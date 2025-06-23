import sys
import os
import ctypes
import json
import csv
import sqlite3
import shutil
from datetime import datetime
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from pywxdump import get_wx_info, WX_OFFS, get_wx_db, batch_decrypt, MsgHandler
from pywxdump.db.utils.common_utils import dat2img

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def get_account_info():
    # 获取当前运行的微信账号信息
    print("正在检测微信进程...")
    wx_info_list = get_wx_info(WX_OFFS, is_print=True)

    if not wx_info_list:
        print("\n未检测到正在运行的微信进程，请确保：")
        print("1. 微信电脑版已经打开")
        print("2. 已经登录微信账号")
        print("3. 以管理员权限运行此脚本")
        return None

    # 输出微信信息
    for info in wx_info_list:
        print("\n==== 微信账号信息 ====")
        print("微信进程ID:", info.get("pid"))
        print("微信版本:", info.get("version"))
        print("微信账号:", info.get("account"))
        print("手机号:", info.get("mobile"))
        print("昵称:", info.get("nickname"))
        print("邮箱:", info.get("mail"))
        print("wxid:", info.get("wxid"))
        print("数据库密钥:", info.get("key"))
        print("微信文件夹路径:", info.get("wx_dir"))
        print("-" * 40)
    
    return wx_info_list[0] if wx_info_list else None

def export_data(data, filename, format='json'):
    """导出数据到文件"""
    if not data:
        print(f"没有数据可导出到 {filename}")
        return False
        
    os.makedirs('exports', exist_ok=True)
    filepath = os.path.join('exports', filename)
    
    try:
        if format == 'json':
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        elif format == 'csv':
            if not data:
                return False
                
            # 根据数据类型选择字段
            if 'wxid' in data[0]:  # 联系人数据
                fieldnames = [
                    'wxid',         # 用户ID
                    'wx_account',   # 微信号
                    'remark',       # 备注名
                    'nickname',     # 昵称
                    'type',         # 联系人类型
                    'labels'        # 标签列表
                ]
            else:  # 聊天记录数据
                fieldnames = [
                    'id',                 # 本地ID
                    'CreateTime',         # 时间
                    'room_name',          # 聊天室名称
                    'talker',             # 发送者
                    'msg',                # 消息内容
                    'type_name',          # 消息类型
                    'is_sender',          # 是否为发送者
                    'MsgSvrID',           # 消息ID
                    'extra',              # 额外信息
                    'src',                # 原始文件路径
                    'decrypted_image'     # 解密后的图片路径
                ]
            
            with open(filepath, 'w', encoding='utf-8', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
                writer.writeheader()
                for item in data:
                    # 确保所有字段都有值，没有的用空字符串代替
                    row = {}
                    for field in fieldnames:
                        row[field] = item.get(field, '')
                    writer.writerow(row)
            
            print(f"已导出 {len(data)} 条记录到 {filepath}")
        return True
    except Exception as e:
        print(f"导出数据失败: {str(e)}")
        return False

class ContactHandler:
    def __init__(self, db_path):
        self.db_path = db_path
        self.conn = sqlite3.connect(db_path)
        self.cursor = self.conn.cursor()
    
    def close(self):
        if self.cursor:
            self.cursor.close()
        if self.conn:
            self.conn.close()
    
    def convert_value(self, value):
        """转换值为可序列化的格式"""
        if isinstance(value, bytes):
            try:
                # 尝试解码为UTF-8字符串
                return value.decode('utf-8')
            except UnicodeDecodeError:
                # 如果解码失败，转换为十六进制字符串
                return value.hex()
        return value
    
    def get_all_contacts(self):
        """获取所有联系人信息"""
        try:
            # 只选择有用的字段
            query = """
            SELECT 
                UserName,    -- 用户ID(wxid)
                Alias,      -- 微信号
                Remark,     -- 备注名
                NickName,   -- 昵称
                Type,       -- 联系人类型
                LabelIDList -- 标签列表
            FROM Contact
            """
            self.cursor.execute(query)
            rows = self.cursor.fetchall()
            
            # 将结果转换为字典列表
            contacts = []
            for row in rows:
                contact = {
                    'wxid': self.convert_value(row[0]),
                    'wx_account': self.convert_value(row[1]) if row[1] else '',
                    'remark': self.convert_value(row[2]) if row[2] else '',
                    'nickname': self.convert_value(row[3]) if row[3] else '',
                    'type': '普通好友' if row[4] == 0 else '其他来源',
                    'labels': self.convert_value(row[5]) if row[5] else ''
                }
                contacts.append(contact)  # 移除了字段过滤，保留所有字段
            
            return contacts
        except sqlite3.Error as e:
            print(f"读取联系人信息出错: {str(e)}")
            return []

def process_media_file(wx_dir, relative_path, save_dir):
    """处理媒体文件（图片、视频等）
    Args:
        wx_dir: 微信主目录路径
        relative_path: 数据库中存储的相对路径
        save_dir: 保存目录
    Returns:
        bool: 是否成功处理
    """
    try:
        # 构建完整的源文件路径
        # 移除开头的 FileStorage\ 如果存在
        if relative_path.startswith('FileStorage\\'):
            relative_path = relative_path[12:]
        
        # 获取WeChat Files目录（wx_dir的父目录）
        wechat_files_dir = os.path.dirname(wx_dir)
        # 源文件完整路径
        source_path = os.path.join(wechat_files_dir, 'FileStorage', relative_path)
        
        if not os.path.exists(source_path):
            print(f"文件不存在: {source_path}")
            return False
            
        # 创建目标目录
        # 保持原始的目录结构
        relative_dir = os.path.dirname(relative_path)
        target_dir = os.path.join(save_dir, relative_dir)
        os.makedirs(target_dir, exist_ok=True)
        
        # 目标文件路径
        target_path = os.path.join(save_dir, relative_path)
        
        # 复制文件
        shutil.copy2(source_path, target_path)
        print(f"已复制文件: {target_path}")
        return True
    except Exception as e:
        print(f"处理文件出错: {str(e)}")
        return False

def read_chat_messages(wx_info):
    if not wx_info:
        return
    
    wx_dir = wx_info.get("wx_dir")
    wxid = wx_info.get("wxid")
    key = wx_info.get("key")
    
    if not all([wx_dir, wxid, key]):
        print("缺少必要的信息（wx_dir/wxid/key）")
        return

    print("\n==== 开始读取微信数据 ====")
    
    if not os.path.exists(wx_dir):
        print(f"微信文件夹不存在: {wx_dir}")
        return
        
    msg_folder = os.path.join(wx_dir, "MSG")
    if not os.path.exists(msg_folder):
        print(f"MSG文件夹不存在: {msg_folder}")
        return
    
    print("\n尝试获取数据库文件...")
    parent_dir = os.path.dirname(wx_dir)
    print(f"搜索目录: {parent_dir}")
    
    # 获取所需的数据库文件
    db_types = ["MicroMsg", "MSG"]
    db_paths = get_wx_db(msg_dir=parent_dir, db_types=db_types, wxids=wxid)
    
    if not db_paths:
        print("未找到数据库文件")
        return

    print(f"\n找到 {len(db_paths)} 个数据库文件:")
    for db in db_paths:
        print(f"类型: {db['db_type']}, 路径: {db['db_path']}")

    # 创建解密后的数据库保存目录
    decrypt_dir = os.path.join(os.path.dirname(__file__), "decrypted")
    os.makedirs(decrypt_dir, exist_ok=True)
    print(f"\n解密后的数据库将保存在: {decrypt_dir}")
    
    # 创建媒体文件保存目录
    media_dir = os.path.join('exports', 'media')
    os.makedirs(media_dir, exist_ok=True)
    
    # 解密数据库
    print("\n正在解密数据库...")
    db_files = [db['db_path'] for db in db_paths]
    success, result = batch_decrypt(key, db_files, decrypt_dir, is_print=True)
    
    if not success:
        print(f"解密失败: {result}")
        return
    
    # 处理数据库
    for db in db_paths:
        decrypted_db_name = 'de_' + os.path.basename(db['db_path'])
        decrypted_db_path = os.path.join(decrypt_dir, 'Multi' if db['db_type'] == 'MSG' else '.', decrypted_db_name)
        
        if not os.path.exists(decrypted_db_path):
            print(f"解密后的数据库文件不存在: {decrypted_db_path}")
            continue
            
        print(f"\n读取数据库: {decrypted_db_path}")
        try:
            if db['db_type'] == 'MSG':
                db_config = {
                    "key": f"wx_{db['db_type']}",
                    "type": "sqlite",
                    "path": decrypted_db_path
                }
                
                msg_handler = MsgHandler(db_config)
                
                # 获取聊天记录数量
                msg_count = msg_handler.get_m_msg_count()
                if msg_count:
                    print(f"\n总聊天记录数: {msg_count.get('total', 0)}")
                
                # 获取所有聊天记录
                print("\n正在获取聊天记录...")
                all_messages = []
                media_files = []  # 存储媒体文件路径
                page_size = 1000
                page = 0
                
                # 创建解密后图片保存目录
                decrypted_img_dir = os.path.join('exports', 'decrypted_images')
                os.makedirs(decrypted_img_dir, exist_ok=True)
                
                while True:
                    try:
                        messages, wxids = msg_handler.get_msg_list(
                            start_index=int(page*page_size), 
                            page_size=int(page_size)
                        )
                        if not messages:
                            break
                        
                        # 处理消息
                        for msg in messages:
                            try:
                                # 转换时间戳为可读格式
                                if 'CreateTime' in msg and msg['CreateTime']:
                                    create_time = msg['CreateTime']
                                    # 如果已经是格式化的时间字符串，直接使用
                                    if isinstance(create_time, str) and '-' in create_time:
                                        msg['CreateTime'] = create_time
                                    else:
                                        # 如果是时间戳，则转换
                                        if isinstance(create_time, str):
                                            create_time = int(create_time)
                                        msg['CreateTime'] = datetime.fromtimestamp(create_time).strftime('%Y-%m-%d %H:%M:%S')
                                
                                # 处理图片消息
                                if msg.get('type_name') == '图片' and msg.get('src', '').startswith('FileStorage'):
                                    img_path = msg['src']
                                    if img_path.startswith('FileStorage\\'):
                                        img_path = img_path[12:]
                                    
                                    # 构建完整的源文件路径
                                    source_path = os.path.join(parent_dir, wxid, 'FileStorage', img_path)
                                    
                                    if os.path.exists(source_path):
                                        # 解密图片
                                        rc, fmt, md5, out_bytes = dat2img(source_path)
                                        if rc:
                                            # 创建保存解密图片的目录结构
                                            save_dir = os.path.join(decrypted_img_dir, os.path.dirname(img_path))
                                            os.makedirs(save_dir, exist_ok=True)
                                            
                                            # 保存解密后的图片
                                            filename = os.path.basename(source_path)
                                            save_path = os.path.join(save_dir, f"{filename}{fmt}")
                                            
                                            with open(save_path, "wb") as f:
                                                f.write(out_bytes)
                                            
                                            # 添加解密后的图片路径到消息中
                                            msg['decrypted_image'] = os.path.abspath(save_path)
                                            print(f"已解密图片: {msg['decrypted_image']}")
                                        else:
                                            msg['decrypted_image'] = ''
                                            print(f"图片解密失败: {source_path}")
                                    else:
                                        msg['decrypted_image'] = ''
                                        print(f"图片文件不存在（可能未被加载到本地）: {img_path}")
                                else:
                                    msg['decrypted_image'] = ''
                                
                            except (ValueError, TypeError) as e:
                                print(f"消息处理错误: {str(e)}, 原始值: {msg}")
                                continue
                        
                        all_messages.extend(messages)
                        print(f"已获取 {len(all_messages)} 条消息...")
                        page += 1
                    except Exception as e:
                        print(f"获取消息出错: {str(e)}")
                        break
                
                # 导出聊天记录
                if all_messages:
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    if export_data(all_messages, f'chat_messages_{timestamp}.json', 'json'):
                        print(f"\n聊天记录已导出到: exports/chat_messages_{timestamp}.json")
                    if export_data(all_messages, f'chat_messages_{timestamp}.csv', 'csv'):
                        print(f"聊天记录已导出到: exports/chat_messages_{timestamp}.csv")
                
                msg_handler.close()
            
            elif db['db_type'] == 'MicroMsg':
                print("\n正在读取联系人信息...")
                contact_handler = ContactHandler(decrypted_db_path)
                contacts = contact_handler.get_all_contacts()
                contact_handler.close()
                
                if contacts:
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    if export_data(contacts, f'contacts_{timestamp}.json', 'json'):
                        print(f"\n联系人信息已导出到: exports/contacts_{timestamp}.json")
                    if export_data(contacts, f'contacts_{timestamp}.csv', 'csv'):
                        print(f"联系人信息已导出到: exports/contacts_{timestamp}.csv")
                    print(f"\n共导出 {len(contacts)} 个联系人信息")
                else:
                    print("未找到联系人信息")
                
        except Exception as e:
            print(f"读取数据库出错: {str(e)}")

def write_log(log_file, message):
    """写入日志到文件"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f"[{timestamp}] {message}\n")

def test_decrypt_images(wx_info):
    """测试解密微信图片文件"""
    if not wx_info:
        return
    
    # 配置解密限制
    MAX_DECRYPT_IMAGES = 100  # 最大解密图片数量
    current_decrypt_count = 0  # 当前已解密数量
    
    wx_dir = wx_info.get("wx_dir")
    wxid = wx_info.get("wxid")
    if not wx_dir or not wxid:
        print("未找到微信目录或wxid")
        return
        
    # 获取WeChat Files目录（wx_dir的父目录）
    wechat_files_dir = os.path.dirname(wx_dir)
    
    # 创建解密后图片保存目录
    decrypted_img_dir = os.path.join('exports', 'decrypted_images')
    os.makedirs(decrypted_img_dir, exist_ok=True)
    
    # 创建日志文件
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = os.path.join('exports', f'decrypt_images_log_{timestamp}.txt')
    
    write_log(log_file, "=== 微信图片解密日志 ===")
    write_log(log_file, f"开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    write_log(log_file, f"微信ID: {wxid}")
    write_log(log_file, f"解密后图片保存目录: {os.path.abspath(decrypted_img_dir)}")
    write_log(log_file, f"图片解密数量限制: {MAX_DECRYPT_IMAGES}")
    write_log(log_file, "-" * 50)
    write_log(log_file, "\n重要说明：")
    write_log(log_file, "1. 微信采用按需加载的方式处理图片，只有当图片被查看过后才会下载到本地")
    write_log(log_file, "2. 如果看到'图片文件不存在'的提示，通常是因为该图片从未被打开过")
    write_log(log_file, "3. 要获取这些图片，需要在微信中手动打开它们，之后再次运行本程序")
    write_log(log_file, "4. 对于群聊中的图片：")
    write_log(log_file, "   - 自己发送的图片会保存在本地")
    write_log(log_file, "   - 他人发送的图片需要打开查看后才会保存")
    write_log(log_file, f"5. 程序限制最多解密 {MAX_DECRYPT_IMAGES} 张图片")
    write_log(log_file, "-" * 50)
    
    # 遍历导出的聊天记录，查找图片消息
    exports_dir = os.path.join(os.path.dirname(__file__), 'exports')
    if not os.path.exists(exports_dir):
        write_log(log_file, "未找到导出的聊天记录")
        return
        
    # 存储解密结果的字典
    decrypted_images = {}
    
    # 从JSON文件中读取并更新图片消息
    json_files = [f for f in os.listdir(exports_dir) if f.startswith('chat_messages_') and f.endswith('.json')]
    for filename in json_files:
        try:
            json_path = os.path.join(exports_dir, filename)
            with open(json_path, 'r', encoding='utf-8') as f:
                messages = json.load(f)
                
            image_paths = []
            for msg in messages:
                if msg.get('type_name') == '图片' and msg.get('src', '').startswith('FileStorage'):
                    image_info = {
                        'path': msg['src'],
                        'sender': msg.get('talker', '未知'),
                        'room': msg.get('room_name', '私聊'),
                        'is_sender': msg.get('is_sender', 0),
                        'time': msg.get('CreateTime', '未知'),
                        'msg_id': msg.get('id'),  # 用于后续更新消息
                        'json_file': filename  # 记录来源文件
                    }
                    image_paths.append(image_info)
                    
        except Exception as e:
            write_log(log_file, f"读取文件出错 {filename}: {str(e)}")
            continue
    
        if not image_paths:
            continue
            
        total_images = len(image_paths)
        write_log(log_file, f"\n找到 {total_images} 个待解密图片")
        print(f"开始处理 {total_images} 个图片，处理日志将写入: {log_file}")
        
        # 按时间排序
        image_paths.sort(key=lambda x: x['time'])
        success_count = 0
        
        for index, img_info in enumerate(image_paths, 1):
            # 检查是否达到解密限制
            if current_decrypt_count >= MAX_DECRYPT_IMAGES:
                write_log(log_file, f"\n已达到最大解密数量限制 ({MAX_DECRYPT_IMAGES} 张)")
                print(f"\n已达到最大解密数量限制 ({MAX_DECRYPT_IMAGES} 张)")
                break
                
            try:
                write_log(log_file, f"\n[{index}/{total_images}] 处理图片:")
                write_log(log_file, f"时间: {img_info['time']}")
                write_log(log_file, f"发送者: {img_info['sender']}")
                write_log(log_file, f"聊天场景: {'群聊-' + img_info['room'] if '@chatroom' in img_info['room'] else '私聊'}")
                write_log(log_file, f"{'(自己发送)' if img_info['is_sender'] else '(他人发送)'}")
                
                img_path = img_info['path']
                if img_path.startswith('FileStorage\\'):
                    img_path = img_path[12:]
                
                source_path = os.path.join(wechat_files_dir, wxid, 'FileStorage', img_path)
                
                if not os.path.exists(source_path):
                    write_log(log_file, f"图片文件不存在（可能未被加载到本地）: {img_path}")
                    continue
                
                rc, fmt, md5, out_bytes = dat2img(source_path)
                if not rc:
                    write_log(log_file, f"解密失败: {source_path}")
                    continue
                    
                save_dir = os.path.join(decrypted_img_dir, os.path.dirname(img_path))
                os.makedirs(save_dir, exist_ok=True)
                
                filename = os.path.basename(source_path)
                save_path = os.path.join(save_dir, f"{filename}{fmt}")
                
                with open(save_path, "wb") as f:
                    f.write(out_bytes)
                
                success_count += 1
                current_decrypt_count += 1
                write_log(log_file, f"成功解密: {os.path.abspath(save_path)}")
                print(f"\r当前进度: {index}/{total_images}, 成功: {success_count}, 剩余可解密: {MAX_DECRYPT_IMAGES - current_decrypt_count}", end="")
                
                # 记录解密结果
                decrypted_images[img_info['msg_id']] = os.path.abspath(save_path)
                
            except Exception as e:
                write_log(log_file, f"处理图片出错: {str(e)}")
                continue
            
            # 如果达到限制，提前退出循环
            if current_decrypt_count >= MAX_DECRYPT_IMAGES:
                write_log(log_file, f"\n已达到最大解密数量限制 ({MAX_DECRYPT_IMAGES} 张)")
                print(f"\n已达到最大解密数量限制 ({MAX_DECRYPT_IMAGES} 张)")
                break
        
        print()  # 换行
        
        # 更新JSON文件中的消息
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                messages = json.load(f)
            
            # 添加解密后的图片路径
            for msg in messages:
                if msg.get('type_name') == '图片' and msg.get('id') in decrypted_images:
                    msg['decrypted_image'] = decrypted_images[msg['id']]
            
            # 保存更新后的JSON文件
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(messages, f, ensure_ascii=False, indent=2)
            
            write_log(log_file, f"\n已更新聊天记录文件: {filename}")
            
        except Exception as e:
            write_log(log_file, f"更新聊天记录文件失败 {filename}: {str(e)}")
            
        # 如果达到限制，提前退出文件处理
        if current_decrypt_count >= MAX_DECRYPT_IMAGES:
            break
    
    write_log(log_file, f"\n处理完成")
    write_log(log_file, f"总计处理了 {len(json_files)} 个聊天记录文件")
    write_log(log_file, f"解密成功的图片总数: {len(decrypted_images)}")
    write_log(log_file, f"达到限制数量: {'是' if current_decrypt_count >= MAX_DECRYPT_IMAGES else '否'}")
    write_log(log_file, f"\n结束时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    write_log(log_file, "\n提示：")
    write_log(log_file, "1. 如果是别人发送的图片不在本地，这是正常的，因为图片存储在发送者的目录中")
    write_log(log_file, "2. 如果是自己发送的图片不在本地，可能是已经被清理掉了")
    write_log(log_file, "3. 成功解密的图片通常来自最近的聊天记录，较老的图片可能已被清理")
    write_log(log_file, "4. 文件名可能会随时间变化，程序会尝试在相同目录下查找可用的图片文件")
    if current_decrypt_count >= MAX_DECRYPT_IMAGES:
        write_log(log_file, f"5. 由于达到了 {MAX_DECRYPT_IMAGES} 张的解密限制，部分图片未被处理")
    
    print(f"\n处理完成，详细日志已写入: {log_file}")
    print(f"成功解密并更新记录: {len(decrypted_images)} 张图片")
    if current_decrypt_count >= MAX_DECRYPT_IMAGES:
        print(f"注意：已达到 {MAX_DECRYPT_IMAGES} 张的解密限制，部分图片未被处理")

def main():
    # 检查是否以管理员权限运行
    if not is_admin():
        print("请以管理员权限运行此脚本！")
        print("原因：需要管理员权限才能读取微信进程内存")
        return

    # 获取账号信息
    wx_info = get_account_info()
    if wx_info:
        # 读取聊天记录
        read_chat_messages(wx_info)
        # 测试解密图片
        test_decrypt_images(wx_info)
        input("\n按回车键退出...")

if __name__ == "__main__":
    main()