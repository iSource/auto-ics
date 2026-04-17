import imaplib
import email
from email.header import decode_header
import os
import re
import json
import base64
from icalendar import Calendar, Event
from datetime import datetime
from openai import OpenAI
from dateutil import parser
from bs4 import BeautifulSoup
from dotenv import load_dotenv

# 加载 .env 环境变量
load_dotenv()

# 配置 DeepSeek
client = OpenAI(
    api_key=os.getenv("DEEPSEEK_API_KEY"), 
    base_url="https://api.deepseek.com"
)

CACHE_FILE = "processed_emails.json"

def imap_utf7_decode(text):
    """解码 IMAP Modified UTF-7 编码的文件夹名称"""
    def decode_part(m):
        s = m.group(1).replace(',', '/')
        # 补全 base64 填充
        pad = len(s) % 4
        if pad: s += '=' * (4 - pad)
        try:
            return base64.b64decode(s).decode('utf-16-be')
        except:
            return m.group(0)

    # 替换 &...- 部分，排除 &- (代表 & 本身)
    text = re.sub(r'&([^-]+)-', decode_part, text)
    return text.replace('&-', '&')

def load_cache():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r') as f:
                return set(json.load(f))
        except:
            return set()
    return set()

def save_cache(cache):
    with open(CACHE_FILE, 'w') as f:
        json.dump(list(cache), f)

def decode_str(s):
    if s is None: return ""
    value, encoding = decode_header(s)[0]
    if isinstance(value, bytes):
        try:
            return value.decode(encoding or 'gbk')
        except:
            return value.decode('utf-8', 'ignore')
    return value

def get_new_emails():
    """获取未处理过的邮件内容和 ID"""
    new_emails = []
    processed_cache = load_cache()
    
    try:
        mail = imaplib.IMAP4_SSL("imap.qq.com", 993)
        mail.login(os.getenv("EMAIL_USER"), os.getenv("EMAIL_PASS"))
        
        result, folders = mail.list()
        for folder_info in folders:
            try:
                folder_raw = folder_info.decode()
                # 提取原始文件夹名
                folder_name_raw = folder_raw.split(' "/" ')[-1].strip('"')
                # 解码文件夹名
                folder_name = imap_utf7_decode(folder_name_raw)
                
                if any(x in folder_name.lower() for x in ['drafts', 'deleted', 'junk', 'sent', '草稿', '已删除', '垃圾', '已发送']):
                    continue
                    
                print(f"正在检查文件夹: {folder_name}")
                mail.select(f'"{folder_name_raw}"', readonly=True)
                
                status, messages = mail.search(None, '(FROM "12306")')
                
                if status == "OK" and messages[0]:
                    msg_ids = messages[0].split()
                    for msg_id in reversed(msg_ids[-20:]):
                        res, msg_data = mail.fetch(msg_id, "(RFC822)")
                        for response in msg_data:
                            if isinstance(response, tuple):
                                msg = email.message_from_bytes(response[1])
                                message_id = msg.get('Message-ID')
                                
                                if message_id in processed_cache:
                                    continue

                                subject = decode_str(msg.get('Subject'))
                                print(f"    - 发现新邮件: {subject}")

                                content = ""
                                if msg.is_multipart():
                                    for part in msg.walk():
                                        ctype = part.get_content_type()
                                        if ctype == "text/plain":
                                            content = part.get_payload(decode=True).decode('gbk', 'ignore')
                                            break
                                        elif ctype == "text/html":
                                            html = part.get_payload(decode=True).decode('gbk', 'ignore')
                                            content = BeautifulSoup(html, 'html.parser').get_text(separator='\n')
                                else:
                                    content = msg.get_payload(decode=True).decode('gbk', 'ignore')
                                
                                if content:
                                    new_emails.append({'id': message_id, 'content': content, 'subject': subject})
            except Exception as e:
                # 注意：这里 folder_name 可能是解码后的，folder_name_raw 是原始的
                print(f"跳过文件夹 {folder_name}: {e}")
                continue
        
        return new_emails
    except Exception as e:
        print(f"邮件连接出错: {e}")
        return []

def parse_with_deepseek(text):
    try:
        prompt = """你是一个行程提取专家。从邮件中提取火车票信息并返回 JSON。
重点：
1. 识别操作类型 (action): 'book' (购票/候补成功/改签后的新票), 'cancel' (退票/改签前的旧票/候补失败)。
2. 提取信息：车次, 出发站, 到达站, 出发时间(YYYY-MM-DD HH:MM), 座位号, 检票口。
3. 如果是改签邮件，请提取新票信息并标记 action 为 'book'，同时如果邮件中提到旧票，请识别其信息并返回两条记录。

返回格式示例：
{
  "trips": [
    {"action": "book", "train_no": "G123", "start_station": "北京", "end_station": "上海", "start_time": "2024-01-01 12:00", "seat": "1车1A", "gate": "A1"}
  ]
}"""
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": prompt},
                {"role": "user", "content": f"内容：\n{text}"}
            ],
            response_format={'type': 'json_object'}
        )
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        print(f"AI 识别出错: {e}")
        return None

def update_ics(trips_data):
    file_path = "trips.ics"
    if os.path.exists(file_path):
        with open(file_path, 'rb') as f:
            cal = Calendar.from_ical(f.read())
    else:
        cal = Calendar()
        cal.add('prodid', '-//DeepSeek Trip Bot//')
        cal.add('version', '2.0')

    changed = False
    trips = trips_data.get('trips', [])
    if not trips and 'action' in trips_data:
        trips = [trips_data]

    for data in trips:
        mapped_data = {
            'action': data.get('action', 'book'),
            'train_no': data.get('车次') or data.get('train_no'),
            'start_station': data.get('出发站') or data.get('start_station'),
            'end_station': data.get('到达站') or data.get('end_station'),
            'start_time': data.get('出发时间') or data.get('start_time'),
            'seat': data.get('座位号') or data.get('seat'),
            'gate': data.get('检票口') or data.get('gate')
        }
        
        if not all([mapped_data['train_no'], mapped_data['start_station'], mapped_data['start_time']]):
            continue

        uid = f"{mapped_data['train_no']}-{mapped_data['start_time']}".replace(" ", "").replace(":", "")
        
        existing_event = None
        new_subcomponents = []
        for component in cal.subcomponents:
            if component.name == "VEVENT" and str(component.get('uid')) == uid:
                existing_event = component
            else:
                new_subcomponents.append(component)

        if mapped_data['action'] == 'cancel':
            if existing_event:
                cal.subcomponents = new_subcomponents
                print(f"      [删除] 已取消行程：{mapped_data['train_no']} ({mapped_data['start_time']})")
                changed = True
            continue
        
        if existing_event:
            print(f"      [跳过] 行程 {mapped_data['train_no']} 已存在")
            continue

        # 根据车次选择 Emoji
        train_no = mapped_data['train_no'].upper()
        emoji = "🚅" if any(train_no.startswith(x) for x in ['G', 'D', 'C']) else "🚆"
        
        event = Event()
        # 美化标题：Emoji 座位 | 车次 | 出发 ➔ 到达
        seat = mapped_data.get('seat')
        if seat:
            summary = f"{emoji} {seat} | {train_no} | {mapped_data['start_station']} ➔ {mapped_data['end_station']}"
        else:
            summary = f"{emoji} {train_no} | {mapped_data['start_station']} ➔ {mapped_data['end_station']}"
        
        event.add('summary', summary)
        event.add('dtstart', parser.parse(mapped_data['start_time']))
        event.add('description', f"座位：{mapped_data.get('seat', '')}\n检票口：{mapped_data.get('gate', '')}")
        event.add('location', f"{mapped_data['start_station']}站")
        event.add('uid', uid)
        
        cal.add_component(event)
        print(f"      [写入] 已成功添加行程：{mapped_data['train_no']} ({mapped_data['start_time']})")
        changed = True

    if changed:
        with open(file_path, "wb") as f:
            f.write(cal.to_ical())
    return True

if __name__ == "__main__":
    if not os.path.exists("trips.ics"):
        with open("trips.ics", "w") as f:
            f.write("BEGIN:VCALENDAR\nVERSION:2.0\nEND:VCALENDAR")

    processed_cache = load_cache()
    new_emails = get_new_emails()

    if new_emails:
        print(f"\n开始解析 {len(new_emails)} 封新邮件...")
        for mail_item in new_emails:
            print(f"  -> 正在解析: {mail_item['subject']}")
            info = parse_with_deepseek(mail_item['content'])
            if info:
                update_ics(info)
                processed_cache.add(mail_item['id'])
        
        save_cache(processed_cache)
        print(f"\n处理完成！已记录 {len(new_emails)} 封邮件 ID。")
    else:
        print("所有邮件均已处理过，无需调用 API。")
