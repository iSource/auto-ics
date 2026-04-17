import imaplib
import email
from email.header import decode_header
import os
import re
import json
from icalendar import Calendar, Event
from datetime import datetime
from openai import OpenAI
from dateutil import parser

# 配置 DeepSeek
client = OpenAI(
    api_key=os.getenv("DEEPSEEK_API_KEY"), 
    base_url="https://api.deepseek.com"
)

def get_email_content():
    try:
        mail = imaplib.IMAP4_SSL("imap.qq.com", 993)
        mail.login(os.getenv("EMAIL_USER"), os.getenv("EMAIL_PASS"))
        
        # 获取所有文件夹列表
        result, folders = mail.list()
        # 倒序查找，因为自定义文件夹通常在后面
        for folder_info in reversed(folders):
            try:
                # 解析文件夹名称
                folder_name = folder_info.decode().split(' "/" ')[-1].strip('"')
                print(f"正在检查文件夹: {folder_name}")
                
                mail.select(f'"{folder_name}"', readonly=True)
                # 搜索 12306 邮件
                status, messages = mail.search(None, '(FROM "12306")')
                
                if status == "OK" and messages[0]:
                    print(f"发现邮件！正在提取...")
                    latest_num = messages[0].split()[-1]
                    res, msg_data = mail.fetch(latest_num, "(RFC822)")
                    
                    for response in msg_data:
                        if isinstance(response, tuple):
                            msg = email.message_from_bytes(response[1])
                            if msg.is_multipart():
                                for part in msg.walk():
                                    if part.get_content_type() == "text/plain":
                                        return part.get_payload(decode=True).decode()
                            else:
                                return msg.get_payload(decode=True).decode()
            except Exception as e:
                print(f"跳过文件夹 {folder_name}: {e}")
                continue
        
        print("未发现 12306 邮件。")
        return None
    except Exception as e:
        print(f"邮件连接出错: {e}")
        return None

def parse_with_deepseek(text):
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "你是一个行程提取专家。从邮件中提取火车票信息并返回JSON。"},
                {"role": "user", "content": f"提取车次, 出发站, 到达站, 出发时间(YYYY-MM-DD HH:MM), 座位号, 检票口。内容：\n{text}"}
            ],
            response_format={'type': 'json_object'}
        )
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        print(f"AI 识别出错: {e}")
        return None

def update_ics(data):
    file_path = "trips.ics"
    if os.path.exists(file_path):
        with open(file_path, 'rb') as f:
            cal = Calendar.from_ical(f.read())
    else:
        cal = Calendar()
        cal.add('prodid', '-//DeepSeek Trip Bot//')
        cal.add('version', '2.0')

    uid = f"{data['train_no']}-{data['start_time']}".replace(" ", "")
    for component in cal.walk():
        if component.name == "VEVENT" and component.get('uid') == uid:
            print("行程已存在。")
            return

    event = Event()
    event.add('summary', f"🚆 {data['start_station']} - {data['end_station']} ({data['train_no']})")
    event.add('dtstart', parser.parse(data['start_time']))
    event.add('description', f"座位：{data.get('seat', '')}\n检票口：{data.get('gate', '')}")
    event.add('location', f"{data['start_station']}站")
    event.add('uid', uid)
    
    cal.add_component(event)
    with open(file_path, "wb") as f:
        f.write(cal.to_ical())
    print(f"已更新日历：{data['train_no']}")

if __name__ == "__main__":
    # 强制初始化文件，防止 Git 报错
    if not os.path.exists("trips.ics"):
        with open("trips.ics", "w") as f:
            f.write("BEGIN:VCALENDAR\nVERSION:2.0\nEND:VCALENDAR")

    content = get_email_content()
    if content:
        info = parse_with_deepseek(content)
        if info:
            update_ics(info)