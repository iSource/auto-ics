import imaplib
import email
import os
import re
import json
from icalendar import Calendar, Event
from datetime import datetime
from openai import OpenAI
from dateutil import parser

# 配置 DeepSeek API (兼容 OpenAI 格式)
client = OpenAI(
    api_key=os.getenv("DEEPSEEK_API_KEY"), 
    base_url="https://api.deepseek.com"
)

def get_email_content():
    try:
        mail = imaplib.IMAP4_SSL("imap.qq.com", 993)
        mail.login(os.getenv("EMAIL_USER"), os.getenv("EMAIL_PASS"))
        mail.select("INBOX")
        
        # 搜索最近 5 封来自 12306 的邮件，确保不漏掉
        status, messages = mail.search(None, '(FROM "12306")')
        if status != "OK" or not messages[0]:
            return None
        
        latest_num = messages[0].split()[-1]
        res, msg_data = mail.fetch(latest_num, "(RFC822)")
        
        for response in msg_data:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])
                # 获取邮件主题
                subject = decode_header(msg["Subject"])[0][0]
                if isinstance(subject, bytes):
                    subject = subject.decode()
                
                # 过滤掉退票、改签等不含订票信息的邮件
                if "订单支付成功" not in subject and "时刻表" not in subject:
                    # 如果你的12306邮件主题不同，可以根据实际情况微调这里的过滤条件
                    pass

                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == "text/plain":
                            return part.get_payload(decode=True).decode()
                else:
                    return msg.get_payload(decode=True).decode()
        return None
    except Exception as e:
        print(f"邮件读取出错: {e}")
        return None

def parse_with_deepseek(text):
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "你是一个行程提取专家。从邮件中提取火车票信息并返回严格的JSON格式。"},
                {"role": "user", "content": f"提取以下信息：车次, 出发站, 到达站, 出发时间(YYYY-MM-DD HH:MM), 座位号, 检票口。邮件内容：\n{text}"}
            ],
            response_format={'type': 'json_object'} # DeepSeek 支持 JSON Mode
        )
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        print(f"DeepSeek 识别失败: {e}")
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

    # 使用车次和出发时间作为唯一标识防止重复
    uid = f"{data['train_no']}-{data['start_time']}".replace(" ", "")
    for component in cal.walk():
        if component.name == "VEVENT" and component.get('uid') == uid:
            print("该行程已在日历中，无需重复添加。")
            return

    event = Event()
    event.add('summary', f"🚆 {data['start_station']} - {data['end_station']} ({data['train_no']})")
    event.add('dtstart', parser.parse(data['start_time']))
    # 假设高铁行程通常在2小时左右，设置一个默认结束时间（可选）
    # event.add('dtend', parser.parse(data['start_time']) + timedelta(hours=2))
    event.add('description', f"座位：{data.get('seat', '未提取')}\n检票口：{data.get('gate', '未提取')}")
    event.add('location', f"{data['start_station']}站")
    event.add('uid', uid)
    
    cal.add_component(event)
    with open(file_path, "wb") as f:
        f.write(cal.to_ical())
    print(f"成功添加行程：{data['train_no']}")

if __name__ == "__main__":
    content = get_email_content()
    if content:
        info = parse_with_deepseek(content)
        if info:
            update_ics(info)