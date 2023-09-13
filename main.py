import smtplib
import ssl
import time
from email.message import EmailMessage
from email.header import Header
from email.mime.multipart import MIMEMultipart
import json
from email.mime.text import MIMEText
from email.utils import formataddr
from utils.excel_collector import IterXlsx
import base64



context = ssl.create_default_context()
cf = json.load(open('./config.json', 'r'))
sender = cf['sender_name']
tmp = open('./Template', 'r')
tmp = tmp.read()
xlsx = cf['excel']




def build_email(html_message: str, receiver_addr: str, receiver_name: str):
    msg = MIMEMultipart()
    msg['Subject'] = f"{receiver_name} 请查收你的面试通知"
    sender_name = base64.b64encode(cf['sender_name'].encode('utf-8')).decode()
    msg['From'] = f"=?UTF-8?B?{sender_name}=?= <{cf['email_addr']}>"
    receiver_name = base64.b64encode(receiver_name.encode('utf-8')).decode()
    msg['To'] = f"=?UTF-8?B?{receiver_name}=?= <{receiver_addr}>"
    msg.attach(MIMEText(html_message, 'html', 'utf-8'))
    return msg


def main():
    smtp_url = cf['smtp_server']['server']
    smtp_port = cf['smtp_server']['port']
    with smtplib.SMTP_SSL(smtp_url, smtp_port ,context=context) as smtp_cli:
        email_addr = cf['email_addr']
        smtp_key = cf['smtp_key']
        smtp_cli.login(user=email_addr, password=smtp_key)
        for member in IterXlsx(xlsx):
            msg = build_email(tmp.replace('{{.Name}}', member.name), member.mail, member.name)
            smtp_cli.sendmail(from_addr=email_addr, to_addrs=member.mail, msg=msg.as_string())
            print(f"Send: {msg}")
            time.sleep(0.5)

        # msg = build_email(tmp.replace("{$msg}", "hello"), "540122658@qq.com", "mr_yang")
        # smtp_cli.sendmail(from_addr=email_addr, to_addrs="540122658@qq.com", msg=msg.as_string())



if __name__ == '__main__':
        main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
