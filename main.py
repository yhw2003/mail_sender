import smtplib
import ssl
import time
from email.message import EmailMessage
import json
from utils.excel_collector import IterXlsx


context = ssl.create_default_context()
cf = json.load(open('./config.json', 'r'))
sender = cf['sender_name']
tmp = open('Templet', 'r')
tmp = tmp.read()
xlsx = cf['excel']




def build_email(html_message: str, receiver_addr: str, receiver_name: str):
    msg = EmailMessage()
    msg['From'] = f"{cf['sender_name']} {cf['email_addr']}"
    msg['To'] = f"{receiver_name} {receiver_addr}"
    msg.set_content(html_message)
    msg['Subject'] = f"{cf['sub']}"
    return msg


def main():
    smtp_url = cf['smtp_server']['server']
    smtp_port = cf['smtp_server']['port']
    with smtplib.SMTP_SSL(smtp_url, smtp_port ,context=context) as smtp_cli:
        email_addr = cf['email_addr']
        smtp_key = cf['smtp_key']
        smtp_cli.login(user=email_addr, password=smtp_key)
        for member in IterXlsx(xlsx):
            msg = build_email(tmp.replace('{$name}', member.name), member.mail, member.name)
            smtp_cli.sendmail(from_addr=email_addr, to_addrs=member.mail, msg=msg.as_string())
            print(f"Send: {msg}")
            time.sleep(0.5)

        # msg = build_email(tmp.replace("{$msg}", "hello"), "540122658@qq.com", "mr_yang")
        # smtp_cli.sendmail(from_addr=email_addr, to_addrs="540122658@qq.com", msg=msg.as_string())



if __name__ == '__main__':
        main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
