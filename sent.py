import smtplib
from email.mime.text import MIMEText
import json

# 1. 讀取模板
def read_template(file_path):
    with open(file_path, "r", encoding="utf-8") as file:
        return file.read()

# 2. 替換模板中的變數
def fill_template(template, variables):
    return template.format(**variables)

# 3. 發送郵件
def send_email(smtp_server, port, sender_email, sender_password, recipient_email, subject, content):
    msg = MIMEText(content, "plain", "utf-8")
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = recipient_email

    with smtplib.SMTP_SSL(smtp_server, port) as server:
        server.login(sender_email, sender_password)
        status=server.sendmail(sender_email, recipient_email, msg.as_string())
        if status=={}:
            print(f"郵件已成功發送給 {recipient_email}")
        else:
            print("郵件傳送失敗!")
        

# 主程式
if __name__ == "__main__":
    
    smtp_server = "smtp.gmail.com"
    port = 465
    sender_email = input("請輸入信箱帳號")
    sender_password = input("請輸入信箱應用程式專用密碼")  # 請使用應用程式專用密碼
    subject = "請確認您的平板是否已開機連網"

    # 信件模板檔案路徑
    template_path = "email_template.txt"
    # 填入output json檔
    jsonFile = open('./test.json','r',encoding="utf-8")
    recipient_json = json.load(jsonFile)

    for name in recipient_json:
        # 替換變數
        tabletID = ""
        for id in recipient_json[name]["TabletID"]:
            tabletID += (id+"\n")
        
        variables = {
            "name": name,
            "TabletID": tabletID,
            "sender_name": "小明"
        }

        # 讀取模板並替換變數
        template = read_template(template_path)
        email_content = fill_template(template, variables)

        recipient_email = recipient_json[name]["email"]
    
        send_email(smtp_server, port, sender_email, sender_password, recipient_email, subject, email_content)
