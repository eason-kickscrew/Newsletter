import smtplib
from email import encoders  # 用於附加檔案編碼
from email.mime.base import MIMEBase  # 用於承載附加檔案
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart  # email內容載體
from email.mime.text import MIMEText  # 用於製作文字內文
from email.utils import formataddr

#設定 Gmail帳戶與應用程式密碼
gmail_user = 'ownbimonthly2025@gmail.com' #Gmail帳戶
gmail_password = 'vrgw aerj xmuv rhyq' #應用程式密碼
  
def sending_email(reciver, subject, html_template):
    mail = MIMEMultipart()
    # 設定寄件人名稱
    sender_name = 'own 一個人生活'  # 你想顯示的寄件人名稱
    mail['From'] = formataddr((sender_name, gmail_user))
    to_emails = [reciver]
    mail['To'] = ', '.join(to_emails)
    mail['Subject'] = subject  # 修正：直接使用傳入的 subject
    mail.attach(MIMEText(html_template, 'html'))

    #設定smtp伺服器並寄發信件    
    smtpserver = smtplib.SMTP_SSL("smtp.gmail.com",465)
    smtpserver.ehlo()
    smtpserver.login(gmail_user, gmail_password)
    smtpserver.sendmail(gmail_user, to_emails , mail.as_string())
    smtpserver.quit()


