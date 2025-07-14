# %%
import imaplib  # 新增：用於收信
import re
from email import message_from_bytes
from email.header import decode_header
from email.utils import parseaddr

#設定 Gmail帳戶與應用程式密碼
gmail_user = 'ownbimonthly@gmail.com' #Gmail帳戶
gmail_password = 'xonv umvk qwsm gcbj' #應用程式密碼
# gmail_user = 'eason88091372104@gmail.com' #Gmail帳戶
# gmail_password = 'tbqa dzgg hlrw aqbo' #應用程式密碼

def fetch_emails():
    """連接 Gmail 並拿取每一封郵件的標題與寄件人"""
    imap_server = 'imap.gmail.com'
    mail = imaplib.IMAP4_SSL(imap_server)
    mail.login(gmail_user, gmail_password)
    mail.select('inbox')
    status, data = mail.search(None, 'ALL')
    email_ids = data[0].split()
    emails = []
    for eid in email_ids[::-1]:
        status, msg_data = mail.fetch(eid, '(RFC822)')
        if status != 'OK':
            continue
        msg = message_from_bytes(msg_data[0][1])
        from_addr = parseaddr(msg.get('From'))[1]
        print(f"From: {from_addr}")
        subject, encoding = decode_header(msg.get('Subject'))[0]
        if isinstance(subject, bytes):
            subject = subject.decode(encoding or 'utf-8', errors='ignore')
        print(f"{subject}")
        # 如果有 In-Reply-To 或 References，印出主旨
        if msg.get('In-Reply-To') or msg.get('References'):
            print("回信主旨:", subject)
        body = ''
        if subject=='Delivery Status Notification (Failure)' and msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == 'text/plain':
                    charset = part.get_content_charset() or 'utf-8'
                    body += part.get_payload(decode=True).decode(charset, errors='ignore')
                    none_mail = body.split(' ')[3]
                    emails.append(none_mail)
    mail.logout()
    return emails
mails = fetch_emails()
# %% 

for mail in mails:
    print(f"From: {mail['from']} | Subject: {mail['subject']}")

# 範例：列印所有被退信的 gmail 收件人
if __name__ == "__main__":
    bounced = collect_bounced_recipients()
    print("被 Gmail 風控退信的收件人：")
    for addr in bounced:
        print(addr)