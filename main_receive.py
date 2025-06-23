# %%
import imaplib  # 新增：用於收信
import re
from email import message_from_bytes
from email.header import decode_header
from email.utils import parseaddr

#設定 Gmail帳戶與應用程式密碼
gmail_user = 'ownbimonthly@gmail.com' #Gmail帳戶
gmail_password = 'xonv umvk qwsm gcbj' #應用程式密碼

def fetch_emails():
    """連接 Gmail 並拿取每一封郵件的標題與寄件人"""
    imap_server = 'imap.gmail.com'
    mail = imaplib.IMAP4_SSL(imap_server)
    mail.login(gmail_user, gmail_password)
    mail.select('inbox')
    status, data = mail.search(None, 'ALL')
    email_ids = data[0].split()
    emails = []
    for eid in email_ids:
        status, msg_data = mail.fetch(eid, '(RFC822)')
        if status != 'OK':
            continue
        msg = message_from_bytes(msg_data[0][1])
        subject, encoding = decode_header(msg.get('Subject'))[0]
        if isinstance(subject, bytes):
            subject = subject.decode(encoding or 'utf-8', errors='ignore')
        from_addr = parseaddr(msg.get('From'))[1]
        emails.append({'subject': subject, 'from': from_addr})
    mail.logout()
    return emails

def collect_bounced_recipients():
    """收集所有被 Gmail 風控退信的收件人信箱 (僅收集 gmail.com)"""
    bounced_recipients = set()
    mails = fetch_emails()
    for mail in mails:
        # 只處理退信郵件
        if (
            'delivery status notification' in mail['subject'].lower() or
            'undeliverable' in mail['subject'].lower() or
            'mail delivery failed' in mail['subject'].lower()
        ):
            # 解析郵件內容
            imap_server = 'imap.gmail.com'
            with imaplib.IMAP4_SSL(imap_server) as imap:
                imap.login(gmail_user, gmail_password)
                imap.select('inbox')
                status, data = imap.search(None, f'SUBJECT "{mail["subject"]}"')
                if status == 'OK' and data[0]:
                    eid = data[0].split()[-1]
                    status, msg_data = imap.fetch(eid, '(RFC822)')
                    if status == 'OK':
                        msg = message_from_bytes(msg_data[0][1])
                        body = ''
                        if msg.is_multipart():
                            for part in msg.walk():
                                if part.get_content_type() == 'text/plain':
                                    charset = part.get_content_charset() or 'utf-8'
                                    body += part.get_payload(decode=True).decode(charset, errors='ignore')
                        else:
                            charset = msg.get_content_charset() or 'utf-8'
                            body = msg.get_payload(decode=True).decode(charset, errors='ignore')
                        # 從內容中找出所有 gmail.com 的信箱
                        found = re.findall(r'[\w\.-]+@gmail\.com', body)
                        bounced_recipients.update(found)
                imap.logout()
    return list(bounced_recipients)

# mails = fetch_emails()
# %% 

for mail in mails:
    print(f"From: {mail['from']} | Subject: {mail['subject']}")

# 範例：列印所有被退信的 gmail 收件人
if __name__ == "__main__":
    bounced = collect_bounced_recipients()
    print("被 Gmail 風控退信的收件人：")
    for addr in bounced:
        print(addr)