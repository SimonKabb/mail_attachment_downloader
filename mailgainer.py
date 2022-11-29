import email
import imaplib
import os
from email.header import decode_header

import pandas
from dotenv import load_dotenv

load_dotenv()
LOGIN = os.getenv('LOGIN')
PASSWORD = os.getenv('PASSWORD')
SMTP_SERV = os.getenv('SMTP_SERV')
PORT = os.getenv('PORT')
def mailgainer():
    imap = imaplib.IMAP4_SSL(SMTP_SERV)
    imap.login(LOGIN, PASSWORD)
    imap.select()
    list_unseen = imap.uid('search', "UNSEEN", "ALL")[-1][0].split()
    if list_unseen ==[]:
        print('no unread messages')
    subject = []
    letter_from = []
    attachement = []
    for i in list_unseen:
        status, data = imap.uid('fetch',i, '(RFC822)')
        msg = email.message_from_bytes(data[0][1])
        try:
            subject.append(decode_header(msg["Subject"])[0][0].decode())
        except AttributeError:
            subject.append(decode_header(msg["Subject"])[0][0])
        except UnicodeDecodeError:
            subject.append(decode_header(msg["Subject"])[0][0])
        letter_from.append(msg['Return-path'])
        for part in msg.walk():
            if part.get_content_disposition() == 'attachment':
                # data = base64.b64decode(img)
                data = part.get_payload(decode=True)
                out = open('img.jpg', 'wb')
                out.write(data)
                out.close
                attachement.append(part)



    df = pandas.DataFrame({'subject': subject, 'letter_from': letter_from, 'attachement': attachement})
    df.to_excel('./mails.xlsx')
    imap.logout()
mailgainer()