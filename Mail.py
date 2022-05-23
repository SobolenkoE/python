from exchangelib import DELEGATE, Account, Credentials,Configuration,NTLM
import socks
import smtplib
import imaplib
# credentials = Credentials(
#     username='ADMSK\\yvsobol3',  # Or myusername@example.com for O365
#     password='Cj,jk222'

socks.setdefaultproxy(socks.PROXY_TYPE_HTTP, addr='bproxy.ug.mts.ru',port=3131,username='yvsobol3', password='Cj,jk222')
socks.wrapmodule(smtplib)
socks.wrapmodule(imaplib)
smtp = smtplib.SMTP()

smtp.connect("smtp.gmail.com", 587)
#
# ORG_EMAIL   = "@gmail.com"
# FROM_EMAIL  = "sobolenkoy" + ORG_EMAIL
# FROM_PWD    = "2588goog"
# SMTP_SERVER = "imap.gmail.com"
# SMTP_PORT   = 993
# mail = imaplib.IMAP4_SSL(SMTP_SERVER)
# mail.login(FROM_EMAIL,FROM_PWD)
#
#
# imap = imaplib.IMAP4_SSL('imap.yandex.ru')
# imap.login('sob.evg@yandex.ru', '2588yan')
mailServer = smtplib.SMTP("smtp.gmail.com", 587) #465 587
mailServer.ehlo()
mailServer.starttls()
mailServer.ehlo()
mailServer.login('sobolenkoy@gmailcom', '2588goog')

#

# )
# config = Configuration(
#     server='e-mail.mts.ru', #192.168.160.90
#     credentials=Credentials(username='ADMSK\\yvsobol3', password='Cj,jk222'),
#     has_ssl=False
# )
#
# account = Account(
#     primary_smtp_address='yvsobol3@mts.ru',
#     config=config,
#         access_type=DELEGATE,
# )
#
# # Print first 100 inbox messages in reverse order
# for item in account.inbox.all().order_by('-datetime_received')[:100]:
#     print(item.subject, item.body, item.attachments)


url = 'e-mail.mts.ru'
conn = smtplib.SMTP(url,993)
conn.starttls()
user,password = ("ADMSK\\yvsobol3","Cj,jk222")
conn.login(user,password)