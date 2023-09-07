import smtplib
import json
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import pandas as pd 
import os


df = pd.read_excel('..\\email list\\mails_clgs.xlsx')
df['emails'].str.strip()
print(df["emails"])

f = open("..\subject_body\email _subject.json", "r")

jsn = json.loads(f.read())
sub = jsn['data']['subject']
print("SUB - ",sub)

f = open("../subject_body/mail_body.html","r")
body = f.read()
print("BODY - ", body)



attachments_path = '../attachments'
attachments = os.listdir(attachments_path)


sender = 'shubhamk531@gmail.com'
psd = 'nfdruhbgexdulvlo'


server = smtplib.SMTP("smtp.gmail.com",587)
server.starttls()
server.login(sender,psd)


def send_mail(email, college):
    print("INSIDE THE FUN")
    receiver = email
    cc = []
    
    
    msg = MIMEMultipart()
    msg['subject']=sub.format(college)
    msg['from']="shubham kumar<shubhamk531@gmail.com>"
    msg['To']=", ".join(receiver)
    msg['cc']=", ".join(cc)
    mail_body=body.format(college)
    msg.attach(MIMEText(mail_body,'html'))
    
    
    for attachment in attachments:
        part=MIMEBase('applications','octet-stream')
        part.set_payload(open(attachments_path + "/" + attachment, 'rb').read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition','attachment',filename=attachment)
        msg.attach(part)
        
    server.sendmail(sender, receiver + cc,msg.as_string())
    print("sent emails to - ",receiver, "college - ",college)
    
    
for index, row in df.iterrows():
    send_mail(row['emails'].split(","),row['college names'])


server.quit()