import smtplib
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os.path as op

from html import HTML
html = HTML.html_content


def automatic_email(html):
    #...Connecting to server
    sendto = 'sheruaravindreddy@applieddatafinance.com'
    user= 'sheruaravindreddy@applieddatafinance.com'
    password = base64.b64decode("############")
    smtpsrv = "smtp.office365.com"
    smtpserver = smtplib.SMTP(smtpsrv,587)
    
    #...Attaching files and content using msg tag
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "Test mail sent by python script"
    msg['From'] = user
    msg['To'] = sendto
    
    html_text = MIMEText(html, 'html')
    msg.attach(html_text)
    
    files = ["C:/Users/sheruaravindreddy/Desktop/Aravind Reddy/Tasks/TU_IDVs.xlsx"]
    for path in files:
        part = MIMEBase('application', "octet-stream")
        with open(path, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename="{}"'.format(op.basename(path)))
        msg.attach(part)
    
    
    smtpserver.ehlo()
    smtpserver.starttls()
    smtpserver.ehlo
    smtpserver.login(user, password)
    
    smtpserver.sendmail(user, sendto, msg.as_string())
    print 'done!'
    smtpserver.close()

automatic_email(html)

# =============================================================================
# year = 2018; month = 8; day = 27; hour = 19; minute = 20
# from datetime import datetime
# start_time = datetime(year, month, day, hour, minute)
# while 1:
#     now = datetime.now()
#     now = now.replace(second=0, microsecond=0)
#     if now == start_time:
#         automatic_email()
#         break
# =============================================================================