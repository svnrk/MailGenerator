import pandas as pd

class Email:

    def __init__(self, name="", address="", mail_subject="", content="", attachment=""):
        self.name = name
        self.address = address
        self.mail_subject = mail_subject
        self.content = content
        self.attachment = attachment

    def __str__(self):
        return ("Address: %s, Topic: %s" % (self.address, self.mail_subject))



excel_path = "/Users/sven-erik/PycharmProjects/MailGenerator/file_example_XLS_50.xls"
df = pd.read_excel(excel_path)
mail_addresses = df["Aadress"]
mail_names = df["Nimi"]
mail_attachment_name = df["Leping"]

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os.path


def send_email(email_recipient,
               email_subject,
               email_message,
               attachment_location = ''):

    email_sender = 'your_email_address@your_server.com' # my stuff

    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = email_recipient
    msg['Subject'] = email_subject

    msg.attach(MIMEText(email_message, 'plain'))

    if attachment_location != '':
        filename = os.path.basename(attachment_location)
        attachment = open(attachment_location, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        "attachment; filename= %s" % filename)
        msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.office365.com', 587) # my stuff
        server.ehlo()
        server.starttls()
        server.login('your_login_name', 'your_login_password') # my stuff
        text = msg.as_string()
        server.sendmail(email_sender, email_recipient, text)
        print('email sent to ' + email_recipient)
        server.quit()
    except:
        print("SMPT server connection error")

    return True

# send_email('bgates@microsoft.com', 'Happy New Year', 'We love Outlook', 'C:\Postcard\NYE.gif')


excel_path = "/Users/sven-erik/PycharmProjects/MailGenerator/file_example_XLS_50.xls"
excel_keyword = "Country"

# last_names = mailAddressesFromExcel(excel_path, excel_keyword)

old = Email(address="22")
print (old)
next_email = Email("svenerikmandmaa@gmail.com", "test", "testtesttest", "false")
print (next_email.address)
print (last_names[5])