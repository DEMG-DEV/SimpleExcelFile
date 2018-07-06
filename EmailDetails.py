import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


class EmailDetails():

    def __init__(self, fromaddr, toaddr, passwd, subject, body, filename):
        self.fromaddr = fromaddr
        self.toaddr = toaddr
        self.filename = filename
        self.passwd = passwd
        self.subject = subject
        self.body = body

    def send_email(self):
        msg = MIMEMultipart()

        msg['From'] = self.fromaddr
        msg['To'] = '%s' % ','.join(self.toaddr)
        msg['Subject'] = self.subject

        body = self.body
        msg.attach(MIMEText(body, 'plain'))

        path = os.path.dirname(os.path.abspath(__file__))
        filename = self.filename
        attachment = open(str(path) + '\\' + filename, "rb")

        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        msg.attach(part)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(self.fromaddr, self.passwd)
        text = msg.as_string()
        server.sendmail(self.fromaddr, self.toaddr, text)
        server.quit()
