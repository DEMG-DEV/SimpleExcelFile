import smtplib, os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


class EmailDetails():

    def __init__(self):
        pass

    def send_email(self):
        fromaddr = "hack.master.m@gmail.com"
        toaddr = "demg@outlook.com"

        msg = MIMEMultipart()

        msg['From'] = fromaddr
        msg['To'] = toaddr
        msg['Subject'] = "Nomina de la semana"

        body = "Hola un cordial saludo, soy David Enrique Mendez Guardado, te adjunto la nomina de la semana."
        msg.attach(MIMEText(body, 'plain'))

        path = os.path.dirname(os.path.abspath(__file__))
        filename = "Sample.xlsx"
        attachment = open(str(path) + '\\' + filename, "rb")

        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        msg.attach(part)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(fromaddr, "Guardado22")
        text = msg.as_string()
        server.sendmail(fromaddr, toaddr, text)
        server.quit()
