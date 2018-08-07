from SimpleExcelFile.EmailDetails import *
import datetime


def send_email(fromaddr, toaddr, passwd, subject, body, filename):
    email = EmailDetails(fromaddr, toaddr, passwd, subject, body, filename)
    email.send_email()


year_init = 2018
month_init = 7
day_init = 23
year_end = 2018
month_end = 7
day_end = 27
filename = "reporte_horas.xlsx"
passwd = '[password]'
subject = 'Reporte de horas de la semana'
fromaddr = '[email]'
toaddr = ['mail1', 'mail2', 'mailn']
message = 'Hola un cordial saludo, soy [nombre], te adjunto el reporte de horas de la semana, hablando con la gente de Internacional de sistemas me dijeron que no habia problema si les enviaba el reporte desde este correo.'
body = message
subject += " " + datetime.datetime(year_init, month_init, day_init).strftime('%d/%m/%y') + " a " + datetime.datetime(
    year_end, month_end, day_end).strftime('%d/%m/%y')
send_email(fromaddr, toaddr, passwd, subject, body, filename)
