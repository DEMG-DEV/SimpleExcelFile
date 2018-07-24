#!/usr/bin/env python
# _*_coding: utf-8_*_

from .ExcelFile import *
from .EmailDetails import *
import datetime


def create_file(filename):
    file = ExcelFile()
    file.set_datetime(datetime.datetime.now())
    file.set_headers()
    file.set_data()
    file.set_styles()
    file.save_as(filename)
    return file.get_week()


def send_email(fromaddr, toaddr, passwd, subject, body, filename):
    email = EmailDetails(fromaddr, toaddr, passwd, subject, body, filename)
    email.send_email()


filename = "reporte_horas.xlsx"
passwd = '[password]'
subject = 'Reporte de horas de la semana'
fromaddr = '[email]'
toaddr = ['mail1', 'mail2', 'mailn']
body = 'Hola un cordial saludo, soy [nombre], te adjunto el reporte de horas de la semana, hablando con la gente de Internacional de sistemas me dijeron que no habia problema si les enviaba el reporte desde este correo.'
init_week, end_week = create_file(filename)
subject += " " + init_week.strftime('%d/%m/%y') + " a " + end_week.strftime('%d/%m/%y')
send_email(fromaddr, toaddr, passwd, subject, body, filename)
