#!/usr/bin/env python
# _*_coding: utf-8_*_

from ExcelFile import *
from EmailDetails import *


def create_file(filename):
    file = ExcelFile()
    file.set_datetime(datetime.datetime(2018, 6, 28))
    file.set_headers()
    file.set_data()
    file.save_as(filename)
    return file.get_week()


def send_email(fromaddr, toaddr, passwd, subject, body, filename):
    email = EmailDetails(fromaddr, toaddr, passwd, subject, body, filename)
    email.send_email()


filename = 'nomina.xlsx'
passwd = ''
subject = 'Reporte de la semana'
fromaddr = ''
toaddr = ['mail1', 'mail2', 'mailn']
body = 'Hola un cordial saludo, soy [nombre], te adjunto el reporte de horas de la semana.'
init_week, end_week = create_file(filename)
subject += " " + init_week.strftime('%d/%m/%y') + " a " + end_week.strftime('%d/%m/%y')
send_email(fromaddr, toaddr, passwd, subject, body, filename)
