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
passwd = 'Guardado22'
subject = 'Nomina de la semana'
fromaddr = 'hack.master.m@gmail.com'
#toaddr = ['admonproyectos@wssgroup.com', 'chernandezt@wssgroup.com', 'adsoto@wssgroup.com']
toaddr = ["demg@outlook.com"]
body = 'Hola un cordial saludo, soy David Enrique Mendez Guardado, te adjunto la nomina de la semana.'
init_week, end_week = create_file(filename)
subject += " " + init_week.strftime('%d/%m/%y') + " a " + end_week.strftime('%d/%m/%y')
send_email(fromaddr, toaddr, passwd, subject, body, filename)
