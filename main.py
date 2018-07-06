#!/usr/bin/env python
# _*_coding: utf-8_*_

from ExcelFile import *
from EmailDetails import *


def create_file():
    file = ExcelFile()
    file.set_headers()
    file.set_data()
    file.save_as('Sample.xlsx')


def send_email():
    email = EmailDetails()
    email.send_email()
