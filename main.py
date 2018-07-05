#!/usr/bin/env python
# _*_coding: utf-8_*_

from ExcelFile import *

file = ExcelFile()
file.set_headers()
file.set_data()
file.save_as('Sample.xlsx')
