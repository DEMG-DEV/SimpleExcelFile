#!/usr/bin/env python
# _*_coding: utf-8_*_

from openpyxl import Workbook
import datetime
from DateDetails import *


class ExcelFile():

    def __init__(self):
        self.file = Workbook()
        self.sheet = self.file.active
        self.date = datetime.datetime.now()
        self.details = DateDetails()
        self.cells = ["B", "C", "D", "E", "F"]
        self.work_hours = "9.5"

    def save_as(self, filename):
        self.file.save(filename)

    def set_headers(self):
        self.sheet['A1'] = self.date.strftime('%B')
        init_week, end_week = self.details.get_init_end_week(self.date.year, int(self.date.strftime('%V')), 4)
        self.sheet['A2'] = 'Week from ' + init_week.strftime('%d/%m/%y') + ' to ' + end_week.strftime('%d/%m/%y')
        self.sheet['A4'] = 'Activity'
        self.sheet['A5'] = 'Programación'
        for i in self.cells:
            self.sheet[i + '4'] = init_week.strftime('%A') + ' ' + str(init_week.day)
            day = init_week.day + 1
            init_week = date(init_week.year, init_week.month, day)

    def set_data(self):
        for i in self.cells:
            self.sheet[i + '5'] = self.work_hours
