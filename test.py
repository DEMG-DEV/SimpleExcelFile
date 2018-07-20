from EmailDetails import *
from ExcelFile import *


def create_file(filename):
    file = ExcelFile()
    file.set_datetime(datetime.datetime.now())
    file.set_headers()
    file.set_data()
    file.set_styles()
    file.save_as(filename)
    return file.get_week()


filename = 'reporte_horas.xlsx'
init_week, end_week = create_file(filename)
