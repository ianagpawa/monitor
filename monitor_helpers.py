import openpyxl
import os

def check_calendar():
    wb = openpyxl.load_workbook('calendar.xlsx')
    print 'Opening calendar'
    for sheet in wb.worksheets:
        pass


def check_month(worksheet):
    pass


check_calendar()
