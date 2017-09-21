import openpyxl
import os

def check_calendar():
    wb = openpyxl.load_workbook('calendar.xlsx')
    print 'Opening calendar'
    for sheet in wb.worksheets:
        pass


def check_month(worksheet):
    i = 3
    cell = worksheet.cell(row=i, column=1)
    while worksheet.cell(row=i, column=1).value != None:
        #   Add checking method here
        print worksheet.cell(row=i, column=1).value
        i += 1


def check_site(worksheet, row_number):
    site = worksheet.cell(row=row_number, column=1)
    i = 2
    while i <= 32:
        pass




# wb = openpyxl.load_workbook('calendar.xlsx')
# sheet = wb['June']
# check_month(sheet)
