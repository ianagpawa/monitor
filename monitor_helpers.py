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
    for row in worksheet.iter_rows(min_row=2, max_col=32, max_row=40):
        if row[0].value:
            print row[0].value
    # while worksheet.cell(row=i, column=1).value != None:
    #     #   Add checking method here
    #     print worksheet.cell(row=i, column=1).value
    #     i += 1


def check_site(row):
    site = row




wb = openpyxl.load_workbook('calendar.xlsx')
sheet = wb['June']
check_month(sheet)
