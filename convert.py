#   Reading Excel
import openpyxl
from helpers import get_values as get_values, find_start as find_start, add_to_calendar as add_to_calendar

# Test data
# from test import allData as allData

# get_values('Invoice 20170710 - OT.xlsx')
# print find_start('Invoice 20170710 - OT.xlsx')
keys = allData.keys()
test_datum = allData[keys[0]]
testing = {keys[0]: test_datum}

add_to_calendar(testing)

#   Writing to Excel
# write_wb = openpyxl.load_workbook('write.xlsx')
# write_sheet = write_wb.get_sheet_by_name("May")
# write_sheet['A1'] = "Hello World"
#
# write_wb.save('write.xlsx')
