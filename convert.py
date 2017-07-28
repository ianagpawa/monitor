#   Reading Excel
import openpyxl
from helpers import get_values as get_values, find_start as find_start

get_values('Invoice 20170710 - OT.xlsx')
# print find_start('Invoice 20170710 - OT.xlsx')



# wb = openpyxl.load_workbook('Invoice 20170710 - OT.xlsx')
# sheet = wb.get_sheet_by_name('Sheet1')

# Example (not ideal)
# d = ws.cell(row=4, column=2, value=10)
#
# cell_range = ws['A1':'C2']
#
# colC = ws['C']
# col_range = ws['C:D']
# row10 = ws[10]
# row_range = ws[5:10]
#
# for row in ws.iter_rows(min_row=1, max_col=3, max_row=2):
#     for cell in row:
#         print(cell)
# <Cell Sheet1.A1>
# <Cell Sheet1.B1>
# <Cell Sheet1.C1>
# <Cell Sheet1.A2>
# <Cell Sheet1.B2>
# <Cell Sheet1.C2>

# for col in ws.iter_cols(min_row=1, max_col=3, max_row=2):
#      for cell in col:
#          print(cell)


# for i in range(1,100):
#     cell = sheet.cell(row=i, column=1)
#     if cell.value:
#         print cell.value


#   Writing to Excel
# write_wb = openpyxl.load_workbook('write.xlsx')
# write_sheet = write_wb.get_sheet_by_name("May")
# write_sheet['A1'] = "Hello World"
#
# write_wb.save('write.xlsx')
