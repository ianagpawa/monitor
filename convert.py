#   Reading Excel
import openpyxl
from helpers import get_values as get_values, find_start as find_start, add_to_calendar as add_to_calendar, get_invoice_number as get_invoice_number, add_invoice as add_invoice, create_calendar as create_calendar, check_name as check_name

import os
# import pkg_resources
# print pkg_resources.get_distribution("openpyxl").version

# def execute():
#     current = os.getcwd()
#     # Windows file system \
#     invoices_folder = current + "\Invoices"
#     os.chdir(invoices_folder)
#     file_names = os.listdir(os.getcwd())
#     for f in file_names:
#         if f != "Done" and f != "Not Done":
#             f_path = f
#             add_invoice(f_path)
#
# execute()



add_invoice('Invoice 20170710 - OT.xlsx')
# add_invoice('Invoice 20170428 - OT.xlsx')
