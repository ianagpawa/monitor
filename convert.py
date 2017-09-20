from helpers import add_invoice as add_invoice

import os
# import pkg_resources
# print pkg_resources.get_distribution("openpyxl").version

def execute():
    current = os.getcwd()
    # Windows file system \
    invoices_folder = current + "\Invoices\\"
    # Linux file system
    invoices_folder = current + "/Invoices/"
    file_names = os.listdir(invoices_folder)
    for file_path in file_names:
        if file_path != "Done" and file_path != "Not Done":
            add_invoice(invoices_folder+file_path)

execute()
