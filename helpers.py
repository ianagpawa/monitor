import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
import datetime

# Find start and end rows of invoice
def find_start(document_name):
    wb = openpyxl.load_workbook(document_name)
    sheet = wb.get_sheet_by_name('Sheet1')
    print 'Finding start row of invoices'
    for i in range(1, 125):
        location = 'A%s' % i
        cell = sheet[location]
        value = cell.internal_value
        if value == 1:
            start = i
        if value == None and i > 2:
            last = i - 1
            print "Finished getting start and last rows"
            return start, last

def get_invoice_number(document_name):
    num = ''
    for i in document_name:
        if i.isdigit():
            num += i
    return num[2:]

def check_name(document_name, start):
    wb = openpyxl.load_workbook(document_name)
    sheet = wb.get_sheet_by_name('Sheet1')
    for i in range(1,15):
        cell = sheet.cell(row=start, column=i)
        if 'Therapist' in cell.value:
            return get_column_letter(i)

# Get values from invoice
def get_values(document_name):
    data = {}

    wb = openpyxl.load_workbook(document_name)
    sheet = wb.get_sheet_by_name('Sheet1')
    start, last = find_start(document_name)

    therapist_c = check_name(document_name, start-1)

    for i in range(start, last):
        date = sheet["B" + str(i)].value
        site = sheet["C" + str(i)].value
        therapist = sheet[therapist_c + str(i)].value
        if site in data:
            arr = data[site]["date"]
            if date not in arr:
                arr.append(date)
        else:
            data[site] = {"therapist": therapist, "date": [date]}
    return data
    # results = open('test.py', 'w')
    # results.write('allData = ' + pprint.pformat(data))
    # results.close()
    # print ("Done.")


# Create new calander sheet
def create_calendar(month):
    wb = openpyxl.load_workbook("calendar.xlsx")
    print 'Creating new sheet'
    new_month_sheet = wb.create_sheet(title=month)
    new_month_sheet.page_setup.orientation = new_month_sheet.ORIENTATION_LANDSCAPE


    first_cell = new_month_sheet.cell(row=1, column=1)
    first_cell.value = month

    print "Creating days"
    for i in range(2, 33):
        cell = new_month_sheet.cell(row=2, column=i)
        column = cell.column

        new_month_sheet.column_dimensions[column].width = "15"
        cell.value = i - 1
    wb.save('calendar.xlsx')
    print "Done.  Created new sheet for " + month

# Add info from parsed invoice, single data obj, to calendar
def add_site(wb, sheet, site):
    for i in range(3, 50):
        cell = sheet.cell(row=i, column=1)
        if cell.value:
            if cell.value == site:
                return i
        else:
            row = sheet.row_dimensions[i]
            row.height = "60"

            cell.value = site
            print 'Adding %s to %s' % (site, sheet.title)
            wb.save('calendar.xlsx')
            return i

def get_initials(name):
    names = name.split(" ")
    finished = ''
    for name in names:
        finished += name[0]
    return finished

# therapist_colors = {
#     "SK": "E7C80E",
#     "IM": "E7C80E",
#     "BR": "E7C80E",
#     "IR": "E7C80E",
#     "MJ": "004C00",
#     "GK": "004C00",
#     "DS": "004C00",
#     "DIO": "660000",
#     "TK": "660000",
#     "TP": "660000",
#     "WS": "660000",
#     "AV": "660000"
# }

colors = {
    "OT": "E7C80E",
    "SP": "004C00",
    "Psy": "660000"
}
def add_to_calendar(site, obj, invoice_num, color):
    dates = obj['date']
    therapist = obj['therapist']
    initials = get_initials(therapist)

    color_fill = PatternFill(start_color= color, end_color= color, fill_type="solid")


    wb = openpyxl.load_workbook("calendar.xlsx")

    error = []
    for da in dates:
        sheet_names = wb.sheetnames
        month = da.strftime("%B")
        day = int(da.strftime("%d"))

        if month not in sheet_names:
            create_calendar(month)
            wb = openpyxl.load_workbook("calendar.xlsx")

        current_sheet = wb[month]
        current_row = add_site(wb, current_sheet, site)
        cell = current_sheet.cell(row=current_row, column=day+1)
        if cell.value:
            if initials in cell.value:
                print 'ERROR: %s already submmited for %s %s at %s!' % (therapist, month, day, site)
                error.append((therapist, month, day, site))
                return

            else:
                cell.value += " and " + ("%s-%s" % (invoice_num, initials))
                font = Font(color=color)
                cell.font = font
                cell.font = Font(color=color)

        else:
            cell.value = "%s-%s" % (invoice_num, initials)
            cell.fill = color_fill


        # if error:
        #     return
        wb.save('calendar.xlsx')
        print "Added %s %s" % (da.strftime("%B %d"), initials)


def add_invoice(document_name):
    modality = document_name[-7:-5]
    color = colors[modality]
    invoice_num = get_invoice_number(document_name)
    invoice_data = get_values(document_name)
    sites = invoice_data.keys()
    for site in sites:
        site_obj = invoice_data[site]
        add_to_calendar(site, site_obj, invoice_num, color)
    print "Finished with %s" % invoice_num
