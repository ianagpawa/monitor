import openpyxl
from openpyxl.styles import PatternFill, Font
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

def check_name(document_name):
    if 'SP' in document_name:
        return "F"
    elif 'OT' or 'Psy' in document_name:
        return "G"
    else:
        return "ERROR - NO MODALITY IN NAME"

# Get values from invoice
def get_values(document_name):
    data = {}
    therapist_c = check_name(document_name)

    wb = openpyxl.load_workbook(document_name)
    sheet = wb.get_sheet_by_name('Sheet1')
    start, last = find_start(document_name)

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

        new_month_sheet.column_dimensions[column].width = "10"
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
            row.height = "50"

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

therapist_colors = {
    "SK": "E7C80E",
    "IM": "E7C80E",
    "BR": "E7C80E",
    "IR": "E7C80E",
    "MJ": "004C00",
    "GK": "004C00",
    "DS": "004C00",
    "DIO": "660000",
    "TK": "660000",
    "TP": "660000",
    "WS": "660000",
    "AV": "660000"
}

def add_to_calendar(site, obj, invoice_num):
    dates = obj['date']
    therapist = obj['therapist']
    initials = get_initials(therapist)

    color = therapist_colors[initials]
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
    invoice_num = get_invoice_number(document_name)
    invoice_data = get_values(document_name)
    sites = invoice_data.keys()
    for site in sites:
        site_obj = invoice_data[site]
        add_to_calendar(site, site_obj, invoice_num)
    print "Finished with %s" % invoice_num
#
# def set_width(month):
#     wb = openpyxl.load_workbook("calendar.xlsx")
#
#     new_month_sheet = wb.create_sheet(title=month)
#     new_month.dimensions.ColumnDimension
#     first_cell = new_month_sheet.cell(row=1, column=1)
#     first_cell.value = month
#
#     print "Creating days"
#     for i in range(2, 33):
#         cell = new_month_sheet.cell(row=2, column=i)
#         cell.value = i - 1
#     wb.save('calendar.xlsx')
