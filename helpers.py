import openpyxl
import pprint
import datetime

# Find start and end rows of invoice
def find_start(document_name):
    wb = openpyxl.load_workbook(document_name)
    sheet = wb.get_sheet_by_name('Sheet1')
    for i in range(1, 100):
        location = 'A%s' % i
        cell = sheet[location]
        value = cell.internal_value
        if value == 1:
            start = i
        if value == None and i > 2:
            last = i - 1
            return start, last

# Get values from invoice
def get_values(document_name):
    data = {}
    wb = openpyxl.load_workbook(document_name)
    sheet = wb.get_sheet_by_name('Sheet1')
    start, last = find_start(document_name)

    for i in range(start, last):
        date = sheet["B" + str(i)].value
        site = sheet["C" + str(i)].value
        therapist = sheet["G" + str(i)].value
        if site in data:
            arr = data[site]["date"]
            if date not in arr:
                arr.append(date)
        else:
            data[site] = {"therapist": therapist, "date": [date]}
    results = open('test.py', 'w')
    results.write('allData = ' + pprint.pformat(data))
    results.close()
    print ("Done.")


# Create new calander sheet
def create_calendar(month):
    wb = openpyxl.load_workbook("calendar.xlsx")
    new_month_sheet = wb.create_sheet(title=month)

    first_cell = new_month_sheet.cell(row=1, column=1)
    first_cell.value = month

    for i in range(2, 33):
        cell = new_month_sheet.cell(row=2, column=i)
        cell.value = i - 1
    wb.save('calendar.xlsx')

# Add info from parsed invoice, single data obj, to calendar
def add_to_calendar(obj):
    site = obj.keys()[0]
    dates = obj[site]['date']
    therapist = obj[site]['therapist']
    # iterate through list
    # get position from date
    #
    # Check if site is already on calender
    # if so, find row
    # if not, add to last row
    #
    # check therapist, and color code
    #

    # fill in cell with color code, and Initials
    # if cell already filled, add only initials

    wb = openpyxl.load_workbook("calendar.xlsx")


    for da in dates:
        sheet_names = wb.sheetnames
        month = da.strftime("%B")
        day = da.strftime("%d")\

        if month not in sheet_names:
            create_calendar(month)
