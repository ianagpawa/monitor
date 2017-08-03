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


def add_site(wb, sheet, site):
    for i in range(3, 50):
        cell = sheet.cell(row=i, column=1)
        if cell.value:
            if cell.value == site:
                return i
        else:
            cell.value = site
            wb.save('calendar.xlsx')
            return i

def get_initials(name):
    names = name.split(" ")
    return names[0][0] + " " + names[1][0]

def add_to_calendar(obj):
    site = obj.keys()[0]
    dates = obj[site]['date']
    therapist = obj[site]['therapist']
    initials = get_initials(therapist)

    wb = openpyxl.load_workbook("calendar.xlsx")


    for da in dates:
        sheet_names = wb.sheetnames
        month = da.strftime("%B")
        day = int(da.strftime("%d"))

        if month not in sheet_names:
            create_calendar(month)

        current_sheet = wb[month]
        current_row = add_site(wb, current_sheet, site)
        cell = current_sheet.cell(row=current_row, column=day+1)
        if cell.value:
            cell.value += " and " + initials
        else:
            cell.value = initials
        wb.save('calendar.xlsx')
