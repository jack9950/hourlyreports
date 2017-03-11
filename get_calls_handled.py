import sys
import xlrd
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException

def get_calls_handled(filename):
# first open using xlrd    book = xlrd.open_workbook(filename)
    try:
        book = xlrd.open_workbook(filename)
    except FileNotFoundError:
        print("File: ", filename)
        print("\nFile not found...Exiting...")
        sys.exit()

    sheet = book.sheet_by_index(1)
    nrows, ncols = sheet.nrows, sheet.ncols

    values = []

    for row in range(6, nrows):
        if sheet.cell_value(row,4) != '':
            #The format is [agent ID, Calls Handled, Sales Calls Handled]
            values.append([sheet.cell_value(row,4), sheet.cell_value(row,11), sheet.cell_value(row,47)])

    return values
