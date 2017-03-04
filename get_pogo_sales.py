import xlrd
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException

def get_pogo_sales(filename):
# first open using xlrd    book = xlrd.open_workbook(filename)
    book = xlrd.open_workbook(filename)
    sheet = book.sheet_by_index(0)
    nrows, ncols = sheet.nrows, sheet.ncols

    values = []

    for row in range(1, nrows):
        if sheet.cell_value(row,6) != '':
            #The format is [agent ID]
            values.append(sheet.cell_value(row,6))

    return values
