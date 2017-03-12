import sys
import xlrd
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException
from datetime import date

def get_fcp_sales(filename):
# first open using xlrd    book = xlrd.open_workbook(filename)
    try:
        book = xlrd.open_workbook(filename)
    except FileNotFoundError:
        print("File: ", filename)
        print("\nFile not found...Exiting...")
        sys.exit()

    sheet = book.sheet_by_index(0)
    nrows, ncols = sheet.nrows, sheet.ncols

    values = []

    for row in range(1, nrows):

        date_of_sale = sheet.cell_value(row,7).split("-")
        date_of_sale = date(int(date_of_sale[0]), int(date_of_sale[1]), int(date_of_sale[2]))

        #Format returned in [agent_id]
        if date_of_sale == date.today():
            agent_id = sheet.cell_value(row,6)
            values.append(agent_id)

    return values
