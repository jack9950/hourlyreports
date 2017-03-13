import sys
import xlrd
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException
from datetime import date
from teams import agent_ids_to_names

def get_fcp_sales_breakdown(*args):
# first open using xlrd    book = xlrd.open_workbook(filename)
    try:
        book = xlrd.open_workbook(args[0])
    except FileNotFoundError:
        print("File: ", filename)
        print("\nFile not found...Exiting...")
        sys.exit()

    sheet = book.sheet_by_index(0)
    nrows, ncols = sheet.nrows, sheet.ncols

    values = []

    if args[1] == '':
        saleDate = date.today()
    else:
        # print("args[1]: ", args[1])
        customDate = args[1]
        # print("customDate: ", customDate, type(customDate))
        year = customDate[4:]
        # print("year: ", year)
        year = int(year)
        month = customDate[0:2]
        month = int(month)
        day = customDate[2:4]
        day = int(day)

        saleDate = date(year, month, day)

    for row in range(1, nrows):

        date_of_sale = sheet.cell_value(row,7).split("-")
        date_of_sale = date(int(date_of_sale[0]), int(date_of_sale[1]), int(date_of_sale[2]))
        account_number = sheet.cell_value(row,1)
        agent_id = sheet.cell_value(row,6)

        if date_of_sale == saleDate and agent_id != '':

            try:
                agent_name = agent_ids_to_names[agent_id]
                values.append([agent_name, account_number])
            except:
                pass

    return values
