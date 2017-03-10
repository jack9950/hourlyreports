import xlrd
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException
from datetime import date
from teams import agent_ids_to_names

def get_fcp_sales_breakdown(filename):
# first open using xlrd    book = xlrd.open_workbook(filename)
    book = xlrd.open_workbook(filename)
    sheet = book.sheet_by_index(0)
    nrows, ncols = sheet.nrows, sheet.ncols

    values = []

    for row in range(1, nrows):

        date_of_sale = sheet.cell_value(row,7).split("-")
        date_of_sale = date(int(date_of_sale[0]), int(date_of_sale[1]), int(date_of_sale[2]))
        account_number = sheet.cell_value(row,1)
        agent_id = sheet.cell_value(row,6)

        if date_of_sale == date.today() and agent_id != '':

            try:
                agent_name = agent_ids_to_names[agent_id]
                values.append([agent_name, account_number])
            except:
                pass

    return values