#We will use this until the HIVE sales are included in the products report from Big Bounce
# We will have to manually download the report from Sonar and then save it as products_sonar_(date).xls
#Remember to save as xls!!!!

import xlrd
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException

def get_hive_sales(filename):
# first open using xlrd    book = xlrd.open_workbook(filename)
    book = xlrd.open_workbook(filename)
    sheet = book.sheet_by_index(0)
    nrows, ncols = sheet.nrows, sheet.ncols

    values = []

    #Hive Lighting Kit

    for row in range(1, nrows):
        if sheet.cell_value(row, 5) == 'Hive Lighting Kit':
            agent_id = (sheet.cell_value(row, 16))
            values.append(agent_id)

    return values
