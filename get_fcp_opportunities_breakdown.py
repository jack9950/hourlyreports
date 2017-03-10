import xlrd
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException
from teams import agent_ids_to_names

def get_fcp_opportunities_breakdown(filename):
# first open using xlrd    book = xlrd.open_workbook(filename)
    book = xlrd.open_workbook(filename)
    sheet = book.sheet_by_index(0)
    nrows, ncols = sheet.nrows, sheet.ncols

    values = []

    for row in range(1, nrows):

        bounce_status = sheet.cell_value(row,3)
        agent_id = sheet.cell_value(row,6)
        account_number = sheet.cell_value(row,1)

        if agent_id != '' and bounce_status == 'Deposit due':
            try:
                agent_name = agent_ids_to_names[agent_id]
                values.append([agent_name,
                               account_number,
                               bounce_status])
            except:
                pass

    return values
