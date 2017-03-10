import xlrd
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException
from teams import agent_ids_to_names

#Sample return:
# [agent_id, [Acct #, Order #, order status], [Acct #, Order #, order status]]
# [2062062, [2092985, 1443822, "Deposit due"], [2092021, 1444496, "Ercot/ISO Processing"] ]

def get_HIVE_renewals_breakdown(filename):
# first open using xlrd    book = xlrd.open_workbook(filename)
    book = xlrd.open_workbook(filename)
    sheet = book.sheet_by_index(0)
    nrows, ncols = sheet.nrows, sheet.ncols

    values = []

    for row in range(1, nrows):
        agent_id = sheet.cell_value(row, 19)
        plan_name = sheet.cell_value(row, 11)
        if(agent_id != None and
          (plan_name == "Home Hero 24" or
           plan_name == "Home Hero 24 - ONC" or
           plan_name == "Home Hero 24 - CNP" or
           plan_name == "Home Hero 24 - AEPC" or
           plan_name == "Home Hero 24 - AEPN" or
           plan_name == "Home Hero 24 - TNMP")):

            try:
                agent_name = agent_ids_to_names[agent_id]
                pogo_account_number = sheet.cell_value(row,1)
                pogo_order_number = sheet.cell_value(row,0)
                plan_name = sheet.cell_value(row,11)
                bounce_status = sheet.cell_value(row,3)

                values.append([agent_name,
                               pogo_account_number,
                               pogo_order_number,
                               plan_name,
                               bounce_status])
            except:
                pass

    return values
