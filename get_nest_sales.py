import xlrd
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException

def get_nest_sales(filename):
# first open using xlrd    book = xlrd.open_workbook(filename)
    book = xlrd.open_workbook(filename)
    sheet = book.sheet_by_index(0)
    nrows, ncols = sheet.nrows, sheet.ncols
    # print("nrows: ", nrows, "ncols: ", ncols)

    values = []

    for row in range(1, nrows):
        # print(sheet.cell_value(row, 17) != None)
        # print(sheet.cell_value(row, 10))
        agent_id = sheet.cell_value(row, 16)
        product_name = sheet.cell_value(row, 5)
        bounce_status = sheet.cell_value(row, 10)
        if (agent_id != "" and
            product_name == "Nest TX" and
    	   (bounce_status == "Accepted" or
    	    bounce_status == "Scheduled" or
    	    bounce_status == "No deposit due" or
    	    bounce_status == "Ercot/ISO Processing" or
    	    bounce_status == "Deposit due in first bill" or
    	    bounce_status == "Deposit paid" or
    	    bounce_status == "Deposit waiver accepted")):
            #The format is [agent ID, Product Name, Bounce Status]
            values.append(sheet.cell_value(row,16))
            # print (sheet.cell_value(row,17), sheet.cell_value(row,6), sheet.cell_value(row,11))

    return values
