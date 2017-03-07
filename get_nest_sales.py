import xlrd
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException

def get_nest_sales(filename):
    # first open the file using xlrd    book = xlrd.open_workbook(filename)
    book = xlrd.open_workbook(filename)
    sheet = book.sheet_by_index(0)
    nrows, ncols = sheet.nrows, sheet.ncols

    values = [] #Will hold the list of agent IDs with NEST sales

    #Collect up the agent IDs with Nest sales add them to the value array and return the array.
    #return format is [agent_id]
    for row in range(1, nrows):
        #The agent ID is on column 16, the product name on column 5 and the bounce status on column 10
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

            values.append(sheet.cell_value(row,16))

    return values
