import xlrd
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException

def get_nest_sales(filename):
# first open using xlrd    book = xlrd.open_workbook(filename)
    book = xlrd.open_workbook(filename)
    sheet = book.sheet_by_index(0)
    nrows, ncols = sheet.nrows, sheet.ncols
    print("nrows: ", nrows, "ncols: ", ncols)

    values = []

    for row in range(1, nrows):
        # print(sheet.cell_value(row, 17) != None)
        # print(sheet.cell_value(row, 10))
        if (sheet.cell_value(row, 16) != "" and
    	    sheet.cell_value(row, 5) == "Nest TX" and
    	   (sheet.cell_value(row, 10) == "Accepted" or
    	    sheet.cell_value(row, 10) == "Scheduled" or
    	    sheet.cell_value(row, 10) == "No deposit due" or
    	    sheet.cell_value(row, 10) == "Ercot/ISO Processing" or
    	    sheet.cell_value(row, 10) == "Deposit due in first bill" or
    	    sheet.cell_value(row, 10) == "Deposit paid" or
    	    sheet.cell_value(row, 10) == "Deposit waiver accepted")):
            #The format is [agent ID, Product Name, Bounce Status]
            values.append(sheet.cell_value(row,16))
            # print (sheet.cell_value(row,17), sheet.cell_value(row,6), sheet.cell_value(row,11))

    return values
