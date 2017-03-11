import xlrd
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException

def get_HIVE_renewals(filename):
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

    #Collect up the warranty sales add them to the value arra and return the array.
    for row in range(1, nrows):
            # agent_id_cell = Column 16 (Column Q)
            # product_name_cell = Column 5 (Column F)
            # bounce_status_cell = Column 10 (Column K)
        agent_id = sheet.cell_value(row, 19)
        #print(agent_id)
        product_name = sheet.cell_value(row, 11)
        #print(product_name)
        bounce_status = sheet.cell_value(row, 3)
        #print(bounce_status)

        if (agent_id != None and
          (product_name == "Home Hero 24" or
           product_name == "Home Hero 24 - ONC" or
           product_name == "Home Hero 24 - CNP" or
           product_name == "Home Hero 24 - AEPC" or
           product_name == "Home Hero 24 - AEPN" or
           product_name == "Home Hero 24 - TNMP") and
          (bounce_status == "Accepted" or
	       bounce_status == "Scheduled" or
	       bounce_status == "No deposit due" or
	       bounce_status == "Ercot/ISO Processing" or
	       bounce_status == "Deposit due in first bill" or
	       bounce_status == "Deposit paid" or
	       bounce_status == "Deposit waiver accepted")):
            #The format is [agent ID]
            values.append(agent_id)
            # print (sheet.cell_value(row,17), sheet.cell_value(row,6), sheet.cell_value(row,11))

    return values
