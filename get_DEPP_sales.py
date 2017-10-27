import xlrd
import sys
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException

def get_DEPP_sales(filename):
# first open using xlrd    book = xlrd.open_workbook(filename)
    try:
        book = xlrd.open_workbook(filename)
    except FileNotFoundError:
        print("File: ", filename)
        print("\nFile not found...Exiting...")
        raise
    sheet = book.sheet_by_index(0)
    nrows, ncols = sheet.nrows, sheet.ncols

    values = []

    #Collect up the warranty sales add them to the value arra and return the array.
    for row in range(1, nrows):
            # agent_id_cell = Column 16 (Column Q)
            # product_name_cell = Column 5 (Column F)
            # bounce_status_cell = Column 10 (Column K)
        agent_id = sheet.cell_value(row, 16)
        product_name = sheet.cell_value(row, 5)
        if(agent_id != None and (product_name == "Surge Protection Plan" or
                                 product_name == "Electric Repair Essentials" or
                                 product_name == "Surge Protection Plan (20% Off)" or
                                 product_name == "Cooling Maintenance Essentials (6 Month Free Trial - Nest Bundle)" or
                                 product_name == "Cooling Repair & Maintenance Essentials" or
                                 product_name == "Electric Repair Essentials (20% Off)") or
                                 product_name == "Heating & Cooling Repair Essentials"):
            #The format is [agent ID, Product Name, Bounce Status]
            values.append(agent_id)
            # print (sheet.cell_value(row,17), sheet.cell_value(row,6), sheet.cell_value(row,11))
            for value in values:
                print(value)
    return values

