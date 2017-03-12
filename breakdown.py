import sys
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from get_pogo_sales_breakdown import get_pogo_sales_breakdown
from get_DEPP_sales_breakdown import get_DEPP_sales_breakdown
from get_HIVE_new_service_breakdown import get_HIVE_new_service_breakdown
from get_HIVE_renewals_breakdown import get_HIVE_renewals_breakdown
from get_fcp_opportunities_breakdown import get_fcp_opportunities_breakdown
from get_fcp_sales_breakdown import get_fcp_sales_breakdown
from teams import agent_ids_to_names
import time
from datetime import datetime
from data_files import homeFolder, callsHandledReportLocation, pogoSalesReportLocation
from data_files import fcpReportLocation, DEPPreportLocation, hiveNewServiceReportLocation, hiveRenewalsReportLocation

if len(sys.argv) == 1: #user did not pass a date argument
    reportDate = ''
elif len(sys.argv) == 2 and len(sys.argv[1]) == 8: #user passed a date argument - must be in format ddmmyyyy
    reportDate = sys.argv[1]
elif len(sys.argv) > 2 or ( len(sys.argv) == 2 and len(sys.argv[1]) != 8 ): #user passed more than one argument
    print("\nInvalid argument(s)...please enter a date in the format: 'ddmmyyyy' \n\n...exiting")
    sys.exit(2)
    #to do - need to write regex to test for invalid characters and invalid dates

firstRow = 4 #first row to start adding agent sales is row 4
left_alignment = Alignment(horizontal='left')

currentDate = datetime.now().strftime("%A %m-%d-%Y")

#Open the template
template = openpyxl.load_workbook(homeFolder + 'template_breakdown.xlsx')
template_sheets = template.get_sheet_names()
template_first_sheet = template.get_sheet_by_name(template_sheets[0])
template_second_sheet = template.get_sheet_by_name(template_sheets[1])
template_third_sheet = template.get_sheet_by_name(template_sheets[2])

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#Start off with the Bounce POGO Sales:
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------

template_first_sheet["A2"] = currentDate #show the date at the top of the sheet

#The list format that will be returned by get_pogo_sales_breakdown is:
#  [ [agent_name, [Acct #, Order #, order status], [Acct #, Order #, order status] ],
#    [agent_name, [Acct #, Order #, order status], [Acct #, Order #, order status] ] ]
bounce_sales = get_pogo_sales_breakdown(pogoSalesReportLocation(reportDate))
bounce_sales.sort() #Sort alphabetically by agent name

row = firstRow #first row to start adding agent sales is row 4
for bounce_sale in bounce_sales:
    template_first_sheet["A" + str(row)].value = bounce_sale[0]
    template_first_sheet["B" + str(row)].value = bounce_sale[1]
    template_first_sheet["C" + str(row)].value = bounce_sale[2]
    template_first_sheet["D" + str(row)].value = bounce_sale[3]
    template_first_sheet["A" + str(row)].alignment = left_alignment
    template_first_sheet["B" + str(row)].alignment = left_alignment
    template_first_sheet["C" + str(row)].alignment = left_alignment
    template_first_sheet["D" + str(row)].alignment = left_alignment
    row += 1

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#Next the DEPP Sales
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------

template_first_sheet["F2"] = currentDate #show the date at the top of the sheet

DEPP_sales = get_DEPP_sales_breakdown(DEPPreportLocation(reportDate))
DEPP_sales.sort()

row = firstRow #first row to start adding agent sales is row 4
for DEPP_sale in DEPP_sales:
    template_first_sheet["F" + str(row)].value = DEPP_sale[0]
    template_first_sheet["G" + str(row)].value = DEPP_sale[1]
    template_first_sheet["H" + str(row)].value = DEPP_sale[2]
    template_first_sheet["I" + str(row)].value = DEPP_sale[3]
    template_first_sheet["J" + str(row)].value = DEPP_sale[4]
    template_first_sheet["F" + str(row)].alignment = left_alignment
    template_first_sheet["G" + str(row)].alignment = left_alignment
    template_first_sheet["H" + str(row)].alignment = left_alignment
    template_first_sheet["I" + str(row)].alignment = left_alignment
    template_first_sheet["J" + str(row)].alignment = left_alignment
    row += 1

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#Next HIVE Sales
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------

template_second_sheet["A2"] = currentDate #show the date at the top of the sheet

HIVE_new_service_sales = get_HIVE_new_service_breakdown(hiveNewServiceReportLocation(reportDate))
HIVE_new_service_sales.sort()

HIVE_renewal_sales = get_HIVE_renewals_breakdown(hiveRenewalsReportLocation(reportDate))
HIVE_renewal_sales.sort()

all_HIVE_sales = []

for sale in HIVE_new_service_sales:
    if sale not in all_HIVE_sales:
        all_HIVE_sales.append(sale)

for sale in HIVE_renewal_sales:
    if sale not in all_HIVE_sales:
        all_HIVE_sales.append(sale)

all_HIVE_sales.sort()

row = firstRow #first row to start adding agent sales is row 4
for HIVE_sale in all_HIVE_sales:
    template_second_sheet["A" + str(row)].value = HIVE_sale[0]
    template_second_sheet["B" + str(row)].value = HIVE_sale[1]
    template_second_sheet["C" + str(row)].value = HIVE_sale[2]
    template_second_sheet["D" + str(row)].value = HIVE_sale[3]
    template_second_sheet["E" + str(row)].value = HIVE_sale[4]
    template_second_sheet["A" + str(row)].alignment = left_alignment
    template_second_sheet["B" + str(row)].alignment = left_alignment
    template_second_sheet["C" + str(row)].alignment = left_alignment
    template_second_sheet["D" + str(row)].alignment = left_alignment
    template_second_sheet["E" + str(row)].alignment = left_alignment
    row += 1

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#Next FCP Sales
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------

template_third_sheet["A2"] = currentDate #show the date at the top of the sheet

fcp_sales = get_fcp_sales_breakdown(fcpReportLocation(reportDate))
fcp_sales.sort()

row = firstRow #first row to start adding agent sales is row 4
for fcp_sale in fcp_sales:
    template_third_sheet["A" + str(row)].value = fcp_sale[0]
    template_third_sheet["B" + str(row)].value = fcp_sale[1]
    template_third_sheet["A" + str(row)].alignment = left_alignment
    template_third_sheet["B" + str(row)].alignment = left_alignment
    row += 1

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#Finally FCP Opportunites
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------

template_third_sheet["D2"] = currentDate #show the date at the top of the sheet

fcp_opportunities = get_fcp_opportunities_breakdown(pogoSalesReportLocation(reportDate))
fcp_opportunities.sort()

row = firstRow #first row to start adding agent sales is row 4
for fcp_opportunity in fcp_opportunities:
    template_third_sheet["D" + str(row)].value = fcp_opportunity[0]
    template_third_sheet["E" + str(row)].value = fcp_opportunity[1]
    template_third_sheet["F" + str(row)].value = fcp_opportunity[2]
    template_third_sheet["D" + str(row)].alignment = left_alignment
    template_third_sheet["E" + str(row)].alignment = left_alignment
    template_third_sheet["F" + str(row)].alignment = left_alignment
    row += 1

finalReportName = 'BreakdownReport'
currentDate = datetime.now().strftime("%m%d%Y")
currentTime = time.strftime("%I%M%S%p")

print("Saving report... \n")

if len(sys.argv) == 1: #user did not pass a date argument
    #print('sys.argv[0]: ', sys.argv[0])
    template.save(homeFolder + finalReportName + "_" + currentDate + "_" + currentTime + ".xlsx")
elif len(sys.argv) == 2:
    template.save(homeFolder + '\\' + reportDate + '\\' + finalReportName + "_" + reportDate + "_" + currentTime + ".xlsx")

print("Done...")
