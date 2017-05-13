import sys
import time
from datetime import datetime, date
import itertools
import win32com.client as win32
import openpyxl
from openpyxl.styles import Alignment
from get_pogo_sales_breakdown import get_pogo_sales_breakdown
from get_DEPP_sales_breakdown import get_DEPP_sales_breakdown
from get_fcp_opportunities_breakdown import get_fcp_opportunities_breakdown
from get_fcp_sales_breakdown import get_fcp_sales_breakdown
from data_files import homeFolder, pogoSalesReportLocation
from data_files import fcpReportLocation, DEPPreportLocation
from breakdownTableFormat import emailStartHtml, emailEndHtml
from breakdownTableFormat import rowOpenTag, rowCloseTag
from breakdownTableFormat import salesDEPPTableOpenTag, FCPTableOpenTag
from breakdownTableFormat import tableCloseTag
from breakdownTableFormat import agentNameOpenTag, agentNameCloseTag

arguments = []

for arg in sys.argv:
    arguments.append(arg)
arguments = arguments[1:]

try:
    int(arguments[0])
    reportDate = arguments[0]
except:
    reportDate = ''

firstRow = 4  # first row to start adding agent sales is row 4
left_alignment = Alignment(horizontal='left')

# Open the template
template = openpyxl.load_workbook(homeFolder + 'template_breakdown.xlsx')
template_sheets = template.get_sheet_names()
template_first_sheet = template.get_sheet_by_name(template_sheets[0])
template_second_sheet = template.get_sheet_by_name(template_sheets[1])

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
#
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------

# template_first_sheet["A2"] = currentDate  # show the date at top of the sheet

# The list format that will be returned by get_pogo_sales_breakdown is:
#  [ [agent_name, [Acct #, Order #, order status],
#                 [Acct #, Order #, order status] ],
#    [agent_name, [Acct #, Order #, order status],
#                 [Acct #, Order #, order status] ] ]
bounce_sales = get_pogo_sales_breakdown(pogoSalesReportLocation(reportDate))
bounce_sales.sort()  # Sort alphabetically by agent name

# The list format that will be returned by get_DEPP_sales_breakdown is:
# [agent_name, pogo_account_number, pogo_order_number,
#  DEPP_name, bounce_status]
DEPP_sales = get_DEPP_sales_breakdown(DEPPreportLocation(reportDate))
DEPP_sales.sort()


rowData = itertools.zip_longest(bounce_sales, DEPP_sales, fillvalue='')

print(rowData)

html = emailStartHtml + salesDEPPTableOpenTag

row = firstRow  # first row to start adding agent sales is row 4
for bounce_sale in bounce_sales:
    agentName = bounce_sale[0]
    accountNumber = bounce_sale[1]
    orderNumber = bounce_sale[2]
    orderStatus = bounce_sale[3]

    html += (rowOpenTag + agentNameOpenTag + agentName + agentNameCloseTag
             + rowCloseTag)

    row += 1

html += tableCloseTag

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# Next the DEPP Sales
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------

# show the date at the top of the sheet
# template_first_sheet["F2"] = currentDate

DEPP_sales = get_DEPP_sales_breakdown(DEPPreportLocation(reportDate))
DEPP_sales.sort()

row = firstRow  # first row to start adding agent sales is row 4
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

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# Next FCP Sales
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------

# show the date at the top of the sheet
# template_second_sheet["A2"] = currentDate

fcp_sales = get_fcp_sales_breakdown(fcpReportLocation(reportDate), reportDate)
fcp_sales.sort()

row = firstRow  # first row to start adding agent sales is row 4
for fcp_sale in fcp_sales:
    template_second_sheet["A" + str(row)].value = fcp_sale[0]
    template_second_sheet["B" + str(row)].value = fcp_sale[1]
    template_second_sheet["A" + str(row)].alignment = left_alignment
    template_second_sheet["B" + str(row)].alignment = left_alignment
    row += 1

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# Finally FCP Opportunites
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------

# show the date at the top of the sheet
# template_second_sheet["D2"] = currentDate

fcp_opportunities = get_fcp_opportunities_breakdown(
    pogoSalesReportLocation(reportDate))
fcp_opportunities.sort()

row = firstRow  # first row to start adding agent sales is row 4
for fcp_opportunity in fcp_opportunities:
    template_second_sheet["D" + str(row)].value = fcp_opportunity[0]
    template_second_sheet["E" + str(row)].value = fcp_opportunity[1]
    template_second_sheet["F" + str(row)].value = fcp_opportunity[2]
    template_second_sheet["D" + str(row)].alignment = left_alignment
    template_second_sheet["E" + str(row)].alignment = left_alignment
    template_second_sheet["F" + str(row)].alignment = left_alignment
    row += 1

# ------------------------------------------------------------------------------
# send email
# ------------------------------------------------------------------------------
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

currentDate = datetime.now().strftime("%m-%d-%y")
currentTime = time.strftime("%#I:%M %p")

try:
    int(arguments[0])
    reportDate = arguments[0]
    reportDate = reportDate[0:2] + '-' + reportDate[2:4] + '-' + reportDate[6:]
    subject = 'DEPP, Sales and FCP Breakdown ' + reportDate + ' End of Business'
    additionalEmailList = "; ".join(arguments[1:])

except:
    reportDate = ''
    subject = 'DEPP, Sales and FCP Breakdown ' + currentDate + ' ' + currentTime
    additionalEmailList = "; ".join(arguments[0:])

mail.To = additionalEmailList + '; jackson.ndiho@iqor.com'
html += emailEndHtml
mail.Subject = subject
mail.HtmlBody = html
mail.send

print("\nEmail sent to: " + additionalEmailList + "; jackson.ndiho@iqor.com.\n\nDone.......")
