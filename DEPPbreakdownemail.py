import sys
import time
from datetime import datetime
import itertools
import win32com.client as win32
from get_pogo_sales_breakdown import get_pogo_sales_breakdown
from get_DEPP_sales_breakdown import get_DEPP_sales_breakdown
from get_fcp_opportunities_breakdown import get_fcp_opportunities_breakdown
from get_fcp_sales_breakdown import get_fcp_sales_breakdown
from data_files import pogoSalesReportLocation
from data_files import fcpReportLocation, DEPPreportLocation
from DEPPbreakdownTableFormat import emailStartHtml, emailEndHtml
from DEPPbreakdownTableFormat import rowOpenTag, rowCloseTag
from DEPPbreakdownTableFormat import salesDEPPTableOpenTag, FCPTableOpenTag
from DEPPbreakdownTableFormat import tableCloseTag, tableGutter, fcpTableGutter
from DEPPbreakdownTableFormat import agentNameOpenTag, agentNameCloseTag
from DEPPbreakdownTableFormat import acctNumOpenTag, acctNumCloseTag
from DEPPbreakdownTableFormat import orderNumOpenTag, orderNumCloseTag
from DEPPbreakdownTableFormat import orderStatusOpenTag, orderStatusCloseTag
from DEPPbreakdownTableFormat import DEPPNameOpenTag, DEPPNameCloseTag
from DEPPbreakdownTableFormat import fcpAgentNameOpenTag, fcpAgentNameCloseTag
from DEPPbreakdownTableFormat import fcpAcctNumOneOpenTag, fcpAcctNumOneCloseTag
from DEPPbreakdownTableFormat import fcpAcctNumTwoOpenTag, fcpAcctNumTwoCloseTag
from DEPPbreakdownTableFormat import fcpOrderStatusOpenTag, fcpOrderStatusCloseTag

arguments = []

for arg in sys.argv:
    arguments.append(arg)
arguments = arguments[1:]

try:
    int(arguments[0])
    reportDate = arguments[0]
except:
    reportDate = ''

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# Bounce and DEPP Sales
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------

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

# remove any duplicates - there is probably a better way to do this!
DUPs_removed = []
for DEPP in DEPP_sales:
    if DEPP not in DUPs_removed:
          DUPs_removed.append(DEPP)
DEPP_sales = DUPs_removed

DEPP_sales.sort()

rowData = itertools.zip_longest(bounce_sales, DEPP_sales, fillvalue=[])

# for row in rowData:
#     print(row, "\n")

html = emailStartHtml + salesDEPPTableOpenTag

for row in rowData:
    # format will be [bounce_sale, DEPP_sales]
    # an empty [] means that it is a partially blank row, and
    # one of the two, bounce_sales or DEPP_sales has more rows than the other
    # we will test for this unevenness by checking the length
    bounceSale = row[0]
    DEPPSale = row[1]
    # print("bounceSale: ", bounceSale)
    
    # if len(bounceSale) > 0:
    #     agentName1 = bounceSale[0]
    #     accountNumber1 = str(int(bounceSale[1]))
    #     orderNumber1 = str(int(bounceSale[2]))
    #     orderStatus1 = bounceSale[3]
    # else:
    #     agentName1 = ''
    #     accountNumber1 = ''
    #     orderNumber1 = ''
    #     orderStatus1 = ''

    if len(DEPPSale) > 0:
        agentName2 = DEPPSale[0]
        accountNumber2 = str(int(DEPPSale[1]))
        orderNumber2 = str(int(DEPPSale[2]))
        DEPPName = DEPPSale[3]
        orderStatus2 = DEPPSale[4]
    else:
        agentName2 = ''
        accountNumber2 = ''
        orderNumber2 = ''
        DEPPName = ''
        orderStatus2 = ''

    html += (rowOpenTag
             # + agentNameOpenTag + agentName1 + agentNameCloseTag
             # + acctNumOpenTag + accountNumber1 + acctNumCloseTag
             # + orderNumOpenTag + orderNumber1 + orderNumCloseTag
             # + orderStatusOpenTag + orderStatus1 + orderStatusCloseTag
             # + tableGutter
             + agentNameOpenTag + agentName2 + agentNameCloseTag
             + acctNumOpenTag + accountNumber2 + acctNumCloseTag
             + orderNumOpenTag + orderNumber2 + orderNumCloseTag
             + DEPPNameOpenTag + DEPPName + DEPPNameCloseTag
             + orderStatusOpenTag + orderStatus2 + orderStatusCloseTag
             + rowCloseTag)

html += tableCloseTag + emailEndHtml


# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# Next FCP Sales and Opportunities
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------

# html += FCPTableOpenTag

# fcp_sales = get_fcp_sales_breakdown(fcpReportLocation(reportDate), reportDate)
# fcp_sales.sort()

# fcp_opportunities = get_fcp_opportunities_breakdown(
#     pogoSalesReportLocation(reportDate))
# fcp_opportunities.sort()

# rowData = itertools.zip_longest(fcp_sales, fcp_opportunities, fillvalue=[])

# for row in rowData:
#     # format will be [fcp_sale, fcp_opportunity]
#     # an empty [] means that it is a partially blank row, and
#     # one of the two, fcp_sale or fcp_opportunity has more rows than the other
#     # we will test for this unevenness by checking the length
#     fcpSale = row[0]
#     # print("bounceSale: ", bounceSale)
#     fcpOpportunity = row[1]
#     if len(fcpSale) > 0:
#         agentName1 = fcpSale[0]
#         accountNumber1 = str(int(fcpSale[1]))
#     else:
#         agentName1 = ''
#         accountNumber1 = ''

#     if len(fcpOpportunity) > 0:
#         agentName2 = fcpOpportunity[0]
#         accountNumber2 = str(int(fcpOpportunity[1]))
#         orderStatus2 = fcpOpportunity[2]
#     else:
#         agentName2 = ''
#         accountNumber2 = ''
#         orderStatus2 = ''

#     html += (rowOpenTag
#              + fcpAgentNameOpenTag + agentName1 + fcpAgentNameCloseTag
#              + fcpAcctNumOneOpenTag + accountNumber1 + fcpAcctNumOneCloseTag
#              + fcpTableGutter
#              + fcpAgentNameOpenTag + agentName2 + fcpAgentNameCloseTag
#              + fcpAcctNumTwoOpenTag + accountNumber2 + fcpAcctNumTwoCloseTag
#              + fcpOrderStatusOpenTag + orderStatus2 + fcpOrderStatusCloseTag
#              + rowCloseTag)

# html += tableCloseTag + emailEndHtml

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
    subject = ('DEPP Breakdown ' + reportDate + ' End of Business')
    additionalEmailList = "; ".join(arguments[1:])

except:
    reportDate = ''
    subject = ('DEPP Breakdown ' + currentDate + ' ' + currentTime)
    additionalEmailList = "; ".join(arguments[0:])

mail.To = additionalEmailList + '; jackson.ndiho@iqor.com'
mail.Subject = subject
mail.HtmlBody = subject + ":" + html
mail.send

print("\nDEPP Breakdown email sent to: " + additionalEmailList
      + "; jackson.ndiho@iqor.com \nat " + currentDate + " " + currentTime 
      + "\n\nDone.......")
