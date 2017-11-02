import sys
import time
from datetime import datetime
import win32com.client as win32

from get_DEPP_sales1 import get_DEPP_sales
from get_DEPP_sales_breakdown import get_DEPP_sales_breakdown

# from get_DEPP_sales2 import get_DEPP_sales, get_DEPP_sales_breakdown

from data_files import DEPPreportLocation
from data_files import tableNames
from data_files import jaelesiaTeam, tekTeam, antwonTeam, jacksonTeam

from DEPPformat import topOfTable
from DEPPformat import agentRowStart, agentRowEnd
from DEPPformat import agentIDStart, agentIDEnd
from DEPPformat import agentNameStart, agentNameEnd
from DEPPformat import DEPPSalesStart, DEPPSalesEnd
from DEPPformat import DEPPSalesStartGreen, DEPPSalesStartNoColor
from DEPPformat import supRowStart, supRowEnd
from DEPPformat import grandTotalRowStart, grandTotalRowEnd
from DEPPformat import supIDStart, supNameStart
from DEPPformat import supDEPPSalesStart
from DEPPformat import gTotalIDStart, gTotalNameStart
from DEPPformat import gTotalDEPPSalesStart

from DEPPbreakdownTableFormat import emailStartHtml, emailEndHtml
from DEPPbreakdownTableFormat import rowOpenTag, rowCloseTag
from DEPPbreakdownTableFormat import salesDEPPTableOpenTag
from DEPPbreakdownTableFormat import tableCloseTag
from DEPPbreakdownTableFormat import agentNameOpenTag, agentNameCloseTag
from DEPPbreakdownTableFormat import acctNumOpenTag, acctNumCloseTag
from DEPPbreakdownTableFormat import orderNumOpenTag, orderNumCloseTag
from DEPPbreakdownTableFormat import orderStatusOpenTag, orderStatusCloseTag
from DEPPbreakdownTableFormat import DEPPNameOpenTag, DEPPNameCloseTag

arguments = []

for arg in sys.argv:
    arguments.append(arg)
arguments = arguments[1:]

try:
    int(arguments[0])
    reportDate = arguments[0]
except:
    reportDate = ''

# Cell Background and Font Styles (to be used to conditionally format cells)
below_goal_text = "9C0006"
below_goal_bg = "FFC7CE"
close_to_goal_text = "9C6500"
close_to_goal_bg = "FFEB9C"
at_or_above_goal_text = "006100"
at_or_above_goal_bg = "C6EFCE"

(jaelesiaDEPPsales, tekDEPPsales, antwonDEPPsales, jacksonDEPPsales,
 totalDEPPsales) = 0, 0, 0, 0, 0

supervisorIDs = {"aervin": 2062007, "jnickerson": 2062001, "tlevon": 2062007,
                 "jacksonn": 2062047, "jabram": 2062017,
                 "iqr_acollins": 2062072, "jmoore": 206223, "mayala": 2062002}



# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# Gather up the DEPP sales from the Products report
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
DEPP_sales_all = get_DEPP_sales(DEPPreportLocation(reportDate))

# remove any duplicates - there's gotta be a better way to do this!
DUPs_removed = []
for DEPP in DEPP_sales_all:
    if DEPP not in DUPs_removed:
          DUPs_removed.append(DEPP)

DEPP_sales_all = DUPs_removed

DEPP_sales = []

for sale in DEPP_sales_all:
    DEPP_sales.append(sale[0])

for id in DEPP_sales:
    if (type(id) == str):
        try:
            DEPP_sales[DEPP_sales.index(id)] = supervisorIDs[id]
        except:
            pass

# Sum up the DEPP sales for each supervisor and for the whole of iQor
for agentID in DEPP_sales:
    if agentID in jaelesiaTeam:
        jaelesiaDEPPsales += 1
        totalDEPPsales += 1
    if agentID in tekTeam:
        tekDEPPsales += 1
        totalDEPPsales += 1
    if agentID in antwonTeam:
        antwonDEPPsales += 1
        totalDEPPsales += 1
    if agentID in jacksonTeam:
        jacksonDEPPsales += 1
        totalDEPPsales += 1

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# Run through each entry in the tableNames and build the HTML string to be
# attached to the body of the email
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
html = topOfTable

for agentRow in tableNames:
    agentID = agentRow[0]
    agentName = agentRow[1]
    DEPPSales = ""
    DEPPSalesStart = DEPPSalesStartNoColor

    # This is executed if it is an agent and not a supervisor
    if (type(agentID) == int):  # only agents have numeric IDs

        # Get the agent DEPP Sales
        DEPPSales = DEPP_sales.count(agentID)

        # if (DEPPSales > 0):
        #     print(agentName, "printing green")
        #     DEPPSalesStart = DEPPSalesStartGreen

        DEPPSales = str(DEPP_sales.count(agentID))

        # Add the HTML string for the agent row
        agentID = str(agentID)
        html += (agentRowStart
                 + agentIDStart + agentID + agentIDEnd
                 + agentNameStart + agentName + agentNameEnd
                 + DEPPSalesStart + DEPPSales + DEPPSalesEnd
                 + agentRowEnd)

    # This is executed if it is a supervisor
    if (agentID == 'jaelesia' or agentID == 'tek' or
            agentID == 'antwon' or agentID == 'jackson'):

        if (agentID == 'jaelesia'):            
            DEPPSales = str(jaelesiaDEPPsales)


        elif (agentID == 'tek'):
            DEPPSales = str(tekDEPPsales) 

        elif (agentID == 'antwon'):
            DEPPSales = str(antwonDEPPsales)

        elif (agentID == 'jackson'):
            DEPPSales = str(jacksonDEPPsales)

        # Add the HTMl string for the supervisor
        agentID = "&nbsp;"
        html += (supRowStart
                 + supIDStart + agentID + agentIDEnd
                 + supNameStart + agentName + agentNameEnd
                 + supDEPPSalesStart + DEPPSales + DEPPSalesEnd
                 + supRowEnd)

    # This is executed if it is grand Total
    if agentID == 'grandTotal':
        DEPPSales = str(totalDEPPsales)

        # Add the HTML string for the Grand Total
        agentID = "&nbsp;"
        html += (grandTotalRowStart
                 + gTotalIDStart + agentID + agentIDEnd
                 + gTotalNameStart + agentName + agentNameEnd
                 + gTotalDEPPSalesStart + DEPPSales + DEPPSalesEnd
                 + grandTotalRowEnd + "</table> <br> <br>")

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# DEPP Sales Breakdown
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------

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

html += salesDEPPTableOpenTag

for DEPP in DEPP_sales:
    # format will be [bounce_sale, DEPP_sales]
    # an empty [] means that it is a partially blank row, and
    # one of the two, bounce_sales or DEPP_sales has more rows than the other
    # we will test for this unevenness by checking the length
    # bounceSale = row[0]
    # DEPPSale = row[1]
    
    # if len(row) > 0:
    agentName2 = DEPP[0]
    accountNumber2 = str(int(DEPP[1]))
    orderNumber2 = str(int(DEPP[2]))
    DEPPName = DEPP[3]
    orderStatus2 = DEPP[4]

    html += (rowOpenTag
             + agentNameOpenTag + agentName2 + agentNameCloseTag
             + acctNumOpenTag + accountNumber2 + acctNumCloseTag
             + orderNumOpenTag + orderNumber2 + orderNumCloseTag
             + DEPPNameOpenTag + DEPPName + DEPPNameCloseTag
             + orderStatusOpenTag + orderStatus2 + orderStatusCloseTag
             + rowCloseTag)

html += tableCloseTag + emailEndHtml




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
    subject = 'iQor DEPP Report ' + reportDate + ' End of Business'
    additionalEmailList = "; ".join(arguments[1:])

except:
    reportDate = ''
    subject = 'iQor DEPP Update ' + currentDate + ' ' + currentTime
    additionalEmailList = "; ".join(arguments[0:])

mail.To = additionalEmailList + '; jackson.ndiho@iqor.com'
mail.Subject = subject
mail.HtmlBody = subject + ":" + html
mail.send

print("\nDEPP Sales email sent to: " + additionalEmailList
      + "; jackson.ndiho@iqor.com \nat " + currentDate + " " + currentTime 
      + "\n\nDone.......")