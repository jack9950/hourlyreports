import sys
import time
from datetime import datetime
import win32com.client as win32
from get_calls_handled import get_calls_handled
from get_pogo_sales import get_pogo_sales
from get_DEPP_sales import get_DEPP_sales
from get_fcp_sales import get_fcp_sales
from data_files import callsHandledReportLocation, pogoSalesReportLocation
from data_files import fcpReportLocation, DEPPreportLocation
from data_files import tableNames
from data_files import jaelesiaTeam, tekTeam, antwonTeam
from tableformat import topOfTable
from tableformat import agentRowStart, agentRowEnd
from tableformat import agentIDStart, agentIDEnd
from tableformat import agentNameStart, agentNameEnd
from tableformat import callsHandledStart, callsHandledEnd
from tableformat import salesCallsHandledStart, salesCallsHandledEnd
from tableformat import bounceSalesStart, bounceSalesEnd
from tableformat import closeRateStartRed, closeRateStartYellow, closeRateStartGreen, closeRateStartNoColor, closeRateEnd
from tableformat import FCPSalesStart, FCPSalesEnd
from tableformat import DEPPSalesStart, DEPPSalesEnd
from tableformat import supRowStart, supRowEnd, grandTotalRowStart, grandTotalRowEnd
from tableformat import supIDStart, supNameStart, supCallsHandledStart
from tableformat import supSalesCallsHandledStart
from tableformat import supBounceSalesStart, supCloseRateStartRed
from tableformat import supCloseRateStartYellow, supCloseRateStartGreen, supCloseRateStartNoColor
from tableformat import supFCPSalesStart, supDEPPSalesStart
from tableformat import gTotalIDStart, gTotalNameStart, gTotalCallsHandledStart
from tableformat import gTotalSalesCallsHandledStart
from tableformat import gTotalBounceSalesStart, gTotalCloseRateStartRed
from tableformat import gTotalCloseRateStartYellow, gTotalCloseRateStartGreen
from tableformat import gTotalFCPSalesStart, gTotalDEPPSalesStart

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


agentIDs = [2062004, 2062026, 2062043, 2062034, 2062053, 2062048, 2062042,
            2062011, 2062030, 2062045, 2062046, 2062016, 2062001, 2062036,
            2062039, 2062025, 2062041, 2062052, 2062037, 2062024, 2062049,
            2062031, 2062044, 2062003, 2062028, 2062022, 2062051, 2062021,
            2062035, 2062007, 2062020, 2062015, 2062040, 2062010, 2062018,
            2062054, 2062032, 2062033, 2062062, 2062070, 2062067, 2062058,
            2062056, 2062066, 2062057, 2062065, 2062060]

(jaelesiaTotalCallsHandled, tekTotalCallsHandled, antwonTotalCallsHandled,
 totalCallsHandled) = 0, 0, 0, 0
(jaelesiaSalesCallsHandled, tekSalesCallsHandled, antwonSalesCallsHandled,
 totalSalesCallsHandled) = 0, 0, 0, 0
(jaelesiaTotalSales, tekTotalSales, antwonTotalSales,
 totalSales) = 0, 0, 0, 0
(jaelesiaFCPsales, tekFCPsales, antwonFCPsales, totalFCPSales) = 0, 0, 0, 0
(jaelesiaDEPPsales, tekDEPPsales, antwonDEPPsales,
 totalDEPPsales) = 0, 0, 0, 0

supervisorIDs = {"aervin": 2062007, "jnickerson": 2062001, "tlevon": 2062007,
                 "jacksonn": 2062047, "jabram": 2062017,
                 "iqr_acollins": 2062072, "jmoore": 206223, "mayala": 2062002}

"""
agentRowStart
+ agentIDStart + agentID + agentIDEnd
+ agentNameStart + agentName + agentNameEnd
+ callsHandledStart + callsHandled + callsHandledEnd
+ salesCallsHandledStart + salesCallsHandled + salesCallsHandledEnd
+ bounceSalesStart + bounceSales + bounceSalesEnd
+ closeRateStartRed + closeRate + closeRateEnd
+ FCPSalesStart + FCPSales + FCPSalesEnd
+ DEPPSalesStart + DEPPSales + DEPPSalesEnd
+ agentRowEnd

"""

html = topOfTable

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# Get the calls handled for each agent
# The format returned is a 2 dimensional array with each agent and their calls
# represented as:
# [agent ID, Calls Handled, Sales Calls Handled] in the return array
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
calls_handled = get_calls_handled(callsHandledReportLocation(reportDate))

# Sum up the calls handled for each supervisor and for the whole of iQor
for item in calls_handled:
    agentID = item[0]
    if agentID in jaelesiaTeam:
        jaelesiaTotalCallsHandled += item[1]
        jaelesiaSalesCallsHandled += item[2]
        totalCallsHandled += item[1]
        totalSalesCallsHandled += item[2]
    if agentID in tekTeam:
        tekTotalCallsHandled += item[1]
        tekSalesCallsHandled += item[2]
        totalCallsHandled += item[1]
        totalSalesCallsHandled += item[2]
    if agentID in antwonTeam:
        antwonTotalCallsHandled += item[1]
        antwonSalesCallsHandled += item[2]
        totalCallsHandled += item[1]
        totalSalesCallsHandled += item[2]

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# Gather up all the orders from the big bounce sales report
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
pogo_sales = get_pogo_sales(pogoSalesReportLocation(reportDate))

# Team leads also submit POGO orders with text POGO ID rather than numeric ID
# Replace the team lead text POGO agent IDs with the numeric AVAYA IDs
for id in pogo_sales:
    if (type(id) == str):
        try:
            pogo_sales[pogo_sales.index(id)] = supervisorIDs[id]
        except:
            pass

# Sum up the POGO sales for each supervisor and for the whole of iQor
for agentID in pogo_sales:
    if agentID in jaelesiaTeam:
        jaelesiaTotalSales += 1
        totalSales += 1
    if agentID in tekTeam:
        tekTotalSales += 1
        totalSales += 1
    if agentID in antwonTeam:
        antwonTotalSales += 1
        totalSales += 1


# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# Gather up the FCP sales from the FCP report
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
fcp_sales = get_fcp_sales(fcpReportLocation(reportDate), reportDate)

# Sum up the fcp sales for each supervisor and for the whole of iQor
for agentID in fcp_sales:
    if agentID in jaelesiaTeam:
        jaelesiaFCPsales += 1
        totalFCPSales += 1
    if agentID in tekTeam:
        tekFCPsales += 1
        totalFCPSales += 1
    if agentID in antwonTeam:
        antwonFCPsales += 1
        totalFCPSales += 1


# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# Gather up the DEPP sales from the Products report
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
DEPP_sales = get_DEPP_sales(DEPPreportLocation(reportDate))

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


# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# Run through each entry in the tableNames and build the HTML string to be
# attached to the body of the email
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
for agentRow in tableNames:
    agentID = agentRow[0]
    agentName = agentRow[1]
    callsHandled = ""
    salesCallsHandled = ""
    bounceSales = ""
    closeRate = ""
    FCPSales = ""
    DEPPSales = ""
    closeRateStart = closeRateStartNoColor
    supCloseRateStart = supCloseRateStartNoColor

    # This is executed if it is an agent and not a supervisor
    if (type(agentID) == int):  # only agent have numeric IDs

        # Get the agent calls handled and sales calls handled
        for item in calls_handled:  # check each nested list
            if (agentID == item[0]):
                callsHandledInteger = item[1]
                salesCallsHandledInteger = item[2]
                if callsHandledInteger > 0:
                    callsHandled = str(int(item[1]))
                    salesCallsHandled = str(int(item[2]))

        # Get the agent Bounce Sales
        if (callsHandled is not ""):
            bounceSales = str(pogo_sales.count(agentID))
            bounceSalesInteger = pogo_sales.count(agentID)

        # Get the agent FCP Sales
        if (callsHandled is not ""):
            FCPSales = str(fcp_sales.count(agentID))
            FCPSalesInteger = fcp_sales.count(agentID)

        # Get the agent Close Rate
        if (salesCallsHandled is not ""):
            if (int(salesCallsHandled) > 0):
                closeRate = ((bounceSalesInteger + FCPSalesInteger) /
                             salesCallsHandledInteger * 100.00)
                closeRate = int(round(closeRate, 0))
                if closeRate > 49:
                    closeRateStart = closeRateStartGreen
                elif closeRate > 39:
                    closeRateStart = closeRateStartYellow
                else:
                    closeRateStart = closeRateStartRed
                closeRate = str(closeRate) + "%"

        # Get the agent DEPP Sales
        if (callsHandled is not ""):
            DEPPSales = str(DEPP_sales.count(agentID))

        # Add the HTML string for the agent row
        agentID = str(agentID)
        html += (agentRowStart
                 + agentIDStart + agentID + agentIDEnd
                 + agentNameStart + agentName + agentNameEnd
                 + callsHandledStart + callsHandled + callsHandledEnd
                 + salesCallsHandledStart + salesCallsHandled + salesCallsHandledEnd
                 + bounceSalesStart + bounceSales + bounceSalesEnd
                 + closeRateStart + closeRate + closeRateEnd
                 + FCPSalesStart + FCPSales + FCPSalesEnd
                 + DEPPSalesStart + DEPPSales + DEPPSalesEnd
                 + agentRowEnd)

    # This is executed if it is a supervisor
    if (agentID == 'jaelesia' or agentID == 'tek' or
            agentID == 'antwon'):

        if (agentID == 'jaelesia'):
            callsHandled = str(int(jaelesiaTotalCallsHandled)
                               ) if jaelesiaTotalCallsHandled else ""
            salesCallsHandled = str(int(jaelesiaSalesCallsHandled)
                                    ) if jaelesiaTotalCallsHandled else ""
            bounceSales = str(jaelesiaTotalSales) if jaelesiaSalesCallsHandled else ""
            FCPSales = str(jaelesiaFCPsales) if jaelesiaSalesCallsHandled else ""
            DEPPSales = str(jaelesiaDEPPsales) if jaelesiaTotalCallsHandled else ""

        elif (agentID == 'tek'):
            callsHandled = str(int(tekTotalCallsHandled)) if tekTotalCallsHandled else ""
            salesCallsHandled = str(int(tekSalesCallsHandled)) if tekTotalCallsHandled else ""
            bounceSales = str(tekTotalSales) if tekSalesCallsHandled else ""
            FCPSales = str(tekFCPsales) if tekSalesCallsHandled else ""
            DEPPSales = str(tekDEPPsales) if tekTotalCallsHandled else ""

        elif (agentID == 'antwon'):
            callsHandled = str(int(antwonTotalCallsHandled)) if antwonTotalCallsHandled else ""
            salesCallsHandled = str(int(antwonSalesCallsHandled)) if antwonTotalCallsHandled else ""
            bounceSales = str(antwonTotalSales) if antwonSalesCallsHandled else ""
            FCPSales = str(antwonFCPsales) if antwonSalesCallsHandled else ""
            DEPPSales = str(antwonDEPPsales) if antwonTotalCallsHandled else ""

        # Calculate Jaelesia's close rate and the colors for her cells
        if (agentID == 'jaelesia'):
            if (jaelesiaSalesCallsHandled is not ""):
                if (int(jaelesiaSalesCallsHandled) > 0):
                    closeRate = ((jaelesiaTotalSales + jaelesiaFCPsales) /
                                 jaelesiaSalesCallsHandled * 100.00)
                    closeRate = int(round(closeRate, 0))
                    if closeRate >= 50:
                        supCloseRateStart = supCloseRateStartGreen
                    elif closeRate >= 40:
                        supCloseRateStart = supCloseRateStartYellow
                    else:
                        supCloseRateStart = supCloseRateStartRed
                    closeRate = str(closeRate) + "%"

        # Calculate Tek's close rate and the colors for his cells
        if (agentID == 'tek'):
            if (tekSalesCallsHandled is not ""):
                if (int(tekSalesCallsHandled) > 0):
                    closeRate = ((tekTotalSales + tekFCPsales) /
                                 tekSalesCallsHandled * 100.00)
                    closeRate = int(round(closeRate, 0))
                    if closeRate >= 50:
                        supCloseRateStart = supCloseRateStartGreen
                    elif closeRate >= 40:
                        supCloseRateStart = supCloseRateStartYellow
                    else:
                        supCloseRateStart = supCloseRateStartRed
                    closeRate = str(closeRate) + "%"

        # Calculate Antwon's close rate and the colors for his cells
        if (agentID == 'antwon'):
            if (antwonSalesCallsHandled is not ""):
                if (int(antwonSalesCallsHandled) > 0):
                    closeRate = ((antwonTotalSales + antwonFCPsales) /
                                 antwonSalesCallsHandled * 100.00)
                    closeRate = int(round(closeRate, 0))
                    if closeRate >= 50:
                        supCloseRateStart = supCloseRateStartGreen
                    elif closeRate >= 40:
                        supCloseRateStart = supCloseRateStartYellow
                    else:
                        supCloseRateStart = supCloseRateStartRed
                    closeRate = str(closeRate) + "%"

        # Add the HTMl string for the supervisor
        agentID = "&nbsp;"
        html += (supRowStart
                 + supIDStart + agentID + agentIDEnd
                 + supNameStart + agentName + agentNameEnd
                 + supCallsHandledStart + callsHandled + callsHandledEnd
                 + supSalesCallsHandledStart + salesCallsHandled + salesCallsHandledEnd
                 + supBounceSalesStart + bounceSales + bounceSalesEnd
                 + supCloseRateStart + closeRate + closeRateEnd
                 + supFCPSalesStart + FCPSales + FCPSalesEnd
                 + supDEPPSalesStart + DEPPSales + DEPPSalesEnd
                 + supRowEnd)

    # This is executed if it is grand Total
    if agentID == 'grandTotal':
        callsHandled = str(int(totalCallsHandled))
        salesCallsHandled = str(int(totalSalesCallsHandled))
        bounceSales = str(totalSales)
        FCPSales = str(totalFCPSales)
        DEPPSales = str(totalDEPPsales)

        if (totalSalesCallsHandled is not ""):
            if (int(totalSalesCallsHandled) > 0):
                closeRate = ((totalSales + totalFCPSales) /
                             totalSalesCallsHandled * 100.00)
                closeRate = int(round(closeRate, 0))
                if closeRate >= 50:
                    supCloseRateStart = gTotalCloseRateStartGreen
                elif closeRate >= 40:
                    supCloseRateStart = gTotalCloseRateStartYellow
                else:
                    supCloseRateStart = gTotalCloseRateStartRed
                closeRate = str(closeRate) + "%"

        # Add the HTML string for the Grand Total
        agentID = "&nbsp;"
        html += (grandTotalRowStart
                 + gTotalIDStart + agentID + agentIDEnd
                 + gTotalNameStart + agentName + agentNameEnd
                 + gTotalCallsHandledStart + callsHandled + callsHandledEnd
                 + gTotalSalesCallsHandledStart + salesCallsHandled + salesCallsHandledEnd
                 + gTotalBounceSalesStart + bounceSales + bounceSalesEnd
                 + supCloseRateStart + closeRate + closeRateEnd
                 + gTotalFCPSalesStart + FCPSales + FCPSalesEnd
                 + gTotalDEPPSalesStart + DEPPSales + DEPPSalesEnd
                 + grandTotalRowEnd)

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
    subject = 'iQor Sales Report ' + reportDate + ' End of Business'
    additionalEmailList = "; ".join(arguments[1:])

except:
    reportDate = ''
    subject = 'iQor Sales Update ' + currentDate + ' ' + currentTime
    additionalEmailList = "; ".join(arguments[0:])

mail.To = additionalEmailList + '; jackson.ndiho@iqor.com'
mail.Subject = subject
mail.HtmlBody = subject + ":" + html
mail.send

print("\nEmail sent to: " + additionalEmailList + "; jackson.ndiho@iqor.com.\n\nDone.......")
