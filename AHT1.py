import sys
import time
from datetime import datetime
import win32com.client as win32
from get_AHT1 import get_AHT
from data_files import trainingTeam, tableNamesJackson
from AHTformat import topOfTable
from AHTformat import agentRowStart, agentRowEnd
from AHTformat import agentIDStart, agentIDEnd
from AHTformat import agentNameStart, agentNameEnd
from AHTformat import signInTimeStart, signInTimeEnd
from AHTformat import callsHandledStart, callsHandledEnd
from AHTformat import AHTStart, AHTEnd
from AHTformat import grandTotalRowStart, grandTotalRowEnd
from AHTformat import grandTotalAgentID, grandTotalAgentName
from AHTformat import grandTotalSignInTime
from AHTformat import grandTotalCallsHandledStart, grandTotalCallsHandledEnd
from AHTformat import grandTotalAHTStart, grandTotalAHTEnd 

arguments = []

for arg in sys.argv:
    arguments.append(arg)
arguments = arguments[1:]

try:
    int(arguments[0])
    reportDate = arguments[0]
except:
    reportDate = ''

homeFolder = 'C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\'

callsHandledReportLocation = homeFolder + 'BounceAgentLevelReport.xls'

calls_handled = get_AHT(callsHandledReportLocation)

jackson_AHT = []
for data in calls_handled:
	if data[0] in trainingTeam:
		jackson_AHT.append(data)

for agentData in jackson_AHT:
	print(agentData)

html = topOfTable
totalCallsHandled = 0
totalAHT = 0
numberOfAgents = 0

for agentData in jackson_AHT:
	agentID = str(agentData[0])
	agentName = agentData[1]
	signInTime = agentData[2]
	callsHandled = str(agentData[3])
	AHT = str(agentData[4])
	totalCallsHandled += agentData[3]
	totalAHT += agentData[4]
	numberOfAgents += 1

	html += (agentRowStart
                 + agentIDStart + agentID + agentIDEnd
                 + agentNameStart + agentName + agentNameEnd
                 + signInTimeStart + signInTime + signInTimeEnd
                 + callsHandledStart + callsHandled + callsHandledEnd
                 + AHTStart + AHT + AHTEnd       
                 + agentRowEnd)

averageAHT = str(int(totalAHT / numberOfAgents))
totalCallsHandled = str(totalCallsHandled)

html += (grandTotalRowStart
			+ grandTotalAgentID
			+ grandTotalAgentName
			+ grandTotalSignInTime
			+ grandTotalCallsHandledStart + totalCallsHandled + grandTotalCallsHandledEnd
			+ grandTotalAHTStart + averageAHT + grandTotalAHTEnd
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
    subject = 'AHT Report ' + reportDate + ' Final'
    additionalEmailList = "; ".join(arguments[1:])

except:
    reportDate = ''
    subject = 'AHT Update ' + currentDate + ' ' + currentTime
    additionalEmailList = "; ".join(arguments[0:])

mail.To = additionalEmailList + '; jackson.ndiho@iqor.com'
mail.Subject = subject
mail.HtmlBody = subject + ":" + html
mail.send

print("\niQor Sales email sent to: " + additionalEmailList
      + "; jackson.ndiho@iqor.com \nat " + currentDate + " " + currentTime 
      + "\n\nDone.......")