import sys
from get_AHT import get_AHT
from data_files import callsHandledReportLocation
from data_files import jacksonTeam, tableNamesJackson

arguments = []

for arg in sys.argv:
    arguments.append(arg)
arguments = arguments[1:]

try:
    int(arguments[0])
    reportDate = arguments[0]
except:
    reportDate = ''


calls_handled = get_AHT(callsHandledReportLocation(reportDate))

jackson_AHT = []
for data in calls_handled:
	if data[0] in jacksonTeam:
		jackson_AHT.append(data)

# for item in jackson_AHT:
#     agentID = item[0]
#     jacksonTotalCallsHandled += item[1]
#     jacksonSalesCallsHandled += item[2]
#     totalCallsHandled += item[1]
#     totalSalesCallsHandled += item[2]

for agentData in jackson_AHT:
	print(agentData)