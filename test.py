import csv
from data_files import homeFolder
from teams import agent_ids_to_names
from data_files import jaelesiaTeam, tekTeam, antwonTeam, jacksonTeam
from get_DEPP_sales2 import get_DEPP_sales, get_DEPP_sales_breakdown


(jaelesiaDEPPsales, tekDEPPsales, antwonDEPPsales, jacksonDEPPsales,
 totalDEPPsales) = 0, 0, 0, 0, 0

supervisorIDs = {"aervin": 2062007, "jnickerson": 2062001, "tlevon": 2062007,
                 "jacksonn": 2062047, "jabram": 2062017,
                 "iqr_acollins": 2062072, "jmoore": 206223, "mayala": 2062002}

DEPPFileName = homeFolder + 'report.csv'
DEPP_sales_all = get_DEPP_sales(DEPPFileName)

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

# for sale in DEPP_sales:
# 	print(sale, '\n')
# print(DEPP_sales)
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

DEPP_sales = get_DEPP_sales_breakdown(DEPPFileName)

# remove any duplicates - there is probably a better way to do this!
DUPs_removed = []
for DEPP in DEPP_sales:
    if DEPP not in DUPs_removed:
          DUPs_removed.append(DEPP)
DEPP_sales = DUPs_removed

DEPP_sales.sort()

for sale in DEPP_sales:
	print(sale, '\n')