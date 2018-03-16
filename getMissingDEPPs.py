import sys
import xlrd
import time
from datetime import datetime
from data_files import agent_ids_to_names
import csv
from data_files import homeFolder
from data_files import jaelesiaTeam, tekTeam, antwonTeam, jacksonTeam

#Sample return:
# [agent_id, [Acct #, Order #, order status], [Acct #, Order #, order status]]
# [2062062, [2092985, 1443822, "Deposit due"], [2092021, 1444496, "Ercot/ISO Processing"] ]

def get_missing_DEPPs(filename, freq='hourly'):

    currentDate = datetime.now().strftime("%m/%d/%Y")
    currentMonth = datetime.now().strftime("%m")

    # print('currentDate', currentDate)
    with open(filename) as DEPPFile:
        DEPPReader = csv.reader(DEPPFile)
        DEPPData = list(DEPPReader)

    values = []
    for row in DEPPData:
        
        agent_id = row[0]
        DEPP_name = row[3]
        reportDate = row[5]
        # print('reportDate = ', reportDate)

        # print('agent_id: ', agent_id, 'DEPP_name: ', DEPP_name)

        if(freq == 'hourly' and reportDate == currentDate):

            if(agent_id != '' and (DEPP_name == "Surge Protection Plan" or
                                     DEPP_name == "Electric Repair Essentials" or
                                     DEPP_name == "Surge Protection Plan (20% Off)" or
                                     DEPP_name == "Cooling Maintenance Essentials (6 Month Free Trial - Nest Bundle)" or
                                     DEPP_name == "Cooling Repair & Maintenance Essentials" or
                                     DEPP_name == "Electric Repair Essentials (20% Off)") or
                                     DEPP_name == "Heating & Cooling Repair Essentials"):
                try:
                    agent_name = agent_ids_to_names[int(agent_id)]
                    pogo_account_number = row[1]
                    pogo_order_number = row[2]
                    DEPP_name = row[3]
                    bounce_status = row[4]

                    values.append([int(agent_id),
                                   int(pogo_account_number),
                                   int(pogo_order_number),
                                   DEPP_name,
                                   bounce_status])
                except:
                    pass

        if(freq == 'MTD' and reportDate[:2] == currentMonth):

            if(agent_id != '' and (DEPP_name == "Surge Protection Plan" or
                                     DEPP_name == "Electric Repair Essentials" or
                                     DEPP_name == "Surge Protection Plan (20% Off)" or
                                     DEPP_name == "Cooling Maintenance Essentials (6 Month Free Trial - Nest Bundle)" or
                                     DEPP_name == "Cooling Repair & Maintenance Essentials" or
                                     DEPP_name == "Electric Repair Essentials (20% Off)") or
                                     DEPP_name == "Heating & Cooling Repair Essentials"):
                try:
                    agent_name = agent_ids_to_names[int(agent_id)]
                    pogo_account_number = row[1]
                    pogo_order_number = row[2]
                    DEPP_name = row[3]
                    bounce_status = row[4]

                    values.append([int(agent_id),
                                   int(pogo_account_number),
                                   int(pogo_order_number),
                                   DEPP_name,
                                   bounce_status])
                except:
                    pass

        if(freq == 'lastMonth' and reportDate[:2] == str(int(currentMonth)-1)):

            if(agent_id != '' and (DEPP_name == "Surge Protection Plan" or
                                     DEPP_name == "Electric Repair Essentials" or
                                     DEPP_name == "Surge Protection Plan (20% Off)" or
                                     DEPP_name == "Cooling Maintenance Essentials (6 Month Free Trial - Nest Bundle)" or
                                     DEPP_name == "Cooling Repair & Maintenance Essentials" or
                                     DEPP_name == "Electric Repair Essentials (20% Off)") or
                                     DEPP_name == "Heating & Cooling Repair Essentials"):
                try:
                    agent_name = agent_ids_to_names[int(agent_id)]
                    pogo_account_number = row[1]
                    pogo_order_number = row[2]
                    DEPP_name = row[3]
                    bounce_status = row[4]

                    values.append([int(agent_id),
                                   int(pogo_account_number),
                                   int(pogo_order_number),
                                   DEPP_name,
                                   bounce_status])
                except:
                    pass

    return values

def get_missing_DEPPs_breakdown(filename, freq='hourly'):
# first open using xlrd    book = xlrd.open_workbook(filename)
    
    # Three frequencies: hourly, MTD and lastMonth

    currentDate = datetime.now().strftime("%m/%d/%Y")

    with open(filename) as DEPPFile:
        DEPPReader = csv.reader(DEPPFile)
        DEPPData = list(DEPPReader)
        # print('DEPPData: ', DEPPData)

    values = []

    for row in DEPPData:
        agent_id = row[0]
        DEPP_name = row[3]
        dateToday = row[5]
        print('dateToday = ', dateToday)

        # print('date from csv: ', row[5])
        # print('row[5] == currentDate: ', row[5] == currentDate)
        # print('agent_id: ', agent_id, 'DEPP_name: ', DEPP_name)
        if(dateToday == 'October'):
            
            if(agent_id != None and (DEPP_name == "Surge Protection Plan" or
                                     DEPP_name == "Electric Repair Essentials" or
                                     DEPP_name == "Surge Protection Plan (20% Off)" or
                                     DEPP_name == "Cooling Maintenance Essentials (6 Month Free Trial - Nest Bundle)" or
                                     DEPP_name == "Cooling Repair & Maintenance Essentials" or
                                     DEPP_name == "Electric Repair Essentials (20% Off)") or
                                     DEPP_name == "Heating & Cooling Repair Essentials"):
                try:
                    agent_name = agent_ids_to_names[int(agent_id)]
                    pogo_account_number = row[1]
                    pogo_order_number = row[2]
                    DEPP_name = row[3]
                    bounce_status = row[4]
                    # print("\n success!")
                    # print('agent_name: ', agent_name)
                    values.append([agent_name,
                                   int(pogo_account_number),
                                   int(pogo_order_number),
                                   DEPP_name,
                                   bounce_status])
                    # print('   values: ', values)
                except:
                    pass

    # DEPPFile.close()
    return values
