import openpyxl
from openpyxl.styles import Font
from get_agent_ids_and_calls import get_agent_ids_and_calls
from get_pogo_sales import get_pogo_sales
from get_nest_sales import get_nest_sales
from get_warranty_sales import get_warranty_sales
from get_fcp_sales import get_fcp_sales
from get_hive_sales import get_hive_sales

jaelesiaTeam = [2062001, 2062011, 2062020, 2062026, 2062036, 2062048, 2062053,
                2062054, 2062057, 2062062]
tekTeam = [2062067, 2062051, 2062035, 2062015, 2062040, 2062010, 2062042,
           2062024, 2062065, 2062060, 2062007]
antwonTeam = [2062039, 2062073, 2062074, 2062052, 2062058, 2062018, 2062049,
              2062076, 2062031, 2062044, 2062003, 2062032, 2062066]

agentIDs = [2062004, 2062026, 2062043, 2062034, 2062053, 2062048, 2062042,
            2062011, 2062030, 2062045, 2062046, 2062016, 2062001, 2062036,
            2062039, 2062025, 2062041, 2062052, 2062037, 2062024, 2062049,
            2062031, 2062044, 2062003, 2062028, 2062022, 2062051, 2062021,
            2062035, 2062007, 2062020, 2062015, 2062040, 2062010, 2062018,
            2062054, 2062032, 2062033, 2062062, 2062070, 2062067, 2062058,
            2062056, 2062066, 2062057, 2062065, 2062060]

jaelesiaTotalCallsHandled = 0
tekTotalCallsHandled = 0
antwonTotalCallsHandled = 0
totalCallsHandled = 0

jaelesiaSalesCallsHandled = 0
tekSalesCallsHandled = 0
antwonSalesCallsHandled = 0
totalSalesCallsHandled = 0

jaelesiaTotalSales = 0
tekTotalSales = 0
antwonTotalSales = 0
totalSales = 0

jaelesiaFCPsales = 0
tekFCPsales = 0
antwonFCPsales = 0
totalFCPSales = 0

jaelesiaNestSales = 0
tekNestSales = 0
antwonNestSales = 0
totalNestSales = 0

jaelesiaDEPPsales = 0
tekDEPPsales = 0
antwonDEPPsales = 0
totalDEPPsales =0

jaelesiaHiveSales = 0
tekHiveSales = 0
antwonHiveSales = 0
totalHiveSales = 0

supervisorIDs = {"aervin":2062007, "jnickerson":2062001, "tlevon": 2062007,
                 "jacksonn": 2062047, "jabram":2062017, "iqr_acollins":2062072,
                 "jmoore":2062023, "mayala":2062002}

homeFolder = 'C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\'

callsHandledReportLocation = homeFolder +'Bounce_Hourly_Sales_Report_03062017.xls'
pogoSalesReportLocation = homeFolder + 'bounce_energy_iqor_report_21.xls'
productsReportLocation = homeFolder + 'BounceEnergyProducts Added2017-03-06.xls'
fcpReportLocation = homeFolder + 'HourlyProducts_Added.xls'
hiveReportLocation = homeFolder + 'products_sonar_03062017.xls'

#Open the template file for editing:
print("\nOpening template file for editing......\n")

template = openpyxl.load_workbook("C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\template.xlsx")
template_sheets = template.get_sheet_names()
template_first_sheet = template.get_sheet_by_name(template_sheets[0])


#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#Open the Bounce Hourly Sales Report from iQor,
#Retrieve data and add to summary tab of template excel file
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------

#Get the calls handled for each agent

print("\nReading agent IDs and call counts.......\n")

#The format returned is a 2 dimensional array with each agent and their calls represented as:
#[agent ID, Calls Handled, Sales Calls Handled] in the return array
calls_handled = get_agent_ids_and_calls(callsHandledReportLocation)

#Write out the call counts to the template file
print("\nWriting call counts to the template file.......\n")
for row in range(3, 60):
    agent_id_cell = "A" + str(row)
    agent_id = template_first_sheet[agent_id_cell].value
    #retrieve each agent ID from the template and check if it is the list.
    #if found write the calls handled and sales calls handled to the template file
    if(agent_id != None):
        for item in calls_handled: #check each nested list
            if agent_id in item and item[1] != 0: #if found and total calls > 0,
                                                  #write calls data to the template
                template_first_sheet["C"+str(row)].value = item[1] #Total Calls Handled
                template_first_sheet["D"+str(row)].value = item[2] #Sales Calls Handled

#Sum up the calls handled for each supervisor and for the whole of iQor
for item in calls_handled:
    agent_id = item[0]
    if agent_id in jaelesiaTeam:
        jaelesiaTotalCallsHandled += item[1]
        jaelesiaSalesCallsHandled += item[2]
        totalCallsHandled += item[1]
        totalSalesCallsHandled += item[2]
    if agent_id in tekTeam:
        tekTotalCallsHandled += item[1]
        tekSalesCallsHandled += item[2]
        totalCallsHandled += item[1]
        totalSalesCallsHandled += item[2]
    if agent_id in antwonTeam:
        antwonTotalCallsHandled += item[1]
        antwonSalesCallsHandled += item[2]
        totalCallsHandled += item[1]
        totalSalesCallsHandled += item[2]

# print("jaelesia: ", jaelesiaTotalCallsHandled, jaelesiaSalesCallsHandled)
# print("tek: ", tekTotalCallsHandled, tekSalesCallsHandled)
# print("antwon: ", antwonTotalCallsHandled, antwonSalesCallsHandled)
# print("totalCallsHandled: ", totalCallsHandled)
# print("totalSalesCallsHandled", totalSalesCallsHandled)


#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#Open the Hourly Orders Placed Report from Big Bounce,
#retrieve date and add to summary tab of template excel file
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#Convert(to xlsx) and open the hourly orders sales report sent by Big Bounce for reading:
#convert_file('bounce_energy_iqor_report.xls')
# wb = openpyxl.load_workbook('bounce_energy_iqor_report_20.xlsx')
# sheets = wb.get_sheet_names()
# sheet = wb.get_sheet_by_name(sheets[0])

#Gather up all the orders from the big bounce sales report:
print("\nGathering up all the orders from the big bounce sales report.......\n")
# pogo_sales = [];
# for i in range(2,100):
#     agent_id_cell = "G" + str(i)
#     if(sheet[agent_id_cell].value != None):
#         pogo_sales.append(sheet[agent_id_cell].value)

pogo_sales = get_pogo_sales(pogoSalesReportLocation)

#For those agents that have their own non numeric POGO logins,
#replace the POGO text usernames with the numeric AVAYA IDs
for id in pogo_sales:
    if (type(id) == str):
        try:
            pogo_sales[pogo_sales.index(id)] = supervisorIDs[id]
        except:
            pass

#Write out the number of sales to the template
print("\nWriting out the number of sales to the template.......\n")
for i in range(3, 50):
    agent_id_cell = "A"+str(i)
    agent_id = template_first_sheet[agent_id_cell].value
    calls_handled_cell = "C"+str(i)
    calls_handled = template_first_sheet[calls_handled_cell].value
    if(agent_id != None and calls_handled != None):
        template_first_sheet["E"+str(i)].value = pogo_sales.count(agent_id)
        #print(agent_id, ": ", pogo_sales.count(agent_id))

#Sum up the POGO sales for each supervisor and for the whole of iQor
for agent_id in pogo_sales:
    if agent_id in jaelesiaTeam:
        jaelesiaTotalSales += 1
        totalSales += 1
    if agent_id in tekTeam:
        tekTotalSales += 1
        totalSales += 1
    if agent_id in antwonTeam:
        antwonTotalSales += 1
        totalSales += 1

# print("jaelesiaTotalSales: ", jaelesiaTotalSales)
# print("tekTotalSales: ", tekTotalSales)
# print("antwonTotalSales: ", antwonTotalSales)
# print("totalSales", totalSales)

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#Retrieve NEST and DEPP data and add to summary tab of template excel file
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------

print("\nGathering NEST and warranty sales......\n")

#Get the NEST sales - the returned calue (nest_sales) is a list of agent IDs
#Every agent ID in the list is an agent ID that has a NEST sale
nest_sales = get_nest_sales(productsReportLocation)

#Get the Warranty sales - the returned calue (warranty_sales) is a list of agent IDs
#Every agent ID in the list is an agent ID that has a Warranty sale
warranty_sales = get_warranty_sales(productsReportLocation)

#Team leads usually submit NEST orders with their text POGO ID rather than the numeric one
#Replace the team lead text POGO agent IDs with the numeric
for id in nest_sales:
    if (type(id) == str):
        try:
            nest_sales[nest_sales.index(id)] = supervisorIDs[id]
        except:
            pass

#Team leads usually submit Warranty orders with their text POGO ID rather than the numeric one
#Replace the team lead text POGO agent IDs with the numeric
for id in warranty_sales:
    if (type(id) == str):
        try:
            nest_sales[nest_sales.index(id)] = supervisorIDs[id]
        except:
            pass

#Write out the products to the template
print("\nWriting the products to the template..........\n")
for i in range(3, 50):
    agent_id_cell = "A"+str(i)
    calls_handled_cell = "C"+str(i)
    agent_id = template_first_sheet[agent_id_cell].value
    calls_handled = template_first_sheet[calls_handled_cell].value
    if(agent_id != None and calls_handled != None):
        template_first_sheet["H"+str(i)].value = nest_sales.count(agent_id)
        template_first_sheet["I"+str(i)].value = warranty_sales.count(agent_id)

for agent_id in nest_sales:
    if agent_id in jaelesiaTeam:
        jaelesiaNestSales += 1
        totalNestSales += 1
    if agent_id in tekTeam:
        tekNestSales += 1
        totalNestSales += 1
    if agent_id in antwonTeam:
        antwonNestSales += 1
        totalNestSales += 1

for agent_id in warranty_sales:
    if agent_id in jaelesiaTeam:
        jaelesiaDEPPsales += 1
        totalDEPPsales += 1
    if agent_id in tekTeam:
        tekDEPPsales += 1
        totalDEPPsales += 1
    if agent_id in antwonTeam:
        antwonDEPPsales += 1
        totalDEPPsales += 1
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#Open the FCP report from Sonar,
#Retrieve data and add to summary tab of template excel file
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
print("\nOpening fcp report......\n")

fcp_sales = get_fcp_sales(fcpReportLocation)

#Write out the FCP sales to the template
print("\nWriting out the FCP sales to the template.......\n")
for i in range(3, 50):
    agent_id_cell = "A"+str(i)
    agent_id = template_first_sheet[agent_id_cell].value
    calls_handled_cell = "C"+str(i)
    calls_handled = template_first_sheet[calls_handled_cell].value
    if(agent_id != None and calls_handled != None):
        template_first_sheet["G"+str(i)].value = fcp_sales.count(agent_id)

for agent_id in fcp_sales:
    if agent_id in jaelesiaTeam:
        jaelesiaFCPsales += 1
        totalFCPSales += 1
    if agent_id in tekTeam:
        tekFCPsales += 1
        totalFCPSales += 1
    if agent_id in antwonTeam:
        antwonFCPsales += 1
        totalFCPSales += 1

# print("jaelesiaFCPsales: ", jaelesiaFCPsales)
# print("tekFCPsales: ", tekFCPsales)
# print("antwonFCPsales: ", antwonFCPsales)
# print("totalFCPSales", totalFCPSales)
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#HIVE Sales: Open the Products report from Sonar,
#Retrieve data and add to summary tab of template excel file
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
hive_sales = get_hive_sales(hiveReportLocation)
print("\nWriting out the HIVE sales to the template.......\n")
for i in range(3, 50):
    agent_id_cell = "A"+str(i)
    agent_id = template_first_sheet[agent_id_cell].value
    calls_handled_cell = "C"+str(i)
    calls_handled = template_first_sheet[calls_handled_cell].value
    if(agent_id != None and calls_handled != None):
        template_first_sheet["j"+str(i)].value = fcp_sales.count(agent_id)

for agent_id in hive_sales:
    if agent_id in jaelesiaTeam:
        jaelesiaHiveSales += 1
        totalHiveSales += 1
    if agent_id in tekTeam:
        tekHiveSales += 1
        totalHiveSales += 1
    if agent_id in antwonTeam:
        antwonHiveSales += 1
        totalHiveSales += 1

#Write out the agent close rates to the template
for i in range(3,50):
    try:
        closeRate = (template_first_sheet["e" + str(i)].value + template_first_sheet["g" + str(i)].value) / template_first_sheet["d" + str(i)].value
        template_first_sheet["f" + str(i)].value = closeRate
    except:
        pass

    #Write out the supervisor and iQor totals to the template
    if template_first_sheet["b" + str(i)].value == "JAELESIA MOORE Total":
        template_first_sheet["c" + str(i)].value = jaelesiaTotalCallsHandled
        template_first_sheet["d" + str(i)].value = jaelesiaSalesCallsHandled
        template_first_sheet["e" + str(i)].value = jaelesiaTotalSales
        template_first_sheet["g" + str(i)].value = jaelesiaFCPsales
        template_first_sheet["h" + str(i)].value = jaelesiaNestSales
        template_first_sheet["i" + str(i)].value = jaelesiaDEPPsales
        template_first_sheet["j" + str(i)].value = jaelesiaHiveSales
        closeRate = (template_first_sheet["e" + str(i)].value + template_first_sheet["g" + str(i)].value) / template_first_sheet["d" + str(i)].value
        template_first_sheet["f" + str(i)].value = closeRate
    if template_first_sheet["b" + str(i)].value == "TEK LEVON Total":
        template_first_sheet["c" + str(i)].value = tekTotalCallsHandled
        template_first_sheet["d" + str(i)].value = tekSalesCallsHandled
        template_first_sheet["e" + str(i)].value = tekTotalSales
        template_first_sheet["g" + str(i)].value = tekFCPsales
        template_first_sheet["h" + str(i)].value = tekNestSales
        template_first_sheet["i" + str(i)].value = tekDEPPsales
        template_first_sheet["j" + str(i)].value = tekHiveSales
        closeRate = (template_first_sheet["e" + str(i)].value + template_first_sheet["g" + str(i)].value) / template_first_sheet["d" + str(i)].value
        template_first_sheet["f" + str(i)].value = closeRate
    if template_first_sheet["b" + str(i)].value == "ANTWON COLLINS Total":
        template_first_sheet["c" + str(i)].value = antwonTotalCallsHandled
        template_first_sheet["d" + str(i)].value = antwonSalesCallsHandled
        template_first_sheet["e" + str(i)].value = antwonTotalSales
        template_first_sheet["g" + str(i)].value = antwonFCPsales
        template_first_sheet["h" + str(i)].value = antwonNestSales
        template_first_sheet["i" + str(i)].value = antwonDEPPsales
        template_first_sheet["j" + str(i)].value = antwonHiveSales
        closeRate = (template_first_sheet["e" + str(i)].value + template_first_sheet["g" + str(i)].value) / template_first_sheet["d" + str(i)].value
        template_first_sheet["f" + str(i)].value = closeRate
    if template_first_sheet["b" + str(i)].value == "Grand Total":
        template_first_sheet["c" + str(i)].value = totalCallsHandled
        template_first_sheet["d" + str(i)].value = totalSalesCallsHandled
        template_first_sheet["e" + str(i)].value = totalSales
        template_first_sheet["g" + str(i)].value = totalFCPSales
        template_first_sheet["h" + str(i)].value = totalNestSales
        template_first_sheet["i" + str(i)].value = totalDEPPsales
        template_first_sheet["j" + str(i)].value = totalHiveSales
        closeRate = (template_first_sheet["e" + str(i)].value + template_first_sheet["g" + str(i)].value) / template_first_sheet["d" + str(i)].value
        template_first_sheet["f" + str(i)].value = closeRate
#-------------------------------------------------------------------------------
#Save the template as final.xlsx
#-------------------------------------------------------------------------------
print("\nSaving final template.......")
template.save("C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\final.xlsx")
print("\nDone.......\n")
