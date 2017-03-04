import openpyxl
from get_agent_ids_and_calls import get_agent_ids_and_calls
from get_pogo_sales import get_pogo_sales
from get_nest_sales import get_nest_sales
from get_warranty_sales import get_warranty_sales

agentIDs = [2062004, 2062026, 2062043, 2062034, 2062053, 2062048, 2062042,
            2062011, 2062030, 2062045, 2062046, 2062016, 2062001, 2062036,
            2062039, 2062025, 2062041, 2062052, 2062037, 2062024, 2062049,
            2062031, 2062044, 2062003, 2062028, 2062022, 2062051, 2062021,
            2062035, 2062007, 2062020, 2062015, 2062040, 2062010, 2062018,
            2062054, 2062032, 2062033, 2062062, 2062070, 2062067, 2062058,
            2062056, 2062066, 2062057, 2062065, 2062060]

supervisorIDs = {"aervin":2062007, "jnickerson":2062001, "tlevon": 2062007,
                 "jacksonn": 2062047, "jabram":2062017, "iqr_acollins":2062072,
                 "jmoore":2062023, "mayala":2062002}

callsHandledReportLocation = 'C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\Bounce_Hourly_Sales_Report_03032017.xls'
pogoSalesReportLocation = 'C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\bounce_energy_iqor_report_21.xls'
productsReportLocation = 'C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\BounceEnergyProducts\ Added2017-03-03.xls'
fcpReportLocation = 'C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\HourlyProducts_Added.xls'

#Open the template file for editing:
print("\nOpening template file for editing......\n")

template = openpyxl.load_workbook("C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\final.xlsx")
template_sheets = template.get_sheet_names()
template_first_sheet = template.get_sheet_by_name(template_sheets[0])


#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#Open the Bounce Hourly Sales Report from iQor,
#Retrieve data and add to summary tab of template excel file
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------

#Open the calls handled report from iQor
print("\nOpening the calls handled report from iQor........\n")

#Gather up all the agent IDs and call counts:
    #Column E: This is the Agent ID
    #Column L: This is the Total Calls Handled
    #Column AV: This is the Sales Calls Handled
print("\nReading agent IDs and call counts.......\n")

calls_handled = get_agent_ids_and_calls(callsHandledReportLocation)

#Write out the call counts to the template file
print("\nWriting call counts to the template file.......\n")
for i in range(3, 60):
    agent_id_cell = "A"+str(i)
    agent_id = template_first_sheet[agent_id_cell].value
    #retrieve each agent ID from the template and check if it is the list.
    #if found write the calls handled and sales calls handled to the template file
    if(agent_id != None):
        for item in calls_handled: #check each nested list
            if agent_id in item and item[1] != 0: #if found and total calls > 0,
                                                  #write calls data to the template
                template_first_sheet["C"+str(i)].value = item[1] #Total Calls Handled
                template_first_sheet["D"+str(i)].value = item[2] #Sales Calls Handled


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
    if(agent_id != None):
        template_first_sheet["E"+str(i)].value = pogo_sales.count(agent_id)
        #print(agent_id, ": ", pogo_sales.count(agent_id))



#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#Retrieve NEST and DEPP data and add to summary tab of template excel file
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------

print("\nGathering NEST and warranty sales......\n")

#Get the NEST sales - the returned calue (nest_sales) is a list of agent IDs
#Every agent ID in the list is an agent ID that has a NEST sale
nest_sales = get_nest_sales('C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\BounceEnergyProducts Added2017-03-03.xls')

#Get the Warranty sales - the returned calue (warranty_sales) is a list of agent IDs
#Every agent ID in the list is an agent ID that has a Warranty sale
warranty_sales = get_warranty_sales('C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\BounceEnergyProducts Added2017-03-03.xls')

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


#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#Open the FCP report from Sonar,
#Retrieve data and add to summary tab of template excel file
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
print("\nOpening fcp report......\n")
wb = openpyxl.load_workbook('C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\fcp_report_03032017.xlsx')
sheets = wb.get_sheet_names()
sheet = wb.get_sheet_by_name(sheets[0])

fcp_sales = []
for i in range(1,2000):
    agent_id_cell = "BJ" + str(i)
    agent_id = sheet[agent_id_cell].value
    if(agent_id != None and type(agent_id) == int and agent_id > 2000000): #FCP Sales
        fcp_sales.append(agent_id)

#Write out the FCP sales to the template
print("\nWriting out the FCP sales to the template.......\n")
for i in range(3, 50):
    agent_id_cell = "A"+str(i)
    agent_id = template_first_sheet[agent_id_cell].value
    calls_handled_cell = "C"+str(i)
    calls_handled = template_first_sheet[calls_handled_cell].value
    if(agent_id != None and calls_handled != None):
        template_first_sheet["G"+str(i)].value = fcp_sales.count(agent_id)


#-------------------------------------------------------------------------------
#Save the template as final.xlsx
#-------------------------------------------------------------------------------
print("\nSaving final template.......")
template.save("C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\final.xlsx")
print("\nDone.......\n")
