import time
from datetime import datetime
currentDate = datetime.now().strftime("%m%d%Y")
currentTime = time.strftime("%I%M%S")
print(currentDate + "_" + currentTime)








# # from get_HIVE_renewals import get_HIVE_renewals
# #
# # homeFolder = 'C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\'
# #
# # callsHandledReportLocation = homeFolder +'Bounce_Hourly_Sales_Report_03082017.xls'
# # pogoSalesReportLocation = homeFolder + 'bounce_energy_iqor_report_18.xls'
# # fcpReportLocation = homeFolder + 'HourlyProducts_Added.xls'
# # #DEPPreportLocation = homeFolder + 'BounceEnergyProducts Added2017-03-08.xls'
# # DEPPreportLocation = homeFolder + 'products_sonar_03082017.xls'
# # hiveNewServiceReportLocation = homeFolder + 'products_sonar_03082017.xls'
# # hiveRenewalsReportLocation = homeFolder + 'hive_renewals_03082017.xls'
#
# # result = get_HIVE_renewals(hiveRenewalsReportLocation)
# #
# # for item in result:
# #     print(item)
#
# import openpyxl
# from openpyxl import Workbook
# from openpyxl.styles import Style, Font, Border, Side, Fill, PatternFill
# from get_agent_ids_and_calls import get_agent_ids_and_calls
# from get_pogo_sales import get_pogo_sales
# #from get_nest_sales import get_nest_sales
# from get_DEPP_sales import get_DEPP_sales
# from get_fcp_sales import get_fcp_sales
# from get_HIVE_new_service import get_HIVE_new_service
# from get_HIVE_renewals import get_HIVE_renewals
#
# wb = Workbook()
# ws = wb.active
#
# below_goal_text = "9C0006"
# below_goal_bg = "FFC7CE"
#
# close_to_goal_text = "9C6500"
# close_to_goal_bg = "FFEB9C"
#
# at_or_above_goal_text = "006100"
# at_or_above_goal_bg = "C6EFCE"
#
# myCell = ws['B5']
# myCell.value = "My Cell"
# myCell.font = Font(name='Calibri', size=11, bold=True)
# myCell.fill = PatternFill("solid", fgColor=agent_below_goal)
# # agent_below_goal = Style()
# # agent_below_goal.fill = PatternFill("solid", fgColor="ff6666")
#
# myCell.style = agent_below_goal
#print(agent_below_goal)

#wb.save("test.xlsx")
