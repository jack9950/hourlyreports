from get_hive_sales import get_hive_sales

homeFolder = 'C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\'

callsHandledReportLocation = homeFolder +'Bounce_Hourly_Sales_Report_03062017.xls'
pogoSalesReportLocation = homeFolder + 'bounce_energy_iqor_report_18.xls'
productsReportLocation = homeFolder + 'BounceEnergyProducts Added2017-03-06.xls'
fcpReportLocation = homeFolder + 'HourlyProducts_Added.xls'
hiveReportLocation = homeFolder + 'products_sonar_03062017.xls'

result = get_hive_sales(hiveReportLocation)

for item in result:
    print(item)
