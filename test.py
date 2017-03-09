from get_HIVE_renewals import get_HIVE_renewals

homeFolder = 'C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\'

callsHandledReportLocation = homeFolder +'Bounce_Hourly_Sales_Report_03082017.xls'
pogoSalesReportLocation = homeFolder + 'bounce_energy_iqor_report_18.xls'
fcpReportLocation = homeFolder + 'HourlyProducts_Added.xls'
#DEPPreportLocation = homeFolder + 'BounceEnergyProducts Added2017-03-08.xls'
DEPPreportLocation = homeFolder + 'products_sonar_03082017.xls'
hiveNewServiceReportLocation = homeFolder + 'products_sonar_03082017.xls'
hiveRenewalsReportLocation = homeFolder + 'hive_renewals_03082017.xls'

result = get_HIVE_renewals(hiveRenewalsReportLocation)

for item in result:
    print(item)
