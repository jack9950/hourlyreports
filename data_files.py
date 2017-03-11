import time
from datetime import datetime

homeFolder = 'C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\'


#callsHandledReportLocation = homeFolder +'Bounce_Hourly_Sales_Report_03102017.xls'
#pogoSalesReportLocation = homeFolder + 'bounce_energy_iqor_report_18.xls'
# fcpReportLocation = homeFolder + 'HourlyProducts_Added.xls'
# DEPPreportLocation = homeFolder + 'products_sonar_03102017.xls'
# hiveNewServiceReportLocation = homeFolder + 'products_sonar_03102017.xls'
# hiveRenewalsReportLocation = homeFolder + 'hive_renewals_03102017.xls'

def pogoSalesReportLocation(*args):
    print('args: ', args == True)

    hour = time.localtime().tm_hour
    if (hour < 8 or hour > 21):
        currentHour = '21'
    else:
        currentHour = time.strftime('%H')

    pogoSalesReportLocation = homeFolder + 'bounce_energy_iqor_report_' + currentHour  + '.xls'
    return pogoSalesReportLocation

def callsHandledReportLocation(*args):
    if args[0]: #if a date is passed, use that to construct the file names (this assumes there is a folder with the date)
        currentDate = args[0]
        print('currentDate: ', currentDate)
    else:
        currentDate = datetime.now().strftime("%A %m-%d-%Y")

    callsHandledReportLocation = homeFolder +'Bounce_Hourly_Sales_Report_' + currentDate + '.xls'

    return callsHandledReportLocation

def fcpReportLocation():
    fcpReportLocation = homeFolder + 'HourlyProducts_Added.xls'
    return fcpReportLocation

def DEPPreportLocation():
    DEPPreportLocation = homeFolder + 'products_sonar_03102017.xls'
    return DEPPreportLocation

def hiveNewServiceReportLocation():
    hiveNewServiceReportLocation = homeFolder + 'products_sonar_03102017.xls'
    return hiveNewServiceReportLocation

def hiveRenewalsReportLocation():
    hiveRenewalsReportLocation = homeFolder + 'hive_renewals_03102017.xls'
    return hiveRenewalsReportLocation
