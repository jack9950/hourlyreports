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
    if args[0]: #if a date is passed, use that to construct the file names
        currentHour = '21'
        currentDate = args[0]
        pogoSalesReportLocation = homeFolder + currentDate + '\\bounce_energy_iqor_report_' + currentHour  + '.xls'
    else:
        hour = time.localtime().tm_hour
        if (hour < 8 or hour > 21):
            currentHour = '21'
        else:
            currentHour = time.strftime('%H')
        pogoSalesReportLocation = homeFolder + 'bounce_energy_iqor_report_' + currentHour  + '.xls'

    return pogoSalesReportLocation

def callsHandledReportLocation(*args):
    if args[0]: #if a date is passed, use that to construct the file names
        currentDate = args[0]
        callsHandledReportLocation = homeFolder + currentDate + '\\Bounce_Hourly_Sales_Report_' + currentDate + '.xls'
        # print('currentDate: ', currentDate)
    else:
        currentDate = datetime.now().strftime("%m%d%Y")
        callsHandledReportLocation = homeFolder +'Bounce_Hourly_Sales_Report_' + currentDate + '.xls'

    return callsHandledReportLocation

def fcpReportLocation(*args):
    if args[0]:
        currentDate = args[0]
        fcpReportLocation = homeFolder + currentDate + '\\HourlyProducts_Added.xls'
    else:
        fcpReportLocation = homeFolder + 'HourlyProducts_Added.xls'

    return fcpReportLocation

def DEPPreportLocation(*args): #if a date is passed, use that to construct the file names

    if args[0]: #if a date is passed, use that to construct the file names
        currentDate = args[0]
        DEPPreportLocation = homeFolder + currentDate + '\\products_sonar_' + currentDate + '.xls'
        # print('currentDate: ', currentDate)
    else:
        currentDate = datetime.now().strftime("%m%d%Y")
        DEPPreportLocation = homeFolder + 'products_sonar_' + currentDate + '.xls'

    return DEPPreportLocation

def hiveNewServiceReportLocation(*args): #if a date is passed, use that to construct the file names

    if args[0]: #if a date is passed, use that to construct the file names
        currentDate = args[0]
        hiveNewServiceReportLocation = homeFolder + currentDate + '\\products_sonar_' + currentDate + '.xls'
        # print('currentDate: ', currentDate)
    else:
        currentDate = datetime.now().strftime("%m%d%Y")
        hiveNewServiceReportLocation = homeFolder + 'products_sonar_' + currentDate + '.xls'

    return hiveNewServiceReportLocation

def hiveRenewalsReportLocation(*args): #if a date is passed, use that to construct the file names

    if args[0]: #if a date is passed, use that to construct the file names
        currentDate = args[0]
        hiveRenewalsReportLocation = homeFolder + currentDate + '\\hive_renewals_' + currentDate + '.xls'
        # print('currentDate: ', currentDate)
    else:
        currentDate = datetime.now().strftime("%m%d%Y")
        hiveRenewalsReportLocation = homeFolder + 'hive_renewals_' + currentDate + '.xls'

    return hiveRenewalsReportLocation
