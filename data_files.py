import time
from datetime import datetime

homeFolder = 'C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\'

jaelesiaTeam = [2062026, 2062062, 2062053, 2062020, 2062048, 2062011,
                2062057, 2062054, 2062090, 2062082, 2062083, 2062084, 2062001]
tekTeam = [2062067, 2062035, 2062015, 2062040, 2062010, 2062024, 2062065,
           2062060, 2062007]
antwonTeam = [2062039, 2062073, 2062074, 2062052, 2062058, 2062018, 2062049,
              2062076, 2062031, 2062044, 2062003, 2062089, 2062066]
jacksonTeam = [2062087, 2062094, 2062095, 2062096, 2062098, 2062099]

tableNames = [[2062026, 'BECERRA, DOLORES'],
              [2062090, 'BOOTH, DEVONAE'],
              [2062062, 'BROWN, ADRIANE'],
              [2062053, 'COLBERT, BETTY'],
              [2062020, 'GABRIEL, TABITHA'],
              [2062082, 'HAYMES, MONICA'],
              [2062048, 'HENRIQUES, PATRICK'],
              [2062083, 'HOUSTON, KAMESHA'],
              [2062084, 'JONES, BROOKE'],
              [2062011, 'JONES, GRACE'],
              [2062057, 'MURPHY, NATASCHA'],
              [2062054, 'ROBINSON, CARRIE'],
              [2062001, 'NICKERSON, JACQUELINE'],
              ['jaelesia', 'JAELESIA MOORE Total'],
              [2062067, 'CARTWRIGHT, GERISHA'],
              [2062015, 'GREEN, REISHA'],
              [2062040, 'HARRIS, SHAMANDA'],
              [2062010, 'HERRERA, MAGDALY'],
              [2062024, 'MALONE, SHEMEKA'],
              [2062007, 'ERVIN, ANGELIQUE'],
              ['tek', 'TEK LEVON Total'],
              [2062039, 'CUELLAR, REYNA'],
              [2062073, 'GUZMAN, HENRY'],
              [2062074, 'IGLESIAS, RAY'],
              [2062052, 'LADAY, JESSICA'],
              [2062058, 'LASTER, SHAWANDA'],
              [2062018, 'MCMURRIN, ANDREADIS'],
              [2062049, 'REDD, TAMERIA'],
              [2062076, 'RHODES, PEGGY'],
              [2062031, 'SLEDGE, DEBRA'],
              [2062089, 'THORNE, ALICE'],
              [2062044, 'WILLIAMS, MARCUS'],
              [2062003, 'WILLIAMS, PAMELA'],
              [2062066, 'MURPH, DOMINIQUE'],
              ['antwon', 'ANTWON COLLINS Total'],
              [2062087,  'JOHN SAMPSON'],
              [2062094, 'WESLEY JACKSON'],
              [2062095, 'THORENT LEWIS'],
              [2062096, 'LATARVEYA MCZEAL'],
              [2062098, 'SHERMEKA SWAYZER'],
              [2062099, 'TRISTAN SHAW'],
              ['jackson', 'JACKSON NDIHO Total'],
              ['grandTotal', 'Grand Total']]


def callsHandledReportLocation(*args):
    if args[0]:
        # if a date, "MTD" or "WTD" is passed,
        # use that to construct the file names
        currentDate = args[0]
        if currentDate == "MTD":
            callsHandledReportLocation = homeFolder + currentDate + '\\Bounce_Engery_Agent_Performance_Rollup.xls'
        else:
            callsHandledReportLocation = homeFolder + currentDate + \
                '\\Bounce_Hourly_Sales_Report_' + currentDate + '.xls'
        # print('currentDate: ', currentDate)
    else:  # No args were passed
        currentDate = datetime.now().strftime("%m%d%Y")
        callsHandledReportLocation = homeFolder + 'Bounce_Hourly_Sales_Report_' + currentDate + '.xls'

    return callsHandledReportLocation


def pogoSalesReportLocation(*args):
    if args[0]:  # if a date, "MTD" or "WTD" is passed, use that to construct the file names
        currentDate = args[0]
        if args[0] == "MTD":
            pogoSalesReportLocation = homeFolder + currentDate + '\\NOPR.xls'
        else:
            currentHour = '21'
            currentDate = args[0]
            pogoSalesReportLocation = homeFolder + currentDate + \
                '\\bounce_energy_iqor_report_' + currentHour + '.xls'
    else:  # No args were passed
        hour = time.localtime().tm_hour
        if (hour < 8 or hour > 21):
            currentHour = '21'
        else:
            currentHour = time.strftime('%H')
        pogoSalesReportLocation = homeFolder + 'bounce_energy_iqor_report_' + currentHour + '.xls'

    return pogoSalesReportLocation


def fcpReportLocation(*args):
    if args[0]:
        currentDate = args[0]
        if args[0] == "MTD":
            fcpReportLocation = homeFolder + currentDate + '\\FCP.xls'
        else:
            currentDate = args[0]
            fcpReportLocation = homeFolder + currentDate + '\\HourlyProducts_Added.xls'
    else:
        fcpReportLocation = homeFolder + 'HourlyProducts_Added.xls'

    return fcpReportLocation


def DEPPreportLocation(*args):  # if a date is passed, use that to construct the file names

    if args[0]:  # if a date is passed, use that to construct the file names
        currentDate = args[0]
        if args[0] == "MTD":
            DEPPreportLocation = homeFolder + currentDate + '\\products_sonar.xls'
        else:
            currentDate = args[0]
            DEPPreportLocation = homeFolder + currentDate + '\\products_sonar_' + currentDate + '.xls'
            # print('currentDate: ', currentDate)
    else:
        currentDate = datetime.now().strftime("%m%d%Y")
        DEPPreportLocation = homeFolder + 'products_sonar_' + currentDate + '.xls'

    return DEPPreportLocation


# if a date is passed, use that to construct the file names
def hiveNewServiceReportLocation(*args):

    if args[0]:  # if a date is passed, use that to construct the file names
        currentDate = args[0]
        if args[0] == "MTD":
            hiveNewServiceReportLocation = homeFolder + currentDate + '\\products_sonar.xls'
        else:
            currentDate = args[0]
            hiveNewServiceReportLocation = homeFolder + currentDate + '\\products_sonar_' + currentDate + '.xls'
            # print('currentDate: ', currentDate)
    else:
        currentDate = datetime.now().strftime("%m%d%Y")
        hiveNewServiceReportLocation = homeFolder + 'products_sonar_' + currentDate + '.xls'

    return hiveNewServiceReportLocation


def hiveRenewalsReportLocation(*args):  # if a date is passed, use that to construct the file names

    if args[0]:  # if a date is passed, use that to construct the file names
        currentDate = args[0]
        if args[0] == "MTD":
            hiveRenewalsReportLocation = homeFolder + currentDate + '\\hive_renewals.xls'
        else:
            currentDate = args[0]
            hiveRenewalsReportLocation = homeFolder + currentDate + '\\hive_renewals_' + currentDate + '.xls'
            # print('currentDate: ', currentDate)
    else:
        currentDate = datetime.now().strftime("%m%d%Y")
        hiveRenewalsReportLocation = homeFolder + 'hive_renewals_' + currentDate + '.xls'

    return hiveRenewalsReportLocation
