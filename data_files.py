import time
from datetime import datetime

homeFolder = 'C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\'

jaelesiaTeam = [2062062, 2062053,2062054,2062090,2062026,2062011,2062105,2062015,2062039,2062024,2062020,2062094]
tekTeam = [2062007, 2062103, 2062052, 2062111, 2062010, 2062057, 2062048, 2062076, 2062098, 2062107, 2062104, 2062100]
antwonTeam = [2062039, 2062018, 2062110, 2062101, 2062096, 2062102, 2062044, 2062109, 2062003, 2062074, 2062058, 2062049, 2062095]
jacksonTeam = [2062067, 2062001]
trainingTeam =[]

tableNames = [[2062062,'BROWN, ADRIANE'],
[2062053,'AGUILAR, BETTY'],
[2062054,'ROBINSON, CARRIE'],
[2062090,'BOOTH, DEVONAE'],
[2062026,'BECERRA, DOLORES'],
[2062011,'JONES, GRACE'],
[2062105,'MOORE, MARQUIS'],
[2062015,'GREEN, REISHA'],
[2062039,'CUELLAR, REYNA'],
[2062024,'MALONE, SHEMEKA'],
[2062020,'GABRIEL, TABITHA'],
[2062094,'JACKSON, WESLEY'],
['jaelesia','JAELESIA MOORE Total'],
[2062007,'ERVIN, ANGELIQUE'],
[2062103,'FLUDD, JESSICA'],
[2062052,'LADAY, JESSICA'],
[2062111,'BURKES, KENEISHA'],
[2062010,'HERRERA, MAGDALY'],
[2062057,'MURPHY, NATASCHA'],
[2062048,'HENRIQUES, PATRICK'],
[2062076,'RHODES, PEGGY'],
[2062098,'SWAYZER, SHERMEKA'],
[2062107,'SIMMONS, TASHARI'],
[2062104,'GOUAUX, TAYLOR'],
[2062100,'BENJAMIN, TRACEY'],
['tek','TEK LEVON Total'],
[2062018,'MCMURRIN, ANDREADIS'],
[2062110,'SUTTON, DANA'],
[2062101,'BROWN, KEYUNNA'],
[2062096,'MCZEAL, LATARVEYA'],
[1234567,'COMBS, LATOYA'],
[2062044,'WILLIAMS, MARCUS'],
[2062003,'WILLIAMS, PAMELA'],
[2062074,'IGLESIAS, RAY'],
[2062058,'LASTER, SHAWANDA'],
[2062049,'REDD, TAMERIA'],
[2062095,'LEWIS, THORENT'],
['antwon','ANTWON COLLINS Total'],
[2062067,'CARTWRIGHT, GERISHA'],
[2062001,'NICKERSON, JACQUELINE'],
['jackson','JACKSON NDIHO Total'],
['grandTotal', 'GRAND TOTAL']
]

              
tableNames2 = [[2062053, 'AGUILAR, BETTY'],
              [2062026, 'BECERRA, DOLORES'],
              [2062062, 'BROWN, ADRIANE'],
              [2062090, 'BOOTH, DEVONAE'],
              [2062020, 'GABRIEL, TABITHA'],
              [2062048, 'HENRIQUES, PATRICK'],
              [2062094, 'JACKSON, WESLEY'],
              [2062011, 'JONES, GRACE'],
              [2062057, 'MURPHY, NATASCHA'],
              [2062054, 'ROBINSON, CARRIE'],
              [2062001, 'NICKERSON, JACQUELINE'],
              ['jaelesia', 'JAELESIA MOORE Total'],
              [2062067, 'CARTWRIGHT, GERISHA'],
              [2062015, 'GREEN, REISHA'],
              [2062010, 'HERRERA, MAGDALY'],
              [2062024, 'MALONE, SHEMEKA'],
              [2062098, 'SWAYZER, SHERMEKA'],
              [2062007, 'ERVIN, ANGELIQUE'],
              ['tek', 'TEK LEVON Total'],
              [2062039, 'CUELLAR, REYNA'],
              [2062074, 'IGLESIAS, RAY'],
              [2062052, 'LADAY, JESSICA'],
              [2062058, 'LASTER, SHAWANDA'],
              [2062095, 'LEWIS, THORENT'],
              [2062018, 'MCMURRIN, ANDREADIS'],
              [2062096, 'MCZEAL, LATARVEYA'],
              [2062049, 'REDD, TAMERIA'],
              [2062076, 'RHODES, PEGGY'],
              [2062044, 'WILLIAMS, MARCUS'],
              [2062003, 'WILLIAMS, PAMELA'],
              [2062066, 'MURPH, DOMINIQUE'],
              ['antwon', 'ANTWON COLLINS Total'],
              [2062100, "BENJAMIN, TRACEY"],
              [2062111, "BURKES, KENEISHA"],
              [2062101, "BROWN, KEYUNNA"],
              [2062102, "COMBS, LATOYA"],
              [2062103, "FLUDD, JESSICA"],
              [2062104, "GOUAUX, TAYLOR"],
              [2062105, "MOORE, MARQUIS"],
              [2062107, "SIMMONS, TASHARI"],
              [2062108, "SMITH, CONSTANCE"],
              [2062109, "SMITH, MICHAEL"],
              [2062088, "SMITH, VICTORIA"],
              [2062110, "SUTTON, DANA"],
              ['jackson', 'JACKSON NDIHO Total'],
              ['grandTotal', 'Grand Total']]




tableNamesJackson = [[2062100, "BENJAMIN, TRACEY"],
                     [2062111, "BURKES, KENEISHA"],
                     [2062101, "BROWN, KEYUNNA"],
                     [2062102, "COMBS, LATOYA"],
                     [2062103, "FLUDD, JESSICA"],
                     [2062104, "GOUAUX, TAYLOR"],
                     [2062105, "MOORE, MARQUIS"],
                     [2062107, "SIMMONS, TASHARI"],
                     [2062108, "SMITH, CONSTANCE"],
                     [2062109, "SMITH, MICHAEL"],
                     [2062088, "SMITH, VICTORIA"],
                     [2062110, "SUTTON, DANA"],
                     ['jackson', 'JACKSON NDIHO Total']]
              

termedAgents = [[2062112, "CAPERS, LATRYSTA"],
                [2062099, 'SHAW, TRISTAN']
                ]


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
        print('pogoSalesReportLocation: ', pogoSalesReportLocation)
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
