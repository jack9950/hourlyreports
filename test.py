from get_nest_sales import get_nest_sales

result = get_nest_sales('C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\products_report_030317.xls')

for item in result:
    print(item)
