from get_warranty_sales import get_warranty_sales

result = get_warranty_sales('C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\BounceEnergyProducts Added2017-03-03.xls')

for item in result:
    print(item)
