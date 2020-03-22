import Creon
import pandas as pd
import time

cr = Creon.Creon()
targetTable = pd.read_csv('C://Users//JHOO//iCloudDrive//stock//stock_data//target_list.csv')

code_list = targetTable['code']
openPrice = []
name = []

for code in code_list:
    print(code)
    openPrice.append(cr.getOpen(code))
    name.append(cr.getName(code))
    time.sleep(0.2)

targetTable['open'] = openPrice
targetTable['name'] = name

targetTable.to_csv('C://Users//JHOO//iCloudDrive//stock//stock_data//target_list2.csv', index = False)