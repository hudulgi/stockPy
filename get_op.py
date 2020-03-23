import Creon
import pandas as pd
import time

cr = Creon.Creon()
targetTable = pd.read_csv('C://Users//JHOO//iCloudDrive//stock//stock_data//target_list.csv')

code_list = targetTable['code'].to_list()
code_list2 = code_list[:]

resultDf = pd.DataFrame(columns=['code', 'name', 'open'])

while code_list2:
    for code in code_list:
        if code not in code_list2:
            continue
        op, name, marketFlag = cr.get_open(code)
        print(code, name, op, marketFlag)

        if marketFlag == '2':
            temp = pd.Series([code, name, op], index=['code', 'name', 'open'])
            resultDf = resultDf.append(temp, ignore_index=True)
            code_list2.remove(code)
        elif marketFlag == '4':
            code_list2.remove(code)

        time.sleep(0.3)

print(resultDf)
#targetTable.to_csv('C://Users//JHOO//iCloudDrive//stock//stock_data//target_list2.csv', index = False)
