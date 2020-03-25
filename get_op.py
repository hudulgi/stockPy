import Creon
import pandas as pd
import datetime
import time

cr = Creon.Creon()
targetTable = pd.read_csv('C://Users//JHOO//iCloudDrive//stock//stock_data//target_list.csv')

code_list = targetTable['code'].to_list()
code_list2 = code_list[:]

resultDf = pd.DataFrame(columns=['code', 'name', 'open'])

mrktStart = str(cr.get_start_time())
h = int(mrktStart[:-2])
m = int(mrktStart[-2:])
startTime = datetime.time(h, m, 0)
limitTime = datetime.time(h, m + 30, 0)

while code_list2:
    now = datetime.datetime.now().time()
    if now < startTime:
        print("장 시작 전")
        continue
    elif now > limitTime:
        print("시간 제한으로 종료")
        break

    for code in code_list:
        if code not in code_list2:
            continue
        op, name, marketFlag, state = cr.get_open(code)
        print(code, name, op, marketFlag, state)

        if state == 'Y':
            code_list2.remove(code)
        elif marketFlag == '2':
            temp = pd.Series([code, name, op], index=['code', 'name', 'open'])
            resultDf = resultDf.append(temp, ignore_index=True)
            code_list2.remove(code)

        time.sleep(0.3)

print(resultDf)
resultDf.to_csv('C://Users//JHOO//iCloudDrive//stock//stock_data//target_open.csv', index=False)
