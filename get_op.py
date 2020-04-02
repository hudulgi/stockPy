import Creon
import pandas as pd
import datetime
import time
from dt_alimi import *


def date_split(line):
    date = line.split('-')
    date2 = datetime.date(int(date[0]), int(date[1]), int(date[2]))
    return date2


bot.sendMessage(myId, "시가 불러오기를 시작합니다.")

finishCode = 0

cr = Creon.Creon()

targetTable = pd.read_csv('C://Users//JHOO//iCloudDrive//stock//stock_data//target_list.csv')
code_list = targetTable['code'].to_list()
code_list2 = code_list[:]

ignoreItem = pd.read_csv('C://Users//JHOO//iCloudDrive//stock//stock_data//ignore_item.csv')
nowDate = datetime.datetime.now().date()
for item in ignoreItem.iterrows():
    if item[1]['code'] not in code_list:
        continue
    startDate = date_split(item[1]['start'])
    endDate = date_split(item[1]['end'])
    if nowDate >= startDate:
        if nowDate <= endDate:
            code_list2.remove(item[1]['code'])
            print('제외: %s %s' % (item[1]['code'], item[1]['name']))
            bot.sendMessage(myId, '제외: %s %s' % (item[1]['code'], item[1]['name']))

resultDf = pd.DataFrame(columns=['code', 'name', 'open'])

mrktStart = str(cr.get_start_time())
h = int(mrktStart[:-2])
m = int(mrktStart[-2:])
startTime = datetime.time(h, m, 0)
limitTime = datetime.time(h, m + 10, 0)

while code_list2:
    nowTime = datetime.datetime.now().time()
    if nowTime < startTime:
        print("장 시작 전")
        continue
    elif nowTime > limitTime:
        print("시간 제한으로 종료")
        finishCode = 2
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
        finishCode = 1

print(resultDf)
resultDf.to_csv('C://Users//JHOO//iCloudDrive//stock//stock_data//open_price//op_' + nowDate.strftime("%y%m%d") +
                '.csv', index=False)
bot.sendMessage(myId, "시가 불러오기를 완료하였습니다. 종료코드:%i" % finishCode)
