import Creon
import pandas as pd
import datetime
import time


def date_split(line):
    date = line.split('-')
    date2 = datetime.date(int(date[0]), int(date[1]), int(date[2]))
    return date2


def import_target_code(_name):
    _targetTable = pd.read_csv(data_path + '\\' + _name)
    _code_list = _targetTable['code'].to_list()
    return _code_list


cr = Creon.Creon()

base_path = "C:\\CloudStation\\dt_data"
data_path = base_path + "\\target"
open_path = base_path + "\\daily_data\\open_price"

target_name = ["target_list_2021.csv", "target_ETF.csv"]

code_list_tot = []
for tg in target_name:
    code_list_tot += import_target_code(tg)

nowDate = datetime.datetime.now().date()
resultDf = pd.DataFrame(columns=['code', 'name', 'open'])

for code in code_list_tot:
    op, name, marketFlag, state = cr.get_open(code)
    print(code, name, op, marketFlag, state)

    if state == 'Y':
        continue
    elif marketFlag == '2':
        temp = pd.Series([code, name, op], index=['code', 'name', 'open'])
        resultDf = resultDf.append(temp, ignore_index=True)

    time.sleep(0.3)


print(resultDf)
resultDf.to_csv(open_path + '\\op_' + nowDate.strftime("%y%m%d") + '_report.csv', index=False)
