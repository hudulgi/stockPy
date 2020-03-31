import Creon
from datetime import date, timedelta, datetime
import time
from sqlalchemy import create_engine
from DB import *
import pymysql
pymysql.install_as_MySQLdb()


class StockDB:
    def __init__(self):
        sql_eng = "mysql+mysqldb://%s:%s@%s:%s/%s" % (DBInfo['user'], DBInfo['password'], DBInfo['host'],
                                                      DBInfo['port'], DBInfo['db'])
        self.engine = create_engine(sql_eng, encoding=DBInfo['charset'])
        self.conn = pymysql.connect(host=DBInfo['host'],
                                    port=int(DBInfo['port']),
                                    user=DBInfo['user'],
                                    password=DBInfo['password'],
                                    db=DBInfo['db'],
                                    charset=DBInfo['charset'])
        self.cursor = self.conn.cursor()

    def close(self):
        self.conn.close()

    def select_max_date(self,table_name):
        sql = 'select max(Date) from ' + table_name
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        return result[0]

    def insert_chart(self,data, table_name):
        data.to_sql(name=table_name, con=self.engine, if_exists='append', index_label='Date')
        self.conn.commit()

    def create_table(self,table_name):
        sql = 'SHOW TABLES LIKE \'' + table_name + '\''
        result = self.cursor.execute(sql)
        if not result:
            sql = 'create table ' + table_name + \
                  '(Date date primary key,Open Decimal,High Decimal,Low Decimal,Close Decimal, Volume Decimal);'
            self.cursor.execute(sql)
            self.conn.commit()


cr = Creon.Creon()

# 현재 날짜, 시간 받아오기
now_hour = datetime.now().hour
now_day = date.today()

# DB 접속 및 초기화
myDb = StockDB()

# 종목 리스트 받아오기
# 1 코스피, 2 코스닥
code_kospi = cr.requestCode(1)
code_kosdaq = cr.requestCode(2)
codelist = code_kospi + code_kosdaq

# 연속실행
for code in codelist:
    # 테이블이 생성되지 않았으면 테이블 생성
    myDb.create_table(code)

    # 테이블에 입력된 데이터 중 가장 최근 날짜 획득
    recent_day = myDb.select_max_date(code)
    dailyData = []

    # 주가 데이터 시작날짜 선정
    if recent_day is not None:

        # 시작날짜 정의: DB의 마지막날짜 + 1
        startDate = recent_day + timedelta(1)

        if startDate == now_day:
            print(3)
            print('%s 패스' % code)
            continue
        elif recent_day == now_day:
            print(4)
            print('%s 패스' % code)
            continue

        print(2)
        startDate = startDate.strftime('%Y%m%d')

    else:
        # 기존 데이터가 없는 경우 초기 설정된 시작날짜로 데이터 받아옴
        print(1)
        startDate = '20080101'

    # 종료날짜 정의: 오늘에서 하루 이전
    endDate = now_day - timedelta(1)
    endDate = endDate.strftime('%Y%m%d')

    dailyData = cr.requestData(code, startDate, endDate)

    #if dailyData.empty is False:
    #    if recent_day.strftime('%Y%m%d') == str(dailyData.index.values[0]):
    #        print(5)
    #        print('%s 패스' % code)
    #        continue

    # 주가 데이터 DB에 전송
    myDb.insert_chart(dailyData, code)

    print('%s 전송완료. %s - %s\n' % (code, startDate, endDate))
    time.sleep(0.2)

# DB 접속 종료
myDb.close()
