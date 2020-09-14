from sqlalchemy import create_engine
import pymysql
import pandas as pd
pymysql.install_as_MySQLdb()


class StockDB:
    def __init__(self, uid, password, host, db, port=3307):
        self.engine = create_engine(f"mysql+mysqldb://{uid}:{password}@{host}:{port}/{db}", encoding='utf-8')
        self.conn = pymysql.connect(host=host,
                                    port=port,
                                    user=uid,
                                    password=password,
                                    db=db,
                                    charset='utf8')
        self.cursor = self.conn.cursor()

    def close(self):
        self.conn.close()

    def select_max_date(self, table_name):
        sql = 'select max(Date) from ' + table_name
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        return result[0]

    def insert_chart(self, data, table_name):
        data.to_sql(name=table_name, con=self.engine, if_exists='append')
        self.conn.commit()

    def create_table(self, table_name):
        sql = 'SHOW TABLES LIKE \'' + table_name + '\''
        result = self.cursor.execute(sql)
        if not result:
            sql = 'create table ' + table_name + '(Date date primary key,Open Decimal,High Decimal,Low Decimal,' \
                                                 'Close Decimal, Volume Decimal);'
            self.cursor.execute(sql)
            self.conn.commit()

    def read_table(self, table_name):
        df = pd.read_sql_table(table_name, self.engine)
        return df
