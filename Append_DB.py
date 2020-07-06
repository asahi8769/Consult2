import sqlite3
import pandas as pd
import os
import warnings
from utils.open_zip import open_zipfile

warnings.filterwarnings('ignore')


class MainDB:
    def __init__(self, issue_month):
        self.issue_month = str(issue_month)
        self.year = str(issue_month)[0:4]
        self.month = str(issue_month)[4:6]
        self.issue_months = [self.year + '0' * (2 - len(str(month))) + str(month) for month in
                             [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]]
        self.db_address = os.path.join('Main_DB', 'Main_DB.db')
        self.excel_address = os.path.join('Cookies', self.issue_month + '.xlsx')

    def df_base(self, sheet_name):
        if self.issue_month + '.xlsx' in os.listdir ('Cookies'):
            with open (self.excel_address, 'rb') as file:
                return pd.read_excel (file, sheet_name=sheet_name, skiprows=0)
        else:
            print('No {} in the folder.'.format(self.issue_month + '.xlsx'))
            return None

    @staticmethod
    def df_process(df, columns_to_keep):
        df = df[columns_to_keep]
        df.fillna('', inplace=True)
        df = df.applymap (str)
        return df.apply (pd.to_numeric, errors='ignore')

    def std_cur(self):
        df_std = self.df_base('기준')
        columns_to_keep = ['보상합계', '변제합계', 'C 통화', 'V 통화']
        return self.df_process(df_std, columns_to_keep)

    def exc_cur(self):
        df_exc = self.df_base('환산')
        columns_to_keep = ['고객사', '사정년월', 'No', 'E/D', '통보서','LIST', 'V List No', 'RO-NO', 'OP-CODE', '클레임상태', '원인부품번호',
                           '원인부품명', '업체코드', '업체명', 'VIN-NO', '엔진NO', 'TM NO', '차종모델코드', '사용년수',
                           '주행거리', 'C 분담율', 'C적용환율', '보상합계', '변제대상 합계', 'V 분담율', 'V적용환율',
                           '변제합계']
        return self.df_process(df_exc, columns_to_keep)

    def merge(self):
        df_std = self.std_cur()
        print('Converting {} 1/2...'.format(self.issue_month + '.xlsx'))
        df_exc = self.exc_cur()
        print('Converting {} 2/2...'.format(self.issue_month + '.xlsx'))
        df_exc['보상합계_기준'] = df_std['보상합계']
        df_exc['변제합계_기준'] = df_std['변제합계']
        df_exc['C 통화'] = df_std['C 통화']
        df_exc['V 통화'] = df_std['V 통화']
        df_exc['사정년월'] = self.issue_month
        return df_exc

    def create_db(self):
        if 'Main_DB.db' not in os.listdir ('Main_DB'):
            temp_file = open_zipfile('Cookies\Pocket.zip', 'MainDB_init.txt')
            with open (temp_file, 'rt', encoding='UTF8') as file:
                query = ''.join ([line.strip () for line in file]).split (';')[1]
            with sqlite3.connect (self.db_address) as conn:
                cur = conn.cursor ()
                cur.execute (query)
                conn.commit ()
            self.set_index()
            self.table_setter()

    def table_setter(self):
        self.create_sapmap()
        self.save_sapmap()
        self.create_global()
        self.save_global()
        self.create_cbr()

    def set_index(self):
        temp_file = open_zipfile('Cookies\Pocket.zip', 'MainDB_init.txt')
        with open (temp_file, 'rt', encoding='UTF8') as file:
            query = ''.join([line.strip() for line in file]).split(';')[2]
        with sqlite3.connect (self.db_address) as conn:
            cur = conn.cursor ()
            cur.execute(query)
            conn.commit()

    def create_sapmap(self):
        temp_file = open_zipfile('Cookies\Pocket.zip', 'MainDB_init.txt')
        with open (temp_file, 'rt', encoding='UTF8') as file:
            query = ''.join([line.strip() for line in file]).split(';')[3]
        with sqlite3.connect (self.db_address) as conn:
            cur = conn.cursor ()
            cur.execute(query)
            conn.commit()

    @staticmethod
    def save_sapmap():
        sap_file = open_zipfile('Cookies\Pocket.zip', 'SAP.xlsx')
        with open(sap_file, 'rb') as file:
            df_sap = pd.read_excel(file, sheet_name=0, skiprows=0)
        with sqlite3.connect('Main_DB/Main_DB.db') as conn:
            df_sap.to_sql('SAP_MAPPING', con=conn, if_exists='append', index=None, index_label=None)

    def create_global(self):
        temp_file = open_zipfile('Cookies\Pocket.zip', 'MainDB_init.txt')
        with open (temp_file, 'rt', encoding='UTF8') as file:
            query = ''.join([line.strip() for line in file]).split(';')[4]
        with sqlite3.connect (self.db_address) as conn:
            cur = conn.cursor ()
            cur.execute(query)
            conn.commit()

    @staticmethod
    def save_global():
        sap_file = open_zipfile('Cookies\Pocket.zip', 'Global.xlsx')
        with open(sap_file, 'rb') as file:
            df_sap = pd.read_excel(file, sheet_name=0, skiprows=0)
        with sqlite3.connect('Main_DB/Main_DB.db') as conn:
            df_sap.to_sql('GLOBAL_VENDORS', con=conn, if_exists='append', index=None, index_label=None)

    @staticmethod
    def create_cbr():
        cbr_file = open_zipfile('Cookies\Pocket.zip', 'C_BR.xlsx')
        with open(cbr_file, 'rb') as file:
            df_cbr = pd.read_excel(file, sheet_name=0, skiprows=0)
        with sqlite3.connect('Main_DB/Main_DB.db') as conn1:
            df_cbr.to_sql('C_BRs', con=conn1, if_exists='replace', index=None, index_label=None)

    def delete_db(self):
        with sqlite3.connect (self.db_address) as conn:
            cur = conn.cursor ()
            query = ''' DELETE FROM Main_DB WHERE 사정년월=? '''
            cur.execute (query, (self.issue_month,))
            conn.commit ()

    def save_db(self):
        self.create_db ()
        df_exc = self.merge()
        with sqlite3.connect(self.db_address) as conn:
            try :
                df_exc.to_sql('Main_DB', con=conn, if_exists='append', index=None, index_label=None)
            except sqlite3.IntegrityError:
                print ('Existing {} data will be replaced'.format (self.issue_month ))
                self.delete_db()
                df_exc.to_sql('Main_DB', con=conn, if_exists='append', index=None, index_label=None)
        print('Saved {} in Main database'.format(self.issue_month + '.xlsx'))

    @classmethod
    def save_year(cls, year):
        print('Targets : {}'.format(', '.join(cls(year).issue_months)))
        for month in cls(year).issue_months:
            cls(month).save_db ()


if __name__=="__main__":
    # MainDB (202002).save_db()
    # MainDB.save_year(2019)
    MainDB(201901).create_db()
