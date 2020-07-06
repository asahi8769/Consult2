import sqlite3
import pandas as pd
import warnings
import numpy as np
import gc
import os

warnings.filterwarnings('ignore')


class VendorCollection:
    def __init__(self, df):
        self.df = df[['고객사', '사정년월', '통보서', 'V List No', '원인부품번호', '업체코드', '업체명', '보상합계', 'V 분담율',
                       'V 통화', '변제합계', '보상합계_기준', '변제합계_기준']]
        self.months = None
        self.issue_month = None
        self.df_collect = None
        self.db_address = os.path.join('Main_DB', 'Main_DB.db')
        self.campaign_list = list()
        self.setoff_check = False

    def sap_data(self):
        with sqlite3.connect(self.db_address) as conn:
            df_sap = pd.read_sql('SELECT * FROM SAP_MAPPING', conn)
            return df_sap[['업체코드', 'SAP코드']]

    def global_data(self):
        with sqlite3.connect(self.db_address) as conn:
            df_glo = pd.read_sql('SELECT * FROM GLOBAL_VENDORS', conn)
            return df_glo[['업체코드', '글로벌', '선통보']]

    def pre_processing(self):
        self.df['CW'] = [i[-2] for i in self.df['통보서']]
        self.months = sorted(self.df['사정년월'].unique().tolist())
        if len(self.months) == 1:
            self.issue_month = str(self.months[0])
        elif len(self.months) > 1:
            self.issue_month = '{}-{}'.format(self.months[0], self.months[-1])

    def init_pivot_exc(self):
        self.pre_processing()
        self.df_collect = self.df.pivot_table(
            values='변제합계',
            index=['사정년월', '업체코드', '업체명'],
            columns=['CW'],
            margins=False,
            aggfunc=np.sum,
            fill_value=0)

    def merge_base(self):
        self.init_pivot_exc()
        df_collect_base = self.df.pivot_table(
            values='변제합계_기준',
            index=['사정년월', '업체코드', '업체명'],
            margins=False,
            aggfunc=np.sum,
            fill_value=0)
        self.df_collect.columns = self.df_collect.columns.get_level_values(0)
        self.df_collect = self.df_collect.rename({'C': '캠페인', 'W': '일반'}, axis=1)
        self.df_collect['합계'] = self.df_collect['캠페인'] + self.df_collect['일반']
        self.df_collect['합계_기준'] = df_collect_base['변제합계_기준']
        self.df_collect = self.df_collect[self.df_collect['합계'] != 0]
        del [[df_collect_base]]
        gc.collect()
        self.df_collect.sort_values(by=['합계'], ascending=False, inplace=True)
        self.df_collect.reset_index(inplace=True)
        self.df_collect['비중(%)'] = (self.df_collect['합계'] / self.df_collect['합계'].sum() * 100).apply(round, args=[2])

    def merge_aux(self):
        self.merge_base()
        df_sap = self.sap_data()
        df_glo = self.global_data()
        self.df_collect = self.df_collect.merge(df_sap, on='업체코드', how='left')
        self.df_collect = self.df_collect.merge(df_glo, on='업체코드', how='left')
        del [[df_sap, df_glo]]
        gc.collect()
        col = ['사정년월', '업체코드', 'SAP코드', '업체명', '글로벌', '선통보', '캠페인', '일반', '합계',
               '비중(%)', '합계_기준']
        self.df_collect = self.df_collect[col]
        self.df_collect.fillna('', inplace=True)

    def set_off_check(self):
        if len(self.months) == 1 and 'SALES.xlsx' in os.listdir('Cookies'):
            with open(r'Cookies\SALES.xlsx', 'rb') as file:
                df_sales = pd.read_excel(file)
            df_sales = df_sales[['업체코드', '입고금액 KRW']]
            df_sales = df_sales.pivot_table(
                values='입고금액 KRW',
                index=['업체코드'],
                margins=False,
                margins_name='',
                aggfunc=np.sum,
                fill_value=0)
            df_sales = df_sales.apply(pd.to_numeric, errors='ignore')
            self.df_collect = self.df_collect.merge(df_sales, on='업체코드', how='left')
            self.df_collect = self.df_collect.rename({'입고금액 KRW': '입고금액'}, axis=1)
            self.df_collect['입고금액'] = self.df_collect['입고금액'].apply(pd.to_numeric, errors='ignore')
            self.df_collect.fillna(0, inplace=True)
            self.setoff_check = True
            self.df_collect['검증'] = self.df_collect.apply(self.test, axis=1)

    @staticmethod
    def test(df_collect):
        if df_collect['합계'] < 0:
            return '상계가능'
        elif df_collect['입고금액'] == 0:
            return '매입없음'
        elif df_collect['입고금액'] < df_collect['합계']:
            return '매입대부족'
        elif df_collect['입고금액'] >= df_collect['합계']:
            return '상계가능'
        else:
            return ''

    def customers(self):
        customers = list()
        for vendor in list(self.df_collect['업체코드']):
            customer = ', '.join(set(self.df[self.df['업체코드'] == vendor][['고객사']]['고객사']))
            customers.append(customer)
        self.df_collect['고객사'] = customers

    def convert(self):
        self.merge_aux()
        try:
            self.set_off_check()
        except:
            pass
        if self.setoff_check is True:
            self.df_collect.loc['요약'] = pd.Series({
                '캠페인': self.df_collect['캠페인'].sum(),
                '일반': self.df_collect['일반'].sum(),
                '합계': self.df_collect['합계'].sum(),
                '비중(%)': 100,
                '합계_기준': self.df_collect['합계_기준'].sum(),
                '입고금액': self.df_collect['입고금액'].sum()})
        else:
            self.df_collect.loc['요약'] = pd.Series({
                '캠페인': self.df_collect['캠페인'].sum(),
                '일반': self.df_collect['일반'].sum(),
                '합계': self.df_collect['합계'].sum(),
                '비중(%)': 100,
                '합계_기준': self.df_collect['합계_기준'].sum()})
        self.customers()
        del [[self.df]]
        gc.collect()
        self.df_collect.fillna("", inplace=True)
        return self.df_collect


if __name__ == '__main__':
    from Search_DB import SearchDB

    df1 = VendorCollection(
        SearchDB(customer=None, start_m=202004, end_m=None,
                 ro_no=None, part_no=None, ven_code=None).search()).convert()

    with pd.ExcelWriter('Spawn/test3.xlsx', mode='r') as writer:
        df1.to_excel(writer, sheet_name='Sheet_name_4', index=False)
