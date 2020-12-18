import pandas as pd
import warnings
import numpy as np
import gc
import os
import re

warnings.filterwarnings('ignore')


class ClosingInfo:
    PLANTS = ['HMMA', 'KMMG', 'KMS', 'HMMC', 'HMMR', 'HMB', 'KMM', 'HAOS', 'DWI','HMI', 'KMI', 'CHMC', 'YOUNGSAN']
    MONTHS = {'01': 'JAN', '02': 'FEB', '03': 'MAR', '04': 'APR', '05': 'MAY', '06': 'JUN', '07': 'JUL',
              '08': 'AUG', '09': 'SEP', '10': 'OCT', '11': 'NOV', '12': 'DEC'}

    def __init__(self, df):
        self.df = df[['고객사', '클레임상태', 'C 분담율', 'V 분담율', '변제합계_기준', '보상합계_기준',
                        '통보서', 'E/D', 'V적용환율', '변제합계', '보상합계', '사정년월']]
        self.month = sorted(self.df['사정년월'].unique().tolist(), reverse=True)[0]
        self.df = self.df[self.df['사정년월'] == int(self.month)]
        self.df_inv = None

    def pre_processing(self):
        self.df['분담율검증'] = self.df['C 분담율'] == self.df['V 분담율']
        self.df['금액검증'] = self.df['변제합계_기준'] - self.df['보상합계_기준']
        self.df['Issue_No'] = self.df['통보서'].map(lambda x: str(x)[0:7]) + '-' + self.df['통보서'].map(
            lambda x: str(x)[7:])
        self.df['Month'] = self.df['통보서'].map(lambda x: str(x)[4:6])
        self.df = self.df.replace({'Month': self.MONTHS})
        self.df['Year'] = self.df['통보서'].map(lambda x: str(x)[0:4])
        self.df['Remarks'] = self.df['Month'].map(str) + ' - ' + self.df['Year'].map(str) + ' / ' + self.df[
            'E/D'].map(str)
        self.df.drop(['Month', 'Year'], axis=1, inplace=True)

    def convert(self):
        self.pre_processing()
        for plant in self.PLANTS:
            df_partial = self.df[self.df['고객사'] == plant]
            if len(df_partial) != 0:
                df_partial = self.closing_conditions(df_partial)
                if self.df_inv is None:
                    self.df_inv = df_partial
                else:
                    self.df_inv = pd.concat([self.df_inv, df_partial])
            else :
                pass
        return self.df_inv

    def closing_conditions(self, df_partial):
        int_br = ', '.join(sorted(list(set(df_partial[(df_partial['클레임상태'] == 'B1') & (
                df_partial['분담율검증'] == False)]['C 분담율'].apply(str)))))
        unsettled = len(df_partial[df_partial['클레임상태'] == 'X1'])
        exchange = self.exchange_rate(df_partial)
        c1_amounts = str(round(df_partial[df_partial['클레임상태'] == 'C1']['보상합계_기준'].sum(), 2))
        payment = str (round(df_partial['보상합계'].sum (), 0))
        collect = str (round (df_partial['변제합계'].sum (), 0))
        difference = str(round(df_partial['변제합계'].sum() - df_partial['보상합계'].sum(), 0))
        reimb_rate = str(round(df_partial['변제합계'].sum() / df_partial['보상합계'].sum() * 100, 2))
        df_partial = df_partial[(df_partial['보상합계_기준'] != 0)]
        df_partial = df_partial.pivot_table(
            values='보상합계_기준', index=['고객사', 'Issue_No', 'Remarks'],
            aggfunc=[np.size, np.sum], fill_value=0)
        df_partial.columns = ['_'.join(col).strip() for col in df_partial.columns.values]
        df_partial.reset_index(inplace=True)
        df_partial = df_partial.rename({'sum_보상합계_기준': 'CLAIM_AMOUNT', 'size_보상합계_기준': '건수'}, axis=1)
        df_partial.sort_values(by=['고객사', 'Issue_No'], ascending=False, inplace=True)
        df_partial = df_partial[['고객사', 'Issue_No', '건수', 'CLAIM_AMOUNT', 'Remarks']]
        df_partial.loc['total'] = pd.Series(
            {'고객사': df_partial['고객사'].iloc[0], 'Issue_No': 'Total', '건수': df_partial['건수'].sum(),
             'CLAIM_AMOUNT': df_partial['CLAIM_AMOUNT'].sum(), 'Remarks': '-'})
        df_partial.loc['통분'] = pd.Series(
            {'고객사': df_partial['고객사'].iloc[0], 'Issue_No': '통분', '건수': int_br, 'CLAIM_AMOUNT': '', 'Remarks': ''})
        df_partial.loc['미확정'] = pd.Series(
            {'고객사': df_partial['고객사'].iloc[0], 'Issue_No': '미확정', '건수': unsettled, 'CLAIM_AMOUNT': '', 'Remarks': ''})
        df_partial.loc['환율'] = pd.Series(
            {'고객사': df_partial['고객사'].iloc[0], 'Issue_No': '환율', '건수': exchange, 'CLAIM_AMOUNT': '', 'Remarks': ''})
        df_partial.loc['C1'] = pd.Series(
            {'고객사': df_partial['고객사'].iloc[0], 'Issue_No': 'C1금액', '건수': c1_amounts, 'CLAIM_AMOUNT': '', 'Remarks': ''})
        df_partial.loc['보상'] = pd.Series(
            {'고객사': df_partial['고객사'].iloc[0], 'Issue_No': '보상', '건수': payment, 'CLAIM_AMOUNT': '', 'Remarks': ''})
        df_partial.loc['변제'] = pd.Series(
            {'고객사': df_partial['고객사'].iloc[0], 'Issue_No': '변제', '건수': collect, 'CLAIM_AMOUNT': '', 'Remarks': ''})
        df_partial.loc['차액'] = pd.Series(
            {'고객사': df_partial['고객사'].iloc[0], 'Issue_No': '차액', '건수': difference, 'CLAIM_AMOUNT': '', 'Remarks': ''})
        df_partial.loc['변제율(%)'] = pd.Series(
            {'고객사': df_partial['고객사'].iloc[0], 'Issue_No': '변제율(%)', '건수': reimb_rate, 'CLAIM_AMOUNT': '', 'Remarks': ''})
        return df_partial

    @staticmethod
    def exchange_rate(df_partial):
        # print(df_partial)
        if len(df_partial[df_partial['통보서'].str.contains('CC', regex=True, na=False) & (df_partial['보상합계_기준'] != 0)]) > 0:
            exc_rate_cc = df_partial[df_partial['통보서'].str.contains('CC', regex=True, na=False) & (df_partial['보상합계_기준'] != 0)]['V적용환율'].iloc[0]
            usd_amount = df_partial[df_partial['통보서'].str.contains('CC', regex=True, na=False) & (df_partial['보상합계_기준'] != 0)]['변제합계_기준'].iloc[0]
            krw_amount = df_partial[df_partial['통보서'].str.contains('CC', regex=True, na=False) & (df_partial['보상합계_기준'] != 0)]['변제합계'].iloc[0]
            return 'V환율:{}, CC적용환율:{:,.2f}'.format(exc_rate_cc, round(krw_amount/usd_amount,2))
        elif len(df_partial[df_partial['통보서'].str.contains('CA', regex=True, na=False) & (df_partial['보상합계_기준'] != 0)]) >0:
            exc_rate_cc = df_partial[df_partial['통보서'].str.contains('CA', regex=True, na=False) & (df_partial['보상합계_기준'] != 0)]['V적용환율'].iloc[0]
            usd_amount = df_partial[df_partial['통보서'].str.contains('CA', regex=True, na=False) & (df_partial['보상합계_기준'] != 0)]['변제합계_기준'].iloc[0]
            krw_amount = df_partial[df_partial['통보서'].str.contains('CA', regex=True, na=False) & (df_partial['보상합계_기준'] != 0)]['변제합계'].iloc[0]
            return 'V환율:{}, CC적용환율:{:,.2f}'.format(exc_rate_cc, round(krw_amount/usd_amount,2))
        else :
            return 'V환율:{}'.format(df_partial[(df_partial['V적용환율'] != 0) & (df_partial['보상합계_기준'] != 0)]['V적용환율'].iloc[0])


if __name__ == '__main__':
    from Search_DB import SearchDB

    df = SearchDB(customer=None, start_m=202010, end_m=None, ro_no=None, part_no=None, ven_code=None).search()
    df_4 = ClosingInfo(df).convert()
    print(len(df_4))
    #
    with pd.ExcelWriter('Spawn/test4.xlsx') as writer:
        df_4.to_excel(writer, sheet_name='Sheet_name_4', index=False)