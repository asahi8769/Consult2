import sqlite3
import pandas as pd
import warnings
import numpy as np
import gc
import os

warnings.filterwarnings('ignore')


class ConsultInformation:
    def __init__(self, df):
        self.df = df[['고객사', '사정년월', '통보서', '클레임상태', 'V List No', '원인부품번호', '업체코드', '업체명', '보상합계',
                      'V 분담율', 'V 통화', '변제합계', '보상합계_기준', '변제합계_기준']]
        self.months = None
        self.df_converted = None
        self.db_address = os.path.join('Main_DB', 'Main_DB.db')

    @staticmethod
    def remark(df):
        if df['CW_'] == 'R':
            return '환불'
        elif df['CW_'] == 'F':
            return '수정전표'
        elif df['CW'] == 'C':
            return '캠페인'
        else:
            return ''

    def pre_processing(self):
        self.df = self.df[~(self.df['클레임상태'].isin(['B1', 'B2', 'E2']))]
        self.df['CW'] = [i[-2] for i in self.df['통보서']]
        self.df['CW_'] = [i[-1:] for i in self.df['V List No']]
        self.df['비고'] = self.df.apply(self.remark, axis=1)
        self.months = sorted(self.df['사정년월'].unique().tolist())

    def sap_data(self):
        with sqlite3.connect(self.db_address) as conn:
            df_sap = pd.read_sql('SELECT * FROM SAP_MAPPING', conn)
            return df_sap[['업체코드', 'SAP코드']]

    def convert(self):
        self.pre_processing()
        df_sap = self.sap_data()
        for month in self.months:
            no = 1
            df_month = self.df[self.df['사정년월'] == month]
            for plant_name in ['HMMA', 'KMMG', 'KMS', 'HMMC', 'HMMR', 'HMB', 'KMM', 'HAOS', 'HMI', 'KMI', 'CHMC',
                               'YOUNGSAN', 'DWI']:
                if len(df_month[(df_month['고객사'] == plant_name)]) == 0:
                    pass
                else:
                    for currency in ['KRW', 'USD', 'EUR']:
                        df_currency = df_month[(df_month['V 통화'] == currency) & (df_month['고객사'] == plant_name)]
                        if len(df_currency) == 0:
                            pass
                        else:
                            df_currency = df_currency.pivot_table(values=['변제합계_기준', '변제합계'],
                                                                  index=['사정년월', '고객사', 'V List No', '업체코드',
                                                                         '업체명', 'CW', 'V 분담율', 'V 통화', '비고'],
                                                                  margins=False,
                                                                  aggfunc=[np.sum, np.size])
                            df_currency.columns = ['_'.join(col).strip() for col in df_currency.columns.values]
                            df_currency = df_currency.rename(
                                {'sum_변제합계_기준': '변제합계_기준', 'sum_변제합계': '변제합계_환산',
                                 'size_변제합계': '건수'}, axis=1)
                            df_currency['변제합계_기준'] = (df_currency['변제합계_기준']).apply(round, args=[2])
                            df_currency.sort_values(by=['업체코드'], ascending=True, inplace=True)
                            df_currency.reset_index(inplace=True)
                            df_currency.index += 1

                            filter_1 = df_currency['V 통화'] == 'KRW'
                            df_currency['거래명세서_기준'] = np.where(filter_1, 0, df_currency['변제합계_기준'].apply(abs))
                            df_currency['거래명세서_환산'] = df_currency['변제합계_환산'].apply(abs)
                            df_currency['순번'] = [i for i in range(no, no + len(df_currency))]
                            no += len(df_currency)

                            val_1 = '{}_{}_요약'.format(plant_name, currency)
                            df_currency.loc[val_1] = pd.Series({
                                '사정년월': month,
                                '순번': val_1,
                                '고객사': plant_name,
                                '건수': df_currency['건수'].sum(),
                                'V 통화': currency,
                                '변제합계_기준': df_currency['변제합계_기준'].sum(),
                                '변제합계_환산': df_currency['변제합계_환산'].sum(),
                                '비고': '요약',
                                '거래명세서_기준': df_currency['거래명세서_기준'].sum(),
                                '거래명세서_환산': df_currency['거래명세서_환산'].sum()})
                            df_currency = df_currency.merge(df_sap, on='업체코드', how='left')
                            df_currency.loc[
                                (df_currency['CW'] == 'F') & (df_currency['V 통화'] == 'KRW'), ['변제합계_기준',
                                                                                              '거래명세서_기준']] = 0
                            df_currency = df_currency.fillna('')
                            if self.df_converted is None:
                                self.df_converted = df_currency
                            else:
                                self.df_converted = pd.concat([self.df_converted, df_currency])
                        del [[df_currency]]
                        gc.collect()

                    val_2 = '{}_{}_총합계'.format(month, plant_name)
                    self.df_converted.loc[val_2] = pd.Series({
                        '사정년월': month,
                        '순번': '{}_총합계'.format(plant_name),
                        '고객사': plant_name,
                        '건수': self.df_converted[
                            (self.df_converted['사정년월'] == month) &
                            (self.df_converted['고객사'] == plant_name) &
                            (self.df_converted['비고'] != '요약')]['건수'].sum(),
                        'V 통화': '-',
                        '변제합계_기준': '-',
                        '변제합계_환산': self.df_converted[
                            (self.df_converted['사정년월'] == month) &
                            (self.df_converted['고객사'] == plant_name) &
                            (self.df_converted['비고'] != '요약')]['변제합계_환산'].sum(),
                        '비고': '요약',
                        '거래명세서_기준': '-',
                        '거래명세서_환산': self.df_converted[
                            (self.df_converted['사정년월'] == month) &
                            (self.df_converted['고객사'] == plant_name) &
                            (self.df_converted['비고'] != '요약')][
                            '거래명세서_환산'].sum(), })
            del [[df_month]]
            gc.collect()
            val_3 = str('{}_전체합계'.format(month))
            self.df_converted.loc[val_3] = pd.Series({
                '사정년월': month,
                '순번': val_3,
                '고객사': '-',
                '건수': self.df_converted[
                    (self.df_converted['사정년월'] == month) &
                    (self.df_converted['비고'] != '요약')]['건수'].sum(),
                'V 통화': '-',
                '변제합계_기준': '-',
                '변제합계_환산': self.df_converted[
                    (self.df_converted['사정년월'] == month) &
                    (self.df_converted['비고'] != '요약')]['변제합계_환산'].sum(),
                '비고': '요약',
                '거래명세서_기준': '-',
                '거래명세서_환산': self.df_converted[
                    (self.df_converted['사정년월'] == month) &
                    (self.df_converted['비고'] != '요약')]['거래명세서_환산'].sum()})

        cols = ['사정년월', '순번', '고객사', 'V List No', '업체코드', 'SAP코드', '업체명', 'V 분담율',
                '건수', 'V 통화', '변제합계_기준', '변제합계_환산', '비고', '거래명세서_기준', '거래명세서_환산']

        if len(self.months) == 1:
            self.df_converted = self.df_converted[cols[1:]]
        elif len(self.months) > 1:
            self.df_converted = self.df_converted[cols]
            val_4 = str('{}-{}_전체합계'.format(self.months[0], self.months[-1]))
            self.df_converted.loc[val_4] = pd.Series({
                '사정년월': str('{}-{}'.format(self.months[0], self.months[-1])),
                '순번': val_4,
                '고객사': '-',
                '건수': self.df_converted[(self.df_converted['비고'] != '요약')][
                    '건수'].sum(),
                'V 통화': '-',
                '변제합계_기준': '-',
                '변제합계_환산': self.df_converted[(self.df_converted['비고'] != '요약')][
                    '변제합계_환산'].sum(),
                '비고': '요약',
                '거래명세서_기준': '-',
                '거래명세서_환산': self.df_converted[(self.df_converted['비고'] != '요약')][
                    '거래명세서_환산'].sum()})

        self.df_converted.reset_index(inplace=True, drop=True)
        self.df_converted.index += 1
        return self.df_converted.fillna('')


if __name__ == '__main__':
    from Search_DB import SearchDB

    df1 = ConsultInformation(
        SearchDB(customer=None, start_m=201901, end_m=None,
                 ro_no=None, part_no=None, ven_code=None).search()).convert()
    df2 = ConsultInformation(
        SearchDB(customer=None, start_m=201902, end_m=None,
                 ro_no=None, part_no=None, ven_code=None).search()).convert()
    with pd.ExcelWriter('Spawn/test2.xlsx', mode='a') as writer:
        df1.to_excel(writer, sheet_name='Sheet_name_1', index=False)
        df2.to_excel(writer, sheet_name='Sheet_name_2', index=False)
