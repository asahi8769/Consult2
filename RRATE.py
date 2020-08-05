import pandas as pd
from dateutil import relativedelta
from datetime import datetime
import warnings
import gc
import os
import numpy as np


warnings.filterwarnings('ignore')


class RRateInformation:
    def __init__(self, df):
        self.df = df[['고객사', '사정년월', '통보서', '클레임상태', 'V List No', '원인부품번호', '업체코드', '업체명', '보상합계',
                      'V 분담율', 'V 통화', 'V적용환율', '변제합계', '보상합계_기준', '변제합계_기준']]
        self.months = None
        self.df_reimb = None
        self.db_address = os.path.join('Main_DB', 'Main_DB.db')
        self.campaign_list = list()

    def pre_processing(self):
        self.df['CW'] = [i[-2] for i in self.df['통보서']]
        self.months = sorted(self.df['사정년월'].unique().tolist())

    def convert(self):
        self.pre_processing()
        for month in self.months:
            df_month = self.df[self.df['사정년월'] == month]
            month_issued = datetime.strptime(str(month), '%Y%m') + relativedelta.relativedelta(months=-2)

            for plant_name in ['HMMA', 'KMMG', 'HAOS', 'DWI', 'KMS', 'HMMC', 'HMMR', 'HMB', 'CHMC', 'KMM', 'HMI', 'KMI',
                               'YOUNGSAN']:
                if len(df_month[df_month['고객사'] == plant_name]) == 0:
                    pass
                else:
                    exc_rate = df_month[(df_month['고객사'] == plant_name) &
                                        (df_month['클레임상태'].isin(['B1'])) &
                                        (df_month['통보서'].str.slice(7, 9, 1) != 'WF') &
                                        (df_month['통보서'].str.slice(0, 6, 1) == month_issued.strftime('%Y%m'))]['V적용환율'].iloc[0]
                    exc_rate = float(str(exc_rate).replace(',', ''))

                    def collection_exchange(df_month, plant_name, exc_rate):
                        issue_list = df_month['통보서'].tolist()
                        currency_list = df_month['V 통화'].tolist()
                        collection_list = df_month['변제합계_기준'].tolist()
                        customer_list = df_month['고객사'].tolist()
                        for n, item in enumerate(collection_list):
                            if str(customer_list[n] == plant_name and issue_list[n]).endswith('0WF') and currency_list[n] == 'KRW' :
                                collection_list[n] = round(collection_list[n] / exc_rate,2)
                            else :
                                pass
                        df_month['변제합계_기준'] = collection_list
                        return df_month

                    df_month = collection_exchange(df_month, plant_name, exc_rate)
                    invoice = round(df_month[df_month['고객사'] == plant_name]['보상합계_기준'].sum(), 2)
                    reimbursement = round(df_month[df_month['고객사'] == plant_name]['변제합계_기준'].sum(), 2)

                    # print(exc_rate,  plant_name, reimbursement)

                    payment = df_month[df_month['고객사'] == plant_name]['보상합계'].sum()
                    collection = df_month[df_month['고객사'] == plant_name]['변제합계'].sum()
                    diff = collection - payment
                    reimb_rate = round((collection / payment) * 100, 2)
                    collection_portion = round((collection / df_month['변제합계'].sum()) * 100, 2)
                    c_number = len(df_month[(df_month['고객사'] == plant_name) &
                                            (df_month['통보서'].str.slice(0, 6, 1) == month_issued.strftime('%Y%m'))])
                    v_number = len(df_month[(df_month['고객사'] == plant_name) & (df_month['클레임상태'].isin(['B1', 'B2', 'E2']))])
                    campaign_amount = df_month[(df_month['CW'] == 'C') & (
                            df_month['고객사'] == plant_name)]['변제합계'].sum()
                    campaign_portion = round(
                        df_month[(df_month['CW'] == 'C') & (df_month['고객사'] == plant_name)][
                            '변제합계'].sum() / collection * 100, 2)
                    campaign_numbers = len(
                        set(df_month[(df_month['CW'] == 'C') & (df_month['고객사'] == plant_name)]['원인부품번호']
                            .str.slice(0,6)))
                    campaign_customer = [i for i in set(
                        df_month[(df_month['CW'] == 'C') & (df_month['고객사'] == plant_name)]['원인부품번호'].
                            str.slice(0, 6))]
                    self.campaign_list = self.campaign_list + campaign_customer

                    reimb_exc = int(exc_rate * reimbursement)

                    exc_diff = int(collection - reimb_exc)
                    adj_diff = int(reimb_exc - payment)

                    dic = {'사정년월': month, '고객사': plant_name, '보상합계_기준': invoice, '변제합계_기준': reimbursement,
                           '보상합계': payment, '변제합계': collection, '환율' : exc_rate, '차액': diff, '변제율(%)': reimb_rate,
                           '점유율(변제, %)': collection_portion, '보상건수' : c_number, '변제건수' : v_number,
                           '변제환산' : reimb_exc, '환차손익' : exc_diff, '보정차액' : adj_diff,
                           '캠페인금액': campaign_amount, '캠페인비율(%)': campaign_portion, '캠페인건수': campaign_numbers}

                    df_concat = pd.DataFrame(dic, index=['사정년월'])
                    if self.df_reimb is None:
                        self.df_reimb = df_concat
                    else:
                        self.df_reimb = pd.concat([self.df_reimb, df_concat])
                    del [[df_concat]]
                    gc.collect()

            payment_sum = self.df_reimb[self.df_reimb['사정년월'] == month]['보상합계'].sum()
            collection_sum = self.df_reimb[self.df_reimb['사정년월'] == month]['변제합계'].sum()
            reimb_sum = round((collection_sum / payment_sum * 100), 2)
            por_sum = round((collection_sum / df_month['변제합계'].sum() * 100), 2)
            c_number_sum = self.df_reimb[self.df_reimb['사정년월'] == month]['보상건수'].sum()
            v_number_sum = self.df_reimb[self.df_reimb['사정년월'] == month]['변제건수'].sum()
            campaign_amount_sum = self.df_reimb[self.df_reimb['사정년월'] == month]['캠페인금액'].sum()
            campaign_portion_sum = round(campaign_amount_sum / collection_sum * 100, 2)
            campaign_numbers_sum = self.df_reimb[self.df_reimb['사정년월'] == month]['캠페인건수'].sum()

            reimb_exc_sum = self.df_reimb[self.df_reimb['사정년월'] == month]['변제환산'].sum()
            exc_diff_sum = self.df_reimb[self.df_reimb['사정년월'] == month]['환차손익'].sum()
            adj_diff_sum = self.df_reimb[self.df_reimb['사정년월'] == month]['보정차액'].sum()

            self.df_reimb.loc['{} 요약'.format(str(month))] = {
                '사정년월': '{} 요약'.format(str(month)),
                '고객사': '-',
                '보상합계_기준': '-',
                '변제합계_기준': '-',
                '보상합계': payment_sum,
                '변제합계': collection_sum,
                '환율': '-',
                '차액': collection_sum - payment_sum,
                '변제율(%)': reimb_sum,
                '점유율(변제, %)': por_sum,
                '보상건수' : c_number_sum,
                '변제건수' : v_number_sum,
                '변제환산' : reimb_exc_sum,
                '환차손익': exc_diff_sum,
                '보정차액': adj_diff_sum,
                '캠페인금액': campaign_amount_sum,
                '캠페인비율(%)': campaign_portion_sum,
                '캠페인건수': campaign_numbers_sum}

            del [[df_month]]
            gc.collect()
        if len(self.months) > 1:
            payment_sum = self.df_reimb[self.df_reimb['고객사'] != '-']['보상합계'].sum()
            collection_sum = self.df_reimb[self.df_reimb['고객사'] != '-']['변제합계'].sum()
            reimb_sum = round((collection_sum / payment_sum * 100), 2)
            campaign_amount_sum = self.df_reimb[self.df_reimb['고객사'] != '-']['캠페인금액'].sum()
            campaign_portion_sum = round(campaign_amount_sum / collection_sum * 100, 2)
            campaign_numbers_sum = len(list(set(self.campaign_list)))
            c_number_sum = self.df_reimb[self.df_reimb['고객사'] != '-']['보상건수'].sum()
            v_number_sum = self.df_reimb[self.df_reimb['고객사'] != '-']['변제건수'].sum()
            reimb_exc_sum = self.df_reimb[self.df_reimb['고객사'] != '-']['변제환산'].sum()
            exc_diff_sum = self.df_reimb[self.df_reimb['고객사'] != '-']['환차손익'].sum()
            adj_diff_sum = self.df_reimb[self.df_reimb['고객사'] != '-']['보정차액'].sum()

            self.df_reimb.loc['{}-{}_요약'.format(self.months[0], self.months[-1])] = {
                '사정년월': '{}-{} 요약'.format(self.months[0], self.months[-1]),
                '고객사': '-',
                '보상합계_기준': '-',
                '변제합계_기준': '-',
                '보상합계': payment_sum,
                '변제합계': collection_sum,
                '환율': '-',
                '차액': collection_sum - payment_sum,
                '변제율(%)': reimb_sum,
                '점유율(변제, %)': '-',
                '보상건수': c_number_sum,
                '변제건수': v_number_sum,
                '변제환산': reimb_exc_sum,
                '환차손익': exc_diff_sum,
                '보정차액': adj_diff_sum,
                '캠페인금액': campaign_amount_sum,
                '캠페인비율(%)': campaign_portion_sum,
                '캠페인건수': campaign_numbers_sum}
        self.df_reimb.reset_index (inplace=True)
        col = ['사정년월', '고객사', '보상합계_기준', '변제합계_기준', '보상합계', '변제합계', '환율', '차액', '변제율(%)', '점유율(변제, %)',
               '보상건수', '변제건수', '변제환산', '환차손익', '보정차액', '캠페인금액', '캠페인비율(%)', '캠페인건수']
        self.df_reimb.index += 1
        return self.df_reimb[col]


if __name__ == '__main__':
    from Search_DB import SearchDB
    df1 = RRateInformation(
        SearchDB(customer=None, start_m=202008, end_m=202008,
                 ro_no=None, part_no=None, ven_code=None).search()).convert()
    with pd.ExcelWriter('Spawn/test2.xlsx') as writer:
        df1.to_excel(writer, sheet_name='Sheet_name_3', index=False)

    # with pd.ExcelWriter('Spawn/test2.xlsx', mode='a') as writer:
    #     df1.to_excel(writer, sheet_name='Sheet_name_3', index=False)