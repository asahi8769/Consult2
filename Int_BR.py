from matplotlib import pyplot as plt
from matplotlib.gridspec import GridSpec
from pathlib import Path
import pandas as pd
import numpy as np
import sqlite3
import warnings
import sys
import gc
import os
from Search_DB import SearchDB
from utils.export import export

warnings.filterwarnings('ignore')


class SimpleIntBRSolver:
    def __init__(self, df):
        df = df.copy()
        df = df[df['WFR'] != 'F']
        self.answer = {}
        for i in ('Regular', 'Exchange'):
            df_ = df[df['일반/교류'] == i]
            if len(df_) == 0:
                break
            self.a = df_[(df_['분담율검증'] == 'Integrated BR') & (df_['보상합계_기준'] != 0)]['변제대상 합계'].sum()
            self.b = df_[df_['분담율검증'] == 'Special BR']['보상합계_기준'].sum()
            self.A = df_['변제합계_기준'].sum()
            self.answer[i] = round((self.A - self.b) / self.a, 4)
        print(self.answer)


class IntBR:
    consult_mul_col = ['고객사', '사정년월', '통보서', '`V List No`', '`RO-NO`', '`OP-CODE`', '클레임상태', '원인부품번호', '원인부품명',
                       '업체코드', '업체명', '차종모델코드', '`C 분담율`', '보상합계', '보상합계_기준', '`변제대상 합계`',
                       '`V 통화`', '`V 분담율`', '변제합계', '변제합계_기준']

    def __init__(self, plant, start_m, end_m):
        with sqlite3.connect('Main_DB/Main_DB.db') as conn:
            self.df_cbr = pd.read_sql("SELECT * FROM 'C_BRs'", conn).fillna('')
        self.plant = plant.upper()
        self.start_m = start_m
        self.end_m = end_m
        self.df = SearchDB(
            customer=self.plant, start_m=self.start_m, end_m=self.end_m, ro_no=None, part_no=None,
            ven_code=None, partial=self.consult_mul_col).search()
        self.months = sorted(list(self.df['사정년월'].unique()))
        self.defined_cbr = sorted([float(i) for i in self.df_cbr[self.plant] if i != ''], reverse=True)
        self.re_class = None
        self.unexpected_BR = None
        self.new_cbr = dict()
        self.test_message = None
        self.pre_processing()

    def define_cbr(self, df):
        if df['C 분담율'] in self.defined_cbr:
            return 'Integrated BR'
        else:
            return 'Special BR'

    def pre_processing(self):
        self.df['분담율검증'] = self.df.apply(self.define_cbr, axis=1)
        filter_1 = self.df['통보서'].apply(lambda x: x[-3]) == '2'
        self.df['일반/교류'] = np.where(filter_1, 'Exchange', 'Regular')
        filter_2 = (self.df['클레임상태'] == 'E2') & (self.df['변제대상 합계'].apply(int) >= 0)
        self.df['변제대상 합계'] = np.where(filter_2, -self.df['변제대상 합계'], self.df['변제대상 합계'])
        self.df['WFR'] = [i[-1:] for i in self.df['통보서']]
        self.br_inspection()

    def br_inspection(self):
        found_cbr = sorted(list(
            self.df[(self.df['클레임상태'] == 'B1') & (self.df['C 분담율'] != self.df['V 분담율'])]['C 분담율'].unique()))
        self.unexpected_BR = f"Unexpected BR : {', '.join([str(i) for i in found_cbr if i not in self.defined_cbr])}"

    def extract(self):
        self.df = self.df[self.df['분담율검증'] == 'Integrated BR']
        self.df = self.df[self.df['클레임상태'].isin(['B1', 'E2'])]
        self.df = self.df[self.df['WFR'] != 'F']

    def pivot(self, nw):
        df_pivot = self.df[self.df['일반/교류'] == nw].pivot_table(
            values=['변제대상 합계', '보상합계_기준', '변제합계_기준'],
            index=['업체코드', '업체명', 'C 분담율', 'V 분담율'],
            margins=False,
            margins_name='Total',
            aggfunc=[np.sum, np.size])
        df_pivot.columns = ['_'.join(col).strip() for col in df_pivot.columns]
        df_pivot = df_pivot.rename(
            {'sum_변제대상 합계': '변제대상합계',
             'sum_변제합계_기준': '변제합계',
             'sum_보상합계_기준': '보상합계',
             'size_변제대상 합계': '건수'}, axis=1)
        df_pivot.reset_index(inplace=True)
        df_pivot = df_pivot[[
            '업체코드', '업체명', 'C 분담율', 'V 분담율', '변제대상합계', '변제합계', '보상합계', '건수']]
        df_pivot['건수'] = df_pivot['건수'].apply(int)
        return df_pivot

    def br_portion(self, df_pivot):
        df_pivot['점유율'] = df_pivot['변제대상합계'] / df_pivot['변제대상합계'].sum()
        df_pivot['분담율'] = df_pivot['점유율'] * df_pivot['V 분담율']
        df_pivot['분담율'] = df_pivot['분담율'].apply(round, args=[4])
        return df_pivot

    def calculate(self):
        self.extract()
        with pd.ExcelWriter(f"Spawn/{self.plant}_{self.start_m}-{self.end_m}.xlsx") as writer:
            for nw in sorted(list(self.df['일반/교류'].unique()), reverse=True):
                if len(self.df[self.df['일반/교류'] == nw]) == 0:
                    pass
                else:
                    df_pivot = self.pivot(nw)
                    df_pivot = self.br_portion(df_pivot)
                    df_pivot.sort_values(by=['점유율'], ascending=False, inplace=True)
                    new_nw_cbr = round(df_pivot['분담율'].sum(), 4)
                    df_pivot.loc['total'] = pd.Series({
                        '업체코드': 'Total',
                        '건수': df_pivot['건수'].sum(),
                        '변제대상합계': df_pivot['변제대상합계'].sum(),
                        '보상합계': df_pivot['보상합계'].sum(),
                        '변제합계': df_pivot['변제합계'].sum(),
                        '점유율': df_pivot['점유율'].sum(),
                        '분담율': new_nw_cbr
                    })
                    df_pivot = df_pivot.fillna('')
                    self.new_cbr[nw] = new_nw_cbr
                    df_pivot.to_excel(writer, sheet_name=nw, index=False)
        return f"{self.plant}_{self.start_m}-{self.end_m}.xlsx"

    def cbr_test_col(self, df_test):
        if df_test['분담율검증'] == 'Integrated BR' and df_test['일반/교류'] == 'Regular':
            return self.new_cbr['Regular']
        elif df_test['분담율검증'] == 'Integrated BR' and df_test['일반/교류'] == 'Exchange':
            return self.new_cbr['Exchange']
        else:
            return df_test['C 분담율']

    def test(self):
        self.df = SearchDB(
            customer=self.plant, start_m=self.start_m, end_m=self.end_m, ro_no=None, part_no=None,
            ven_code=None, partial=self.consult_mul_col).search()
        self.pre_processing()
        self.df = self.df[self.df['WFR'] != 'F']
        self.df['신규_C_분담율'] = self.df.apply(self.cbr_test_col, axis=1)
        self.df['신규_보상합계'] = round(self.df['보상합계_기준'] / self.df['C 분담율'] * self.df['신규_C_분담율'], 2)
        avg_old_reimb = round(self.df['변제합계_기준'].sum() / self.df['보상합계_기준'].sum() * 100, 2)
        avg_new_reimb = round(self.df['변제합계_기준'].sum() / self.df['신규_보상합계'].sum() * 100, 2)
        plot_info = list()
        for month in self.months:
            df_month = self.df[self.df['사정년월']==month]
            old_diff = round(df_month['변제합계_기준'].sum() - df_month['보상합계_기준'].sum(), 2)
            new_diff = round(df_month['변제합계_기준'].sum() - df_month['신규_보상합계'].sum(), 2)
            old_reimb = round(df_month['변제합계_기준'].sum() / df_month['보상합계_기준'].sum() * 100, 2)
            new_reimb = round(df_month['변제합계_기준'].sum() / df_month['신규_보상합계'].sum() * 100, 2)
            plot_info.append([old_diff, new_diff, old_reimb, new_reimb])
        return plot_info, avg_old_reimb, avg_new_reimb

    def plot(self):
        plot_info, avg_old_reimb, avg_new_reimb = self.test()
        old_diffs = [i[0]/1000 for i in plot_info]
        new_diffs = [i[1]/1000 for i in plot_info]
        old_reimbs = [i[2] for i in plot_info]
        new_reimbs = [i[3] for i in plot_info]

        if len(old_diffs) >= 3:
            old_3_month_diffs_avg = round(np.mean(old_diffs[-3:]), 2)
            new_3_month_diffs_avg  = round(np.mean(new_diffs[-3:]), 2)
        else:
            old_3_month_diffs_avg = round(np.mean(old_diffs), 2)
            new_3_month_diffs_avg  = round(np.mean(new_diffs), 2)
        self.test_message = new_3_month_diffs_avg - old_3_month_diffs_avg
        avg_old_diff = np.mean(old_diffs)
        avg_new_diff = np.mean(new_diffs)

        fig = plt.figure(constrained_layout=True, figsize=(11, 5))
        gs = GridSpec(1, 2, figure=fig)
        ax1 = fig.add_subplot(gs[0:1, 0:1])
        ax1.set(ylabel='Loss/Surplus(C CUR, K)', title=f"{self.plant} L/S in {self.start_m}~{self.end_m}")
        ax1.grid(True, linestyle='--')
        ax1.axhline(y=0, color='grey', alpha=0.9, linestyle='-')
        ax1.tick_params(axis='x', rotation=45)
        ax1.plot([str(i) for i in self.months], old_diffs, ':o', color='C1', label='Old Difference')
        ax1.plot([str(i) for i in self.months], new_diffs, '-o', color='C0', label='New Difference')
        ax1.axhline(y=avg_old_diff, color='C1', alpha=0.7, linestyle='-.',
                    label='Old Total:{:,.2f}K\n(Monthly:{:,.2f}K)'.format(
                        avg_old_diff, avg_old_diff / len(self.months)))
        ax1.axhline(y=avg_new_diff, color='C0', alpha=0.7, linestyle='--',
                    label='New Total:{:,.2f}K\n(Monthly:{:,.2f}K)\n(3M_Delta:{:,.2f}K)'.format(
                        avg_new_diff, avg_new_diff / len(self.months), self.test_message))
        z = np.polyfit([i for i in range(1, len(self.months) + 1)], old_diffs, 1)
        y_hat = np.poly1d(z)([i for i in range(1, len(self.months) + 1)])
        ax1.plot([str(i) for i in self.months], y_hat, "--", color='grey', alpha=0.5, label='Trend Slope:{:,.2f}'.format(z[0]))
        ax1.legend()

        ax2 = fig.add_subplot(gs[0:1, 1:])
        ax2.set(ylabel='Rate(%)', title=f"{self.plant} R-Rate(%) in {self.start_m}~{self.end_m}")
        ax2.grid(True, linestyle='--')
        ax2.axhline(y=100, color='grey', alpha=0.9, linestyle='-')
        ax2.tick_params(axis='x', rotation=45)
        ax2.plot([str(i) for i in self.months], old_reimbs, ':o', color='C1', label='Old Rate')
        ax2.plot([str(i) for i in self.months], new_reimbs, '-o', color='C0', label='New Rate')
        ax2.axhline(y=avg_old_reimb, color='C1', alpha=0.7, linestyle='-.',
                    label='Old Total Rate:{:,.2f}%'.format(avg_old_reimb))
        ax2.axhline(y=avg_new_reimb, color='C0', alpha=0.7, linestyle='--',
                    label='New Total Rate:{:,.2f}%'.format(avg_new_reimb))
        z = np.polyfit([i for i in range(1, len(self.months) + 1)], old_reimbs, 1)
        y_hat = np.poly1d(z)([i for i in range(1, len(self.months) + 1)])
        ax2.plot([str(i) for i in self.months], y_hat, "--", color='grey', alpha=0.5, label='Trend Slope:{:,.2f}'.format(z[0]))
        ax2.legend()
        plt.savefig(os.path.join('Images', f"{self.plant}_{self.start_m}-{self.end_m}_plot.png"))
        # plt.show()


if __name__ == "__main__":
    obj = IntBR('hmma', 201901, 202003)
    print(obj.unexpected_BR)

    obj.calculate()
    print(obj.new_cbr)  # {'Regular': 0.1987, 'Exchange': 0.1381}

    # obj.new_cbr = {'Regular': 0.3087, 'Exchange': 0.2181}
    # print(obj.new_cbr)
    obj.plot()

    # obj.new_cbr = {'Regular': 0.1087, 'Exchange': 0.2181}
    # print(obj.new_cbr)
    # obj.plot()


    # print(plot_info)
    # print(obj.df)
    # print(obj.df)
    # export(obj.df, 'ver1', interval=1)
    # print(obj.defined_cbr)
