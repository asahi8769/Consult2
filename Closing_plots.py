from matplotlib import pyplot as plt
from matplotlib.gridspec import GridSpec
from matplotlib import cm
from matplotlib.offsetbox import AnchoredText
import pandas as pd
import numpy as np
import warnings
import sys
import os
warnings.filterwarnings ('ignore')


class MakeClosingPlots:
    def __init__(self, filename):
        self.filename = filename
        self.issuemonth = filename.split('.')[0]
        if self.filename in os.listdir('Spawn'):
            with open(os.path.join('Spawn', self.filename), 'rb') as file:
                self.df = pd.read_excel(file, sheet_name='2_변제율', skiprows=0)
        else :
            print('No such file')
            sys.exit()

    def loss_surplus(self):
        df_pivoted = self.df.pivot_table(values='차액', index=['고객사'], margins=False, aggfunc=np.sum)
        df_pivoted = df_pivoted[df_pivoted.index != '-']
        df_pivoted.sort_values(by=['차액'], ascending=False, inplace=True)
        customers = df_pivoted.index.tolist()
        if 'YOUNGSAN' in customers:
            customers[customers.index('YOUNGSAN')] = 'YSG_'
        differences = df_pivoted['차액'].tolist()
        differences = [i / 1000000 for i in differences]
        total_diff = np.sum(differences)
        return customers, differences, total_diff

    def reimb_rate(self):
        df_pivoted = self.df.pivot_table(values=['보상합계', '변제합계'], index=['고객사'], margins=False, aggfunc=np.sum)
        df_pivoted = df_pivoted[df_pivoted.index != '-']
        df_pivoted['변제율(%)'] = df_pivoted['변제합계'] / df_pivoted['보상합계'] * 100
        df_pivoted.sort_values(by=['변제율(%)'], ascending=False, inplace=True)
        customers = df_pivoted.index.tolist()
        if 'YOUNGSAN' in customers:
            customers[customers.index('YOUNGSAN')] = 'YSG_'
        reimb_rate = df_pivoted['변제율(%)'].tolist()
        total_reimb = df_pivoted['변제합계'].sum() / df_pivoted['보상합계'].sum() * 100
        return customers, reimb_rate, total_reimb

    @staticmethod
    def bar_colormap(y, a, cmap='coolwarm_r', amplifier=5, threshold=None):
        """ colormap options :  coolwarm, coolwarm_r, RdBu, bwr, seismic, Spectral, RdYIBu """
        norm_y = [(threshold-i) / np.std (y) for i in y]
        bar_cmap = cm.get_cmap (cmap)
        bar_color = [bar_cmap (x * amplifier) for x in norm_y]
        norm_a = (threshold-a) / np.std (y)
        line_cmap = cm.get_cmap (cmap)
        line_color = line_cmap (norm_a * amplifier)
        return bar_color, line_color

    @staticmethod
    def barplot(ax, x, y, a, bar_color, line_color):
        ax.bar(x, y, align='center', alpha=0.5, color=bar_color)
        ax.axhline(y=a, color=line_color, alpha=0.7, linestyle='-.')
        ax.tick_params(axis='x', rotation=45)
        ax.grid(True, linestyle='--')
        ax.axis(option='tight')

    def plot_data(self):
        fig = plt.figure(constrained_layout=True, figsize=(7, 4))
        gs = GridSpec(1, 2, figure=fig)

        """ plot 1: loss/surplus per customer in the year """
        ax1 = fig.add_subplot(gs[0:1, 0:1])
        ax1.set(ylabel='Loss/Surplus(Scale: million KRW)', title='Loss/Surplus in {}'.format(self.issuemonth.split('_')[0]))
        customers, differences, total_diff = self.loss_surplus()
        bar_color, line_color = self.bar_colormap(differences, total_diff, amplifier=10, threshold=0)
        self.barplot(ax1, customers, differences, total_diff, bar_color, line_color)
        at1 = AnchoredText('Difference = {:,.2f}M KRW'.format(total_diff), prop=dict(size=10, color=line_color),
                           frameon=True, loc='upper right')
        at1.patch.set_boxstyle("round,pad=0.,rounding_size=0.1")
        ax1.add_artist(at1)
        ax1.axhline(y=0, color='grey', alpha=0.9, linestyle='-')

        """ plot 2: reimbursement rate(%) per customer in the year """
        ax2 = fig.add_subplot(gs[0:1, 1:2])
        ax2.set(ylabel='Rate(%)', title='Reimbursement Rate in {}'.format(self.issuemonth.split('_')[0]))
        customers, reimb_rate, total_reimb = self.reimb_rate()
        bar_color, line_color = self.bar_colormap(reimb_rate, total_reimb, amplifier=20, threshold=100)
        self.barplot(ax2, customers, reimb_rate, total_reimb, bar_color, line_color)
        at2 = AnchoredText('Rate = {:,.2f}%'.format(total_reimb).format(total_diff),
                           prop=dict(size=10, color=line_color), frameon=True,
                           loc='upper right')
        at2.patch.set_boxstyle("round,pad=0.,rounding_size=0.1")
        ax2.add_artist(at2)
        ax2.set_ylim([80, 120])
        ax2.axhline(y=100, color='grey', alpha=0.9, linestyle='-')

        """ save image """
        plt.savefig(os.path.join('Images', self.issuemonth+'.png'))
        return plt

if __name__=='__main__':
    obj = MakeClosingPlots('202005_Consult.xlsx')
    a,b,c = obj.loss_surplus()
    print(a)
    print(b)
    print(c)

    a,b,c = obj.reimb_rate()
    print(a)
    print(b)
    print(c)

    obj.plot_data()
