from Append_DB import MainDB
from Search_DB import SearchDB
import pandas as pd

from Consult import ConsultInformation
from RRATE import RRateInformation
from Collection import VendorCollection
from Closing import ClosingInfo


class MainExecute:
    """ to satisfy sql query syntax additional ` is needed as column name delimiter. """
    consult_mul_col = ['고객사', '사정년월', '통보서', '`V List No`', '원인부품번호', '업체코드', '업체명', '보상합계', '`V 분담율`',
                       '`V 통화`', '변제합계', '보상합계_기준', '변제합계_기준', 'V적용환율', '`E/D`', '`C 분담율`', '클레임상태']

    def __init__(self, customer=None, start_m=None, end_m=None, ro_no=None, part_no=None, ven_code=None):
        self.customer = customer
        self.start_m = start_m
        self.end_m = end_m
        self.ro_no = ro_no
        self.part_no = part_no
        self.ven_code = ven_code
        self.attachment = None
        self.plot = None

    def append_m(self):
        if self.start_m:
            MainDB(self.start_m).save_db()
        else :
            print('Error : start_m is missing.')

    def append_y(self):
        if self.start_m:
            MainDB.save_year(int(str(self.start_m)[0:4]))
        else:
            print('Error : start_m(year) is missing.')

    def consult_mul(self):
        if self.start_m and self.end_m:
            df = SearchDB(start_m=self.start_m, end_m=self.end_m, partial=MainExecute.consult_mul_col).search()
            self.attachment = 'Spawn/{}-{}_Consult.xlsx'.format(self.start_m, self.end_m)
            df_1 = ConsultInformation(df).convert()
            df_2 = RRateInformation(df).convert()
            df_3 = VendorCollection(df).convert()
            df_4 = ClosingInfo(df).convert()
            with pd.ExcelWriter(self.attachment) as writer:
                df_1.to_excel(writer, sheet_name='1_품의정보', index=False)
                df_2.to_excel(writer, sheet_name='2_변제율', index=False)
                df_3.to_excel(writer, sheet_name='3_변제정보', index=False)
                df_4.to_excel(writer, sheet_name='4_마감정보', index=False)
        else:
            print('Error : Requirements(start_m, end_m) are missing.')

    def consult_sin(self):
        if self.start_m and not self.end_m:
            df = SearchDB(start_m=self.start_m, partial=MainExecute.consult_mul_col).search()
            self.attachment = 'Spawn/{}_Consult.xlsx'.format(self.start_m)
            df_1 = ConsultInformation(df).convert()
            df_2 = RRateInformation(df).convert()
            df_3 = VendorCollection(df).convert()
            df_4 = ClosingInfo(df).convert()
            with pd.ExcelWriter(self.attachment) as writer:
                df_1.to_excel(writer, sheet_name='1_품의정보', index=False)
                df_2.to_excel(writer, sheet_name='2_변제율', index=False)
                df_3.to_excel(writer, sheet_name='3_변제정보', index=False)
                df_4.to_excel(writer, sheet_name='4_마감정보', index=False)
        else :
            print('Error : Requirements(start_m, ~end_m) are missing.')


if __name__=="__main__":
    # obj = MainExecute(start_m=201910, end_m=201912)
    # obj.consult_mul()

    obj_2 = MainExecute(start_m=201910, end_m=None)
    obj_2.consult_sin()

