from Append_DB import MainDB
import pandas as pd
import sqlite3
import os
from tqdm import tqdm


class SearchDB:
    def __init__(self, customer=None, start_m=None, end_m=None, ro_no=None, part_no=None, ven_code=None, partial=None):
        self.partial = partial
        if self.partial:
            self.query = '''SELECT {} FROM Main_DB WHERE '''.format(', '.join(self.partial))
        else :
            self.query = '''SELECT * FROM Main_DB WHERE '''
        self.db_address = os.path.join ('Main_DB', 'Main_DB.db')
        self.start_m = start_m
        self.end_m = end_m
        self.customer = customer
        self.ro_no = ro_no
        self.part_no = part_no
        self.ven_code = ven_code
        self.prev = False
        self.search_customer()
        self.search_issue_month()
        self.search_ro_no()
        self.search_part_no()
        self.search_ven_code()

    def search_customer(self):
        if self.customer:
            query_add = '''고객사='{}' '''.format(self.customer)
            self.query += query_add
            self.prev = True

    def search_issue_month(self):
        if self.prev:
            if self.start_m and not self.end_m:
                query_add = '''AND 사정년월={} '''.format(self.start_m)
                self.query += query_add
            elif self.start_m and self.end_m:
                query_add = '''AND 사정년월>={} AND 사정년월<={} '''.format (self.start_m, self.end_m)
                self.query += query_add
        else:
            if self.start_m and not self.end_m:
                query_add = '''사정년월='{}' '''.format(self.start_m)
                self.query += query_add
                self.prev = True
            elif self.start_m and self.end_m:
                query_add = '''사정년월>={} AND 사정년월<={} '''.format (self.start_m, self.end_m)
                self.query += query_add
                self.prev = True

    def search_ro_no(self):
        if self.prev:
            if self.ro_no:
                query_add = '''AND `RO-NO`='{}' '''.format(self.ro_no)
                self.query += query_add
        else:
            if self.ro_no:
                query_add = '''`RO-NO`='{}' '''.format(self.ro_no)
                self.query += query_add
                self.prev = True

    def search_part_no(self):
        if self.prev:
            if self.part_no:
                query_add = '''AND 원인부품번호='{}' '''.format(self.part_no)
                self.query += query_add
        else:
            if self.part_no:
                query_add = '''원인부품번호='{}' '''.format(self.part_no)
                self.query += query_add
                self.prev = True

    def search_ven_code(self):
        if self.prev:
            if self.ven_code:
                query_add = '''AND 업체코드='{}' '''.format(self.ven_code)
                self.query += query_add
        else:
            if self.ven_code:
                query_add = '''업체코드='{}' '''.format(self.ven_code)
                self.query += query_add
                self.prev = True

    def search(self):
        with sqlite3.connect (self.db_address) as conn:
            df = pd.read_sql_query (self.query, conn)
        return df

    @staticmethod
    def export(df):
        with pd.ExcelWriter ('Spawn/exported.xlsx', engine='openpyxl') as writer:
            interval = 3000
            for i in tqdm (range (0, len (df), interval)):
                if i == 0:
                    if i + interval <= len (df):
                        df.iloc[i:i + interval].to_excel (
                            writer, sheet_name='exported', index=False, header=True)
                    else:
                        df.iloc[i:len (df)].to_excel (
                            writer, sheet_name='exported', index=False, header=True)
                else:
                    if i + interval <= len (df):
                        df.iloc[i:i + interval].to_excel (
                            writer, sheet_name='exported', index=False, header=False,startrow=i + 1)
                    else:
                        df.iloc[i:len (df)].to_excel (
                            writer, sheet_name='exported', index=False, header=False, startrow=i + 1)


if __name__=="__main__":
    from utils.export import export
    obj = SearchDB (customer=None, start_m=201912, end_m=None, ro_no=None, part_no=None, ven_code=None, partial=None)
    print(f'Excuted query : {obj.query}')
    try:
        df = obj.search()
        export(df, 'search_test', interval=3000)
    except Exception as e:
        print(f'Data exporting failed. ({e})')
