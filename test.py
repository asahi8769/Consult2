
from datetime import datetime
from dateutil import relativedelta


nextmonth = datetime.today() + relativedelta.relativedelta(months=-1)
# print(nextmonth)

date_time_str = '202001'
beforemonth = datetime.strptime(date_time_str, '%Y%m') + relativedelta.relativedelta(months=-2)

print(beforemonth.strftime('%Y%m'))