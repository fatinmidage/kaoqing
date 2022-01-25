from datetime import datetime
from openpyxl import load_workbook
from chinese_calendar import is_workday, is_holiday
import datetime
import calendar

_c = calendar.Calendar()
date = _c.itermonthdates(2022,1)
for each in date:
    print(each)


today = datetime.date(2022,1,25)

print(is_workday(today))