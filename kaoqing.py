from email.mime import image
from openpyxl import load_workbook
from chinese_calendar import holidays, is_workday, is_holiday, is_in_lieu
import calendar


def get_monthrange(year,month):
    """获取指定月份所有日期的一个生成器

    Args:
        year (int)): 指定年份
        month (int): 指定月份

    Returns:
        generator: 所有日期的一个生成器
    """
    _c = calendar.Calendar()
    _month = _c.itermonthdates(year,month)
    monthrange = []
    [monthrange.append(each) for each in _month if each.month==month]
    return monthrange

def get_workdays(monthrange):
    """计算并返回指定月份的工作日数量

    Args:
        monthrange (generator): 月份生成器

    Returns:
        int: 工作日数量
    """
    workday_count = 0
    for each in monthrange:
        if is_workday(each):
            workday_count +=1
    return workday_count

def get_holidays_count(monthrange):
    """计算并返回指定月份的休息日和节假日总和

    Args:
        monthrange (generator): 指定月份的生成器

    Returns:
        int: 总节假日和休息日的数量
    """
    holiday_count = 0
    for each in monthrange:
        if is_holiday(each):
            holiday_count += 1
    return holiday_count

def get_six_workdays(monthrange):
    _count = 0
    for each in monthrange:
        if is_workday(each) or each.isoweekday() == 6:
            _count +=1
    return _count

def get_six_holidays(monthrange):
    holiday_count = 0
    for each in monthrange:
        if is_holiday(each) and each.isoweekday() != 6:
            holiday_count += 1
    return holiday_count

def get_legal_holidays_count(monthrange):
    _count = 0
    for each in monthrange:
        if is_holiday(each) and is_in_lieu(each):
            _count += 1
    return _count

def get_workday_info(year,month):
    _info = {}
    monthrange = get_monthrange(2021,12)
    _info['workday_count'] = get_workdays(monthrange)
    _info['holiday_count'] = get_holidays_count(monthrange)
    _info['legal_holiday'] = get_legal_holidays_count(monthrange)
    _info['six_workdays'] = get_six_workdays(monthrange)
    _info['six_holidays'] = get_six_holidays(monthrange)
    return _info

def update_basic_info(ws, year, month, info):
    ws['b1'] = month
    ws['b2'] = month if month<12 else 1
    print(ws['b2'].value)
    ws['b3'] = year
    ws['b5'] = info['workday_count']
    ws['b6'] = info['holiday_count']
    ws['b7'] = info['legal_holiday']
    ws['b8'] = info['six_workdays']
    ws['b9'] = info['six_holidays']







def main():
    year = 2022
    month = 1

    workday_info = get_workday_info(year,month)

    wb = load_workbook('考勤确认表.xlsx',data_only=True)

    # 更新基础信息表
    ws = wb['基础信息表']
    update_basic_info(ws,year,month,workday_info)

    ws_hengke = wb['恒科']
    ws_hengke['a2'] = '惠州恒科房地产开发有限公司' + str(year) + '年' + str(month) + '月份员工考勤确认统计表'
    ws_hengke['p3'] = '制表日期：' + str(2021) + '-' + str(month) + '-1'

    # 保存
    wb.save(str(month)+'月考勤确认表.xlsx')

if __name__ == '__main__':
    main()
    