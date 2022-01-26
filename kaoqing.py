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



def main():
    monthrange = get_monthrange(2021,12)
    workday_count = get_workdays(monthrange)
    holiday_count = get_holidays_count(monthrange)

    print()

if __name__ == '__main__':
    main()
    