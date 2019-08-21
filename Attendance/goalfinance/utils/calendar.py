from calendar import monthrange
from calendar import weekday

"""
Return a list of a tuple, the structure of the tuple is (weekday, day, isHoliday)
"""
def get_calendar_of_month(year, month):
    calendar_of_month = []
    days_of_month = monthrange(year, month)[1]
    for i in range(1, days_of_month + 1):
        calendar_info = ()
        week_day = weekday(year, month, i)
        isHoliday = False
        if week_day in [5, 6]:
            isHoliday = True
        calendar_info = week_day, i, isHoliday
        calendar_of_month.append(calendar_info)
    return calendar_of_month

def month_to_str(month):
    month_str = str(month)
    if len(month_str) < 2:
        month_str = '0' + month_str 
    return month_str
