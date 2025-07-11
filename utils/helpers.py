import calendar

def getMonthRange(month):
    month_start = f'2024-{month}-01'
    month_end = f'2024-{month}-{calendar.monthrange(2024, int(month))[1]}'
    return month_start, month_end