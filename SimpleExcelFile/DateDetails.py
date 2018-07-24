from datetime import date, timedelta
import calendar


class DateDetails():

    def __init__(self):
        pass

    def get_init_end_week(self, year, week, days):
        d = date(year, 1, 1)
        if d.weekday() > 3:
            d = d + timedelta(7 - d.weekday())
        else:
            d = d - timedelta(d.weekday())
        dlt = timedelta(days=(week - 1) * 7)
        init_week = d + dlt
        end_week = d + dlt + timedelta(days=days)
        return init_week, end_week

    def get_month_day_range(self, date):
        last_day = date.replace(day=calendar.monthrange(date.year, date.month)[1])
        return last_day
