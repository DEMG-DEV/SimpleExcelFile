#!/usr/bin/env python
# _*_coding: utf-8_*_

from datetime import date, timedelta


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
        return d + dlt, d + dlt + timedelta(days=days)
