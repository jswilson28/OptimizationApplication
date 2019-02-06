# This file contains general methods and classes used in other portions of the application.

from datetime import datetime, timedelta
from PyQt5.QtWidgets import QSpinBox, QComboBox, QLabel
from PyQt5.QtGui import QColor, QFont
import os


def day_from_num(num):

    if num < 1 or num > 7:
        return "NULL"

    day_list = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    return day_list[num-1]


def today():

    return datetime.today().strftime('%m%d%y')


def now():

    return datetime.today().strftime('%H%M%S')


def quick_italic():
    font = QFont()
    font.setItalic(True)
    return font


def quick_bold():
    font = QFont()
    font.setBold(True)
    return font


def quick_color(color):
    if color == "yellow":
        return QColor(255, 255, 0)
    elif color == "red":
        # return QColor(255, 145, 164)
        return QColor(246, 218, 176)
    elif color == "green":
        return QColor(51, 255, 51)
    elif color == "grey":
        return QColor(224, 224, 224)
    elif color == "blue":
        return QColor(0, 0, 255)
    elif color == "dark red":
        return QColor(255, 0, 0)
    else:
        return QColor(0, 0, 0)


def duration_between_two_times(start_time, stop_time):

    dt_start = time_to_datetime(start_time)
    dt_stop = time_to_datetime(stop_time)

    if dt_stop < dt_start:
        dt_stop += timedelta(hours=24)

    return int((dt_stop - dt_start).seconds/60)


def cminute_to_tod(cminute):

    cminute = cminute % 1440
    midnight = datetime(100, 1, 1, 0, 0, 0).time()
    return time_and_timedelta(midnight, cminute)


def minutes_to_hours_and_minutes(minutes):

    minutes = int(minutes)
    hours = int(minutes / 60)
    minutes = int(minutes % 60)
    return '{:,.0f}'.format(hours) + ":" + str(minutes).zfill(2)


def time_to_datetime(time):

    return datetime(100, 1, 1, time.hour, time.minute)


def time_and_timedelta(time, minutes):

    return (time_to_datetime(time) + timedelta(minutes=minutes)).time()


def night_differential(nd_am, nd_pm, start_time, duration):

    if not start_time or not duration:
        return 0

    if type(start_time) != int:
        start_minute = cminute(start_time)
    else:
        start_minute = start_time

    end_minute = start_minute + duration

    nd_am_minute = cminute(nd_am)
    nd_pm_minute = cminute(nd_pm)

    night_diff_minutes = 0

    for minute in range(start_minute, end_minute):
        day_of_minute = minute % 1440
        if 0 <= day_of_minute < nd_am_minute:
            night_diff_minutes += 1
        elif nd_pm_minute <= day_of_minute < 1440:
            night_diff_minutes += 1

    if night_diff_minutes > duration:
        print("Error!")
        return None

    return night_diff_minutes


def day_minutes(start_time, duration):
    # this returns the minutes that occur on day t and the number that occur on day t + 1, day t + 2
    midnight = datetime(100, 1, 1, 0, 0, 0).time()
    start_minute = int(duration_between_two_times(midnight, start_time))
    end_minute = int(start_minute + duration)

    today_minutes = 0
    tomorrow_minutes = 0
    next_day_minutes = 0

    for minute in range(start_minute, end_minute):
        if 0 <= minute < 1440:
            today_minutes += 1
        elif 1440 <= minute < 2880:
            tomorrow_minutes += 1
        elif minute >= 2880:
            next_day_minutes += 1

    if next_day_minutes > 0:
        print("Next day minutes! " + str(next_day_minutes))

    return today_minutes, tomorrow_minutes, next_day_minutes


def number(var):

    return '{:,.0f}'.format(float(var))


def number_decimal(var, decimals):

    format_string = '{:,.' + str(decimals) + 'f}'
    return format_string.format(float(var))


def number_too(var):

    if var.isdigit():
        return '{:,.0f}'.format(float(var))
    else:
        return str(var)


def money(money_in):
    if money_in >= 0:
        return '${:,.2f}'.format(money_in)
    else:
        return '-${:,.2f}'.format(-money_in)


def big_money(money_in):
    if money_in >= 0:
        return '${:,.0f}'.format(money_in)
    else:
        return '(${:,.0f})'.format(-money_in)


def percent(var):
    var = var*100
    return '{:,.1f}%'.format(var)


def qtime_to_time(time):

    time = time.toString()
    return datetime.strptime(time, "%H:%M:%S").time()


def cminute(time):

    midnight = datetime(100, 1, 1, 0, 0, 0).time()
    return duration_between_two_times(midnight, time)


def sun_minutes(start_time, duration):

    if not start_time or not duration:
        return 0

    minutes = 0

    for x in range(start_time, start_time + duration):
        if 7*1440 <= x < 8*1440 or 0 <= x < 1440:
            minutes += 1

    return minutes


class QuickSpin(QSpinBox):

    def __init__(self, min, max, val):
        super().__init__()
        self.setMinimum(min)
        self.setMaximum(max)
        self.setValue(val)


class YesNoCombo(QComboBox):

    def __init__(self):
        super().__init__()
        self.addItem("Yes")
        self.addItem("No")


class InputPasser:

    def __init__(self, **kwargs):

        source = kwargs.get('source')
        if not source:
            self.sensitive = False
        elif source == "InSorceror":
            self.sensitive = False
        elif source == "Optimizer":
            self.sensitive = True

        self.pvs_to_pdc = kwargs.get('pvs_to_pdc')
        if not self.pvs_to_pdc:
            self.pvs_to_pdc = 1

        self.pvs_time = kwargs.get('pvs_time')
        if not self.pvs_time:
            self.pvs_time = 14

        self.pdc_time = kwargs.get('pdc_time')
        if not self.pdc_time:
            self.pdc_time = 10

        self.tour_one_time = kwargs.get('tour_one_time')
        if not self.tour_one_time:
            self.tour_one_time = datetime(100, 1, 1, 20, 0, 0).time()

        self.lunch_duration = kwargs.get('lunch_duration')
        if not self.lunch_duration:
            self.lunch_duration = 30

        self.hours_wo_lunch = kwargs.get('hours_wo_lunch')
        if not self.hours_wo_lunch:
            self.hours_wo_lunch = 6

        self.lunch_travel_time = kwargs.get('lunch_travel_time')
        if not self.lunch_travel_time:
            self.lunch_travel_time = 5

        self.lunch_buffer_time = kwargs.get('lunch_buffer_time')
        if not self.lunch_buffer_time:
            self.lunch_buffer_time = 10

        self.max_duration = kwargs.get('max_duration')
        if not self.max_duration:
            self.max_duration = 8

        self.max_mileage = kwargs.get('max_mileage')
        if not self.max_mileage:
            self.max_mileage = 350

        # self.allow_non_postal = kwargs.get('allow_non_postal') == "Yes"
        self.allow_non_postal = kwargs.get('allow_non_postal')
        if not self.allow_non_postal:
            self.allow_non_postal = True

        # self.allow_extension = kwargs.get('allow_extension') == "Yes"
        self.allow_extension = kwargs.get('allow_extension')
        if not self.allow_extension:
            self.allow_extension = True

        self.outpath = kwargs.get('outpath')
        if not self.outpath:
            self.outpath = os.getcwd()

        self.nd_pm = kwargs.get('nd_pm')
        if not self.nd_pm:
            self.nd_pm = datetime(100, 1, 1, 18, 0, 0).time()

        self.nd_am = kwargs.get('nd_am')
        if not self.nd_am:
            self.nd_am = datetime(100, 1, 1, 6, 0, 0).time()

        self.min_check_time = kwargs.get('min_check_time')
        if not self.min_check_time:
            self.min_check_time = 60

        self.max_check_time = kwargs.get('max_check_time')
        if not self.max_check_time:
            self.max_check_time = 480

        self.max_combined_time = kwargs.get('max_combined_time')
        if not self.max_combined_time:
            self.max_combined_time = 1440

        self.round_time = kwargs.get('round_time')
        if not self.round_time:
            self.round_time = 15
