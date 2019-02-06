# This file contains the basic logic for keeping and modifying schedules, after being read in by the reader classes.

from GeneralMethods import duration_between_two_times as dur
from GeneralMethods import time_and_timedelta, time_to_datetime, night_differential, day_minutes, cminute, sun_minutes
from datetime import timedelta, datetime
from LookupReader import create_adjust_code, AddFrequencyCode
import math
import copy


class Stop:

    def __init__(self, **kwargs):

        # basics
        self.arrive_time = kwargs.get('arrive_time')
        self.depart_time = kwargs.get('depart_time')
        self.stop_name = kwargs.get('stop_name')

        self.arrive_cminute = kwargs.get('arrive_cminute')
        self.depart_cminute = kwargs.get('depart_cminute')

        self.original_schedule_name = kwargs.get('original_schedule_name')
        self.nass_code = kwargs.get('nass_code')
        self.load_num = kwargs.get('load_num')
        self.vehicle_type = kwargs.get('vehicle_type')

        self.category = kwargs.get('category')

        self.lunch_after = False
        self.to_remove = False
        self.address = None
        self.post_office_location = False

    def add_address(self, address):

        self.address = address

        if "US POSTAL SERVICE" in address:
            self.post_office_location = True
        elif "P&DC" in self.stop_name:
            self.post_office_location = True
        elif "PVS" in self.stop_name:
            self.post_office_location = True

    def layover(self):

        return dur(self.arrive_time, self.depart_time)

    def shift_stop(self, minutes):

        self.arrive_time = time_and_timedelta(self.arrive_time, minutes)
        self.depart_time = time_and_timedelta(self.depart_time, minutes)

        if self.arrive_cminute:
            self.arrive_cminute += minutes
        if self.depart_cminute:
            self.depart_cminute += minutes

    def change_depart_time(self, minutes):
        self.depart_time = time_and_timedelta(self.depart_time, minutes)
        if self.depart_cminute:
            self.depart_cminute += minutes
        if not self.is_good_stop():
            print("PROBLEM WITH A STOP'S LAYOVER")

    def change_arrive_time(self, minutes):
        self.arrive_time = time_and_timedelta(self.arrive_time, minutes)
        if self.arrive_cminute:
            self.arrive_cminute += minutes
        if not self.is_good_stop():
            print("PROBLEM WITH A STOP'S LAYOVER")

    def is_good_stop(self):
        return self.layover() >= 1

    def is_post_office_location(self):
        name_list = ("P&DC", "P & DC", "PVS", "NDC")
        if any(word in self.stop_name for word in name_list):
            self.post_office_location = True

    def __str__(self):

        address_string = "Address unknown"
        if self.address:
            address_string = "Address known"

        return ("Stop: " + self.stop_name + ", arrive: " + str(self.arrive_time) + ", depart: " +
                str(self.depart_time) + ", " + address_string)


class RoundTrip:

    def __init__(self, **kwargs):

        pass


class Schedule:

    def __init__(self, **kwargs):
        # Source info variables
        self.source = kwargs.get('source')  # hcr, pvs, JDA, or optimizer
        self.source_type = kwargs.get('source_type')  # excel, pdf, or html
        self.source_file = kwargs.get('source_file')  # originating file name

        # Schedule Name
        self.schedule_name = kwargs.get('schedule_name')  # can be set later based on other details

        # Site Info
        self.short_name = kwargs.get('short_name')
        self.pdc_name = kwargs.get('pvs_pdc')
        self.pvs_name = kwargs.get('pvs_name')
        self.same_name = self.pvs_name == self.pdc_name

        # Stops
        self.stops = kwargs.get('stops')
        self.stops_original = copy.deepcopy(self.stops)
        self.stops_postalized = copy.deepcopy(self.stops)

        # Frequency Code
        self.freq_code = str(kwargs.get('freq_code')).zfill(4)
        self.freq_code_original = copy.copy(self.freq_code)
        self.annual_trips = kwargs.get('annual_trips')
        self.day_list = [None, None, None, None, None, None, None]
        self.bin_string = kwargs.get('bin_string')

        # Vehicle Details
        self.vehicle_type = kwargs.get('vehicle_type')  # specific vehicle type
        self.vehicle_category = kwargs.get('vehicle_category')  # 11-Ton/Single/Other
        self.operator_category = None  # MVO/TTO/Other

        # Flags
        self.flags = []
        self.postalized = False  # Do postalized stops meet postal rules
        self.eligible = False  # Insource eligible

        # Schedule Details
        self.daily_hours = None
        self.daily_paid_hours = None
        self.daily_lunch_hours = None

        self.daily_hours_postalized = None
        self.daily_paid_hours_postalized = None
        self.daily_lunch_hours_postalized = None

        self.weekly_hours = None
        self.weekly_paid_hours = None
        self.weekly_lunch_hours = None
        self.weekly_saturday_hours = None
        self.weekly_saturday_lunch_hours = None
        self.weekly_saturday_paid_hours = None
        self.weekly_sunday_hours = None
        self.weekly_sunday_lunch_hours = None
        self.weekly_sunday_paid_hours = None
        self.weekly_night_hours = None
        self.weekly_night_lunch_hours = None
        self.weekly_night_paid_hours = None
        self.weekly_holiday_hours = None

        self.weekly_hours_postalized = None
        self.weekly_paid_hours_postalized = None
        self.weekly_lunch_hours_postalized = None
        self.weekly_saturday_hours_postalized = None
        self.weekly_saturday_lunch_hours_postalized = None
        self.weekly_saturday_paid_hours_postalized = None
        self.weekly_sunday_hours_postalized = None
        self.weekly_sunday_lunch_hours_postalized = None
        self.weekly_sunday_paid_hours_postalized = None
        self.weekly_night_hours_postalized = None
        self.weekly_holiday_hours_postalized = None

        self.annual_hours = None
        self.annual_paid_hours = None
        self.annual_lunch_hours = None
        self.annual_saturday_hours = None
        self.annual_sunday_hours = None
        self.annual_night_hours = None
        self.annual_holiday_hours = None

        self.annual_hours_postalized = None
        self.annual_paid_hours_postalized = None
        self.annual_lunch_hours_postalized = None
        self.annual_saturday_hours_postalized = None
        self.annual_sunday_hours_postalized = None
        self.annual_night_hours_postalized = None
        self.annual_holiday_hours_postalized = None

    def add_flag(self, flag):

        self.flags.append(flag)

    def get_flags(self):

        return self.flags

    def duration_current(self):

        # note this method adds individual travel times and layovers rather than comparing the first
        # and last times of the schedule to account for schedules over 24 hours

        duration = 0
        for x, stop in enumerate(self.stops[:-1]):
            next_stop = self.stops[x + 1]
            duration += stop.layover()
            duration += dur(stop.depart_time, next_stop.arrive_time)

        duration += self.stops[-1].layover()

        return duration

    def duration_postalized(self):

        # note this method adds individual travel times and layovers rather than comparing the first
        # and last times of the schedule to account for schedules over 24 hours

        duration = 0
        for x, stop in enumerate(self.stops_postalized[:-1]):
            next_stop = self.stops[x + 1]
            duration += stop.layover()
            duration += dur(stop.depart_time, next_stop.arrive_time)

        duration += self.stops[-1].layover()

        return duration

    def postal_compliance_check(self, input_passer):

        pvs_time = input_passer.pvs_time
        pdc_time = input_passer.pdc_time
        pvs_to_pdc = input_passer.pvs_to_pdc
        layover_time = pvs_time + pdc_time + pvs_to_pdc

        pvs_name = self.pvs_name
        pdc_name = self.pdc_name
        max_working_time = input_passer.max_duration
        lunch_duration = input_passer.lunch_duration
        hours_wo_lunch = input_passer.hours_wo_lunch

        # check all layovers longer than one minute
        for stop in self.stops:
            if not stop.is_good_stop():
                return False

        # check all travel times at least one minute
        for x, stop in enumerate(self.stops[:-1]):
            next_stop = self.stops[x+1]
            travel_time = dur(stop.depart_time, next_stop.arrive_time)
            if travel_time < 1:
                return False

        # check the PVS and P&DC start situation
        start_check = False
        if self.same_name:
            if self.stops[0].stop_name == pvs_name:
                if self.stops[0].layover() == layover_time:
                    start_check = True
        else:
            if self.stops[0].stop_name == pvs_name:
                if self.stops[1].stop_name == pdc_name:
                    if self.stops[0].layover() >= pvs_time:
                        if self.stops[1].layover() >= pdc_time:
                            start_check = True

        if not start_check:
            return False

        # check stop PVS and P&DC
        stop_check = False
        if self.same_name:
            if self.stops[-1].stop_name == pvs_name:
                if self.stops[-1].layover() == layover_time:
                    stop_check = True
        else:
            if self.stops[-1].stop_name == pvs_name:
                if self.stops[-2].stop_name == pdc_name:
                    if self.stops[-1].layover() >= pvs_time:
                        if self.stops[-2].layover() >= pdc_time:
                            stop_check = True

        if not stop_check:
            return False

        # check lunch situation
        break_time = 0
        lunch_check = False
        duration = self.duration_postalized()
        if duration < hours_wo_lunch * 60:
            lunch_check = True
        elif duration <= lunch_duration + hours_wo_lunch * 120:
            if self.has_good_lunch(lunch_duration, hours_wo_lunch):
                lunch_check = True
                lunch_stop = next(x for x in self.stops if x.stop_name in ("LUNCH", "Lunch"))
                break_time += lunch_stop.layover()
        else:
            print("too long for lunch")

        if not lunch_check:
            return False

        work_time = duration - break_time
        if work_time > max_working_time * 60:
            return False

        return True

    def has_good_lunch(self, lunch_duration, hours_wo_lunch):

        lunch_stops = [x for x in self.stops if x.stop_name in ("LUNCH", "Lunch")]

        if len(lunch_stops) > 1:
            print("Multiple lunches found!!")
            return False

        if len(lunch_stops) < 1:
            return False

        lunch_stop = lunch_stops[0]

        if lunch_stop.layover() < lunch_duration:
            # print("lunch found, but too short")
            return False

        start_time = self.stops[0].arrive_time
        stop_time = self.stops[-1].depart_time

        if dur(start_time, lunch_stop.arrive_time) > hours_wo_lunch * 60:
            # print("lunch found, is too late")
            return False

        if dur(lunch_stop.depart_time, stop_time) > hours_wo_lunch * 60:
            # print("lunch found, is too early")
            return False

        return True

    def add_a_lunch(self, input_passer):

        hours_wo_lunch = input_passer.hours_wo_lunch
        lunch_duration = input_passer.lunch_duration
        allow_extension = input_passer.allow_extension
        allow_non_postal = input_passer.allow_non_postal
        lunch_travel_time = input_passer.lunch_travel_time
        lunch_buffer_time = input_passer.lunch_buffer_time
        total_buffer = lunch_travel_time + lunch_buffer_time

        duration = self.duration_current()
        if duration < hours_wo_lunch * 60:
            return True

        if self.has_good_lunch(lunch_duration, hours_wo_lunch):
            return True

        existing_lunch = [x for x in self.stops if x.stop_name in ("LUNCH", "Lunch")]
        if len(existing_lunch) > 0:
            # print("already tried best we could")
            return False

        start_time = self.stops[0].arrive_time
        stop_time = self.stops[-1].depart_time

        max_td = timedelta(hours=hours_wo_lunch)
        one_day = timedelta(hours=24)

        start_dt = time_to_datetime(start_time)
        stop_dt = time_to_datetime(stop_time)
        if stop_dt < start_dt:
            stop_dt += one_day

        earliest_possible_lunch_dt = stop_dt - max_td
        latest_possible_lunch_dt = start_dt + max_td

        if earliest_possible_lunch_dt > latest_possible_lunch_dt:
            print("Schedule " + self.schedule_name + " is too long for one lunch.")
            return False

        earliest_possible_time = earliest_possible_lunch_dt.time()
        latest_possible_time = latest_possible_lunch_dt.time()

        earliest_possible_minute = max(26, dur(start_time, earliest_possible_time)+total_buffer)
        latest_possible_minute = dur(start_time, latest_possible_time)-total_buffer

        first_start_time = self.stops[0].arrive_time

        postal_stops = []
        extend_stops = []
        non_postal_stops = []
        extend_non_postal_stops = []
        potential_stops = 0

        for x, stop in enumerate(self.stops):
            start_minute = dur(first_start_time, stop.arrive_time)
            stop_minute = dur(first_start_time, stop.depart_time)
            layover_minutes = stop_minute - start_minute

            if stop_minute > earliest_possible_minute:
                if start_minute < latest_possible_minute:
                    if stop.post_office_location:
                        if layover_minutes >= (lunch_duration + 2*(lunch_travel_time + lunch_buffer_time)):
                            postal_stops.append(x)
                            potential_stops += 1
                        elif allow_extension:
                            extend_stops.append(x)
                            potential_stops += 1
                    elif allow_non_postal:
                        if layover_minutes >= (lunch_duration + 2*(lunch_travel_time + lunch_buffer_time)):
                            non_postal_stops.append(x)
                            potential_stops += 1
                        elif allow_extension:
                            extend_non_postal_stops.append(x)
                            potential_stops += 1

        if potential_stops == 0:
            # print("No eligible lunch stops found")
            return False
        # else:
            # print("Found " + str(potential_stops) + " eligible stops for lunch")

        if len(postal_stops) > 0:

            index = postal_stops[0]
            stop = self.stops[index]
            must_start_by = latest_possible_minute - dur(start_time, stop.arrive_time) + total_buffer
            # cant_finish_before = earliest_possible_minute - dur(start_time, stop.arrive_time)
            self.insert_lunch(index, lunch_duration, must_start_by, lunch_travel_time, lunch_buffer_time)

        elif len(non_postal_stops) > 0:
            index = non_postal_stops[0]
            stop = self.stops[index]
            must_start_by = latest_possible_minute - dur(start_time, stop.arrive_time) + total_buffer
            # cant_finish_before = earliest_possible_minute - dur(start_time, stop.arrive_time)
            self.insert_lunch(index, lunch_duration, must_start_by, lunch_travel_time, lunch_buffer_time)

        elif len(extend_stops) > 0:
            index = extend_stops[0]
            for x in extend_stops:
                if self.stops[x].layover() > self.stops[index].layover():
                    index = x

            orig_layover = self.stops[index].layover()
            required_layover = lunch_duration + 2*(lunch_travel_time + lunch_buffer_time)
            pushback_minutes = required_layover - orig_layover

            self.stops[index].change_depart_time(pushback_minutes)

            stop = self.stops[index]
            must_start_by = latest_possible_minute - dur(start_time, stop.arrive_time) + total_buffer
            # cant_finish_before = earliest_possible_minute - dur(start_time, stop.arrive_time)

            self.insert_lunch(index, lunch_duration, must_start_by, lunch_travel_time, lunch_buffer_time)

            for stop in self.stops[index+3:]:
                stop.shift_stop(pushback_minutes)
            for stop in self.stops:
                stop.shift_stop(-round(pushback_minutes/2, 0))

        elif len(extend_non_postal_stops) > 0:
            index = extend_non_postal_stops[0]

            for x in extend_non_postal_stops:
                if self.stops[x].layover() > self.stops[index].layover():
                    index = x

            orig_layover = self.stops[index].layover()
            required_layover = lunch_duration + 2*(lunch_travel_time + lunch_buffer_time)
            pushback_minutes = required_layover - orig_layover

            self.stops[index].change_depart_time(pushback_minutes)

            stop = self.stops[index]
            must_start_by = latest_possible_minute - dur(start_time, stop.arrive_time) + total_buffer
            # cant_finish_before = earliest_possible_minute - dur(start_time, stop.arrive_time)

            self.insert_lunch(index, lunch_duration, must_start_by, lunch_travel_time, lunch_buffer_time)

            for stop in self.stops[index+3:]:
                stop.shift_stop(pushback_minutes)
            for stop in self.stops:
                stop.shift_stop(-round(pushback_minutes/2, 0))

        else:
            print("wtf? eligible lunch stops both found and not found??")

        return True

    def insert_lunch(self, index, lunch_duration, must_start_by, lunch_travel_time, lunch_buffer_time):

        stop = self.stops[index]
        start_range_stop = must_start_by

        arrive_lunch_minute = min(start_range_stop, stop.layover()-(lunch_duration +
                                                                    lunch_travel_time + lunch_buffer_time))

        arrive_lunch_time = time_and_timedelta(stop.arrive_time, arrive_lunch_minute)
        new_depart_time = time_and_timedelta(arrive_lunch_time, -lunch_travel_time)
        depart_lunch_time = time_and_timedelta(arrive_lunch_time, lunch_duration)
        arrive_back_time = time_and_timedelta(depart_lunch_time, lunch_travel_time)

        lunch_stop = Stop(arrive_time=arrive_lunch_time, depart_time=depart_lunch_time, stop_name="LUNCH")
        new_back_stop = Stop(arrive_time=arrive_back_time, depart_time=stop.depart_time, stop_name=stop.stop_name)

        self.stops[index].depart_time = new_depart_time
        self.stops.insert(index + 1, new_back_stop)
        self.stops.insert(index + 1, lunch_stop)