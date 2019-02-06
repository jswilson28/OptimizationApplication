# This file contains the Schedule, Round-trip, and Stop classes.
from GeneralMethods import duration_between_two_times as dur
from GeneralMethods import time_and_timedelta, time_to_datetime, night_differential, day_minutes, cminute, sun_minutes
from datetime import timedelta, datetime
import copy
from LookupReader import create_adjust_code, AddFrequencyCode
import math


class Stop:

    def __init__(self, **kwargs):

        self.arrive_time = kwargs.get('arrive_time')
        self.depart_time = kwargs.get('depart_time')
        self.stop_name = kwargs.get('stop_name')

        self.base_time = False
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

        # if not self.is_good_stop():
            # print("PROBLEM WITH A STOP'S LAYOVER")

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

        if "P&DC" in self.stop_name:
            self.post_office_location = True
        elif "PVS" in self.stop_name:
            self.post_office_location = True

    def __str__(self):

        address_string = "Address unknown"
        if self.address:
            address_string = "Address known"

        return ("Stop: " + self.stop_name + ", arrive: " + str(self.arrive_time) + ", depart: " +
                str(self.depart_time) + ", " + address_string)


class Schedule:

    def __init__(self, **kwargs):
        # source type is the file type of the source information: Excel, PDF, or HTML
        self.source_type = kwargs.get('source_type')

        # source is whether the schedule came from an HCR contract, a PVS document, or a PVS Process (optimization)
        self.source = kwargs.get('source')
        self.source_file = kwargs.get('source_file')

        # HCR info
        self.part = kwargs.get('part')
        self.plate_num = kwargs.get('plate_num')
        self.purpose = kwargs.get('purpose')
        self.mail_class = kwargs.get('mail_class')

        short_plate_num = None
        if self.plate_num:
            short_plate_num = self.plate_num[:5]

        if not self.source_file:
            if short_plate_num:
                self.source_file = short_plate_num

        self.trip_num = kwargs.get('trip_num')

        # PVS info
        self.schedule_num = kwargs.get('schedule_num')
        self.run_num = kwargs.get('run_num')
        self.tractor = kwargs.get('tractor')

        # Opt info
        self.read_in_index = kwargs.get('read_in_index')
        self.num_trips = kwargs.get('num_trips')

        # PVS other info
        self.annual_miles = None
        self.annual_hours = None
        self.weekday_hours = None
        self.saturday_hours = None
        self.sunday_hours = None
        self.holiday_hours = None
        self.calculated_night_hours = None
        self.postalized_night_hours = None
        self.paid_time = None
        self.unassigned_time_listed = None
        self.unassigned_time_calculated = None
        self.schedule_time = None

        # Name info
        self.hcr_pdc = kwargs.get('hcr_pdc')
        self.pvs_pdc = kwargs.get('pvs_pdc')
        self.pvs_name = kwargs.get('pvs_name')
        self.short_name = kwargs.get('short_name')
        if not self.short_name:
            self.short_name = self.pvs_name
        self.same_name = self.pvs_name == self.pvs_pdc

        # Frequency code info
        self.adjust = False
        self.freq_code = kwargs.get('freq_code')
        if not self.freq_code:
            print("FREQUENCY CODE MISSING!")

        self.freq_code.zfill(4)
        self.original_freq_code = str(self.freq_code)

        self.annual_trips = kwargs.get('annual_trips')

        self.freq_code_description = kwargs.get('freq_code_description')

        self.stops = kwargs.get('stops')
        self.original_duration = self.raw_duration()
        self.postalized_duration = 0
        self.original_stops = copy.deepcopy(self.stops)
        self.postalized_stops = copy.deepcopy(self.stops)
        self.pdc_address = kwargs.get('pdc_address')
        self.pdc_tz = kwargs.get('pdc_tz')

        # Vehicle Info
        self.vehicle_type = kwargs.get('vehicle_type')
        if not self.vehicle_type:
            self.vehicle_type = "Unknown Vehicle Type"
        self.vehicle_category = kwargs.get('vehicle_category')
        # self.opt_schedule_num = None
        self.schedule_name = kwargs.get('schedule_name')
        if not self.schedule_name:
            if self.source == "HCR":
                self.schedule_name = str(self.plate_num) + " " + str(self.trip_num)
            elif self.source == "PVS":
                self.schedule_name = self.pvs_pdc + " " + str(self.schedule_num)
            elif self.source == "OPT":
                # self.opt_schedule_num = self.schedule_num
                self.set_new_schedule_name()
            elif self.source == "JDA":
                self.schedule_name = self.short_name + " " + str(self.schedule_num)
            else:
                self.schedule_name = "UK Source " + str(self.schedule_num)

        self.mileage = kwargs.get('mileage')
        if not self.mileage:
            self.mileage = 0
        else:
            self.mileage = float(self.mileage)

        self.original_mileage = copy.copy(float(self.mileage))
        self.postalized_mileage = 0
        self.effective_date = kwargs.get('effective_date')
        self.end_date = kwargs.get('end_date')

        self.flags = []
        self.tour = kwargs.get('tour')

        self.mon = None
        self.tue = None
        self.wed = None
        self.thu = None
        self.fri = None
        self.sat = None
        self.sun = None
        self.bin_string = kwargs.get('bin_string')

        if self.bin_string:
            self.set_days_of_week()

        self.first_stop_depart = self.stops[0].depart_time
        self.last_stop_arrive = self.stops[-1].arrive_time

        self.annual_calculated_duration = None
        self.annual_postalized_duration = None
        self.annual_calculated_mileage = None
        self.annual_postalized_mileage = None

        # this keeps track of how many minutes are added by rounding to the 15s
        self.fluff_minutes = 0

        self.is_postalized = False
        self.can_postalize = False
        self.is_eligible = False
        self.tried_to_postalize = False
        self.has_been_postalized = False
        self.cant_eligible = []
        self.cant_postalize = []
        self.is_spotter_schedule = False
        self.pre_cleaned_stops = None
        self.pre_cleaned_duration = None

        self.round_trips = []
        # self.is_spotter_schedule = False
        self.schedule_type = None
        self.holiday = False
        self.network_schedule = False
        self.cross_state_lines = False

    def set_days_of_week(self):

        self.mon = int(self.bin_string[0]) == 1
        self.tue = int(self.bin_string[1]) == 1
        self.wed = int(self.bin_string[2]) == 1
        self.thu = int(self.bin_string[3]) == 1
        self.fri = int(self.bin_string[4]) == 1
        self.sat = int(self.bin_string[5]) == 1
        self.sun = int(self.bin_string[6]) == 1

    def add_flag(self, flag):

        self.flags.append(flag)

    def raw_duration(self):

        duration = 0

        for x, stop in enumerate(self.stops[:-1]):
            next_stop = self.stops[x+1]
            duration += stop.layover()
            duration += dur(stop.depart_time, next_stop.arrive_time)

        duration += self.stops[-1].layover()
        # old_duration = dur(self.stops[0].arrive_time, self.stops[-1].depart_time)

        return duration

    def postal_compliance_check(self, input_passer):
        # return true iff schedule is already postal compliant
        # print("Checking postal compliance of: " + self.schedule_name)

        pvs_time = input_passer.pvs_time
        pdc_time = input_passer.pdc_time
        pvs_to_pdc = input_passer.pvs_to_pdc
        layover_time = pvs_time + pdc_time + pvs_to_pdc

        pvs_name = self.pvs_name
        pdc_names = [self.hcr_pdc, self.pvs_pdc]
        max_working_time = input_passer.max_duration
        lunch_duration = input_passer.lunch_duration
        hours_wo_lunch = input_passer.hours_wo_lunch

        # check all layovers longer than one minute
        for stop in self.stops:
            if not stop.is_good_stop():
                # print("Found stop with layover under one minute")
                return False

        # check all travel times at least one minute
        for x, stop in enumerate(self.stops[0:-1]):
            next_stop = self.stops[x+1]
            travel_time = dur(stop.depart_time, next_stop.arrive_time)
            if travel_time < 1:
                # print(self.schedule_name, ": found a travel time under one minute")
                # print(stop.depart_time, next_stop.arrive_time)
                return False

        # check the PVS and P&DC start situation
        start_check = False
        if self.same_name:
            if self.stops[0].stop_name == pvs_name:
                if self.stops[0].layover() == layover_time:
                    start_check = True
        else:
            if self.stops[0].stop_name == pvs_name:
                if self.stops[1].stop_name in pdc_names:
                    if self.stops[0].layover() >= pvs_time:
                        if self.stops[1].layover() >= pdc_time:
                            start_check = True

        if not start_check:
            # print("does not start at PVS/P&DC correctly")
            return False

        # check stop PVS and P&DC
        stop_check = False
        if self.same_name:
            if self.stops[-1].stop_name == pvs_name:
                if self.stops[-1].layover() == layover_time:
                    stop_check = True
        else:
            if self.stops[-1].stop_name == pvs_name:
                if self.stops[-2].stop_name in pdc_names:
                    if self.stops[-1].layover() >= pvs_time:
                        if self.stops[-2].layover() >= pdc_time:
                            stop_check = True

        if not stop_check:
            # print("does not terminate at PVS/P&DC correctly")
            return False

        # check lunch situation
        break_time = 0
        lunch_check = False
        duration = self.raw_duration()
        if duration < hours_wo_lunch*60:
            lunch_check = True
        elif duration <= lunch_duration + hours_wo_lunch*120:
            if self.has_good_lunch(lunch_duration, hours_wo_lunch):
                lunch_check = True
                lunch_stop = next(x for x in self.stops if x.stop_name in ("LUNCH", "Lunch"))
                break_time += lunch_stop.layover()
        else:
            print("too long for lunch")

        if not lunch_check:
            return False

        # lastly check duration
        work_time = duration - break_time
        if work_time > max_working_time*60:
            # print("exceeds maximum working duration")
            return False

        # print("Is postal compliant")
        return True

    def postal_compliance_possible(self, input_passer, address_book):

        if self.is_postalized:
            return True

        if self.tried_to_postalize:
            return self.can_postalize

        max_duration = input_passer.max_duration * 60
        hours_wo_lunch = input_passer.hours_wo_lunch
        lunch_duration = input_passer.lunch_duration

        self.add_pvs_and_pdc(input_passer, address_book)

        duration = self.raw_duration()

        lunch_needed = duration > hours_wo_lunch*60

        if lunch_needed:
            if not self.has_good_lunch(lunch_duration, hours_wo_lunch):
                if self.add_a_lunch(input_passer):
                    duration = self.raw_duration()
                    max_duration += lunch_duration
                    lunch_check = True
                else:
                    lunch_check = False

            else:
                lunch_check = True
                lunch_stop = next(x for x in self.stops if x.stop_name in ("LUNCH", "Lunch", "lunch"))
                max_duration += lunch_stop.layover()
        else:
            lunch_check = True

        duration_check = duration <= max_duration

        # reset stops
        self.stops = copy.deepcopy(self.original_stops)
        # self.mileage = copy.copy(self.original_mileage)

        if duration_check and lunch_check:
            return True
        if not lunch_check:
            self.cant_postalize.append("Lunch Impossible")
        if not duration_check:
            self.cant_postalize.append("Duration")

        return False

    def add_a_lunch(self, input_passer):

        hours_wo_lunch = input_passer.hours_wo_lunch
        lunch_duration = input_passer.lunch_duration
        allow_extension = input_passer.allow_extension
        allow_non_postal = input_passer.allow_non_postal
        lunch_travel_time = input_passer.lunch_travel_time
        lunch_buffer_time = input_passer.lunch_buffer_time
        total_buffer = lunch_travel_time + lunch_buffer_time

        duration = self.raw_duration()

        if duration < hours_wo_lunch * 60:
            # print("tried to add lunch to schedule that doesn't need one!")
            return True

        if self.has_good_lunch(lunch_duration, hours_wo_lunch):
            # print("tried to add lunch to schedule that already had one!")
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
        if self.same_name:
            if self.stops[-1].depart_time == new_back_stop.depart_time:
                new_end_time = time_and_timedelta(new_back_stop.depart_time, 15)
                self.stops[-1].depart_time = new_end_time

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

        if dur(start_time, lunch_stop.arrive_time) > hours_wo_lunch*60:
            # print("lunch found, is too late")
            return False

        if dur(lunch_stop.depart_time, stop_time) > hours_wo_lunch*60:
            # print("lunch found, is too early")
            return False

        return True

    def trim_start_and_stop(self):

        names = (self.hcr_pdc, self.pvs_pdc, self.pvs_name)

        test_stops = [x for x in self.stops if x.stop_name in names]
        if len(test_stops) == len(self.stops):
            return False

        while self.stops[0].stop_name in names:
            self.first_stop_depart = self.stops[0].depart_time
            del self.stops[0]

        while self.stops[-1].stop_name in names:
            self.last_stop_arrive = self.stops[-1].arrive_time
            del self.stops[-1]

        return True

    def add_start_and_stop(self, input_passer):

        pdc_time = input_passer.pdc_time
        pvs_time = input_passer.pvs_time
        pvs_to_pdc = input_passer.pvs_to_pdc
        layover_time = pdc_time + pvs_time + pvs_to_pdc

        depart_time = self.first_stop_depart

        if self.same_name:
            new_arrive_time = time_and_timedelta(depart_time, -layover_time)
            start_stop = Stop(arrive_time=new_arrive_time, depart_time=depart_time, stop_name=self.pvs_name)
            if self.pdc_address:
                start_stop.add_address(self.pdc_address)

            self.stops.insert(0, start_stop)

        else:
            pdc_depart = depart_time
            pdc_arrive = time_and_timedelta(depart_time, -pdc_time)
            pvs_depart = time_and_timedelta(pdc_arrive, -pvs_to_pdc)
            pvs_arrive = time_and_timedelta(pvs_depart, -pvs_time)

            pdc_stop = Stop(arrive_time=pdc_arrive, depart_time=pdc_depart, stop_name=self.pvs_pdc)
            if self.pdc_address:
                pdc_stop.add_address(self.pdc_address)

            pvs_stop = Stop(arrive_time=pvs_arrive, depart_time=pvs_depart, stop_name=self.pvs_name)

            self.stops.insert(0, pdc_stop)
            self.stops.insert(0, pvs_stop)

        arrive_time = self.last_stop_arrive

        if self.same_name:
            new_depart_time = time_and_timedelta(arrive_time, layover_time)
            stop_stop = Stop(arrive_time=arrive_time, depart_time=new_depart_time, stop_name=self.pvs_name)
            if self.pdc_address:
                stop_stop.add_address(self.pdc_address)

            self.stops.append(stop_stop)

        else:
            pdc_arrive = arrive_time
            pdc_depart = time_and_timedelta(pdc_arrive, pdc_time)
            pvs_arrive = time_and_timedelta(pdc_depart, pvs_to_pdc)
            pvs_depart = time_and_timedelta(pvs_arrive, pvs_time)

            pdc_stop = Stop(arrive_time=pdc_arrive, depart_time=pdc_depart, stop_name=self.pvs_pdc)
            if self.pdc_address:
                pdc_stop.add_address(self.pdc_address)

            pvs_stop = Stop(arrive_time=pvs_arrive, depart_time=pvs_depart, stop_name=self.pvs_name)

            self.stops.append(pdc_stop)
            self.stops.append(pvs_stop)

    def add_pvs_and_pdc(self, input_passer, address_book):

        if not self.trim_start_and_stop():
            self.is_spotter_schedule = True
            return

        if self.first_stop_depart == self.stops[0].depart_time:
            if address_book.check_existing_trips(self.pvs_pdc, self.stops[0].stop_name):
                travel_time = address_book.get_a_trip(self.pvs_pdc, self.stops[0].stop_name)
                distance = travel_time[2]
                travel_time = travel_time[3]
            else:
                travel_time, distance = address_book.add_a_trip(self.pvs_pdc, self.stops[0].stop_name,
                                                                self.pdc_address, self.stops[0].address)

            first_stop_arrive_time = self.stops[0].arrive_time
            self.first_stop_depart = time_and_timedelta(first_stop_arrive_time, -travel_time)
            self.mileage += distance/1609

        if self.last_stop_arrive == self.stops[-1].arrive_time:
            if address_book.check_existing_trips(self.stops[-1].stop_name, self.pvs_pdc):
                travel_time = address_book.get_a_trip(self.stops[-1].stop_name, self.pvs_pdc)
                distance = travel_time[2]
                travel_time = travel_time[3]
            else:
                travel_time, distance = address_book.add_a_trip(self.stops[-1].stop_name, self.pvs_pdc,
                                                                self.stops[-1].address, self.pdc_address)

            last_stop_depart_time = self.stops[-1].depart_time
            self.last_stop_arrive = time_and_timedelta(last_stop_depart_time, travel_time)
            self.mileage += distance/1609

        self.add_start_and_stop(input_passer)

    def convert_to_pvs(self):

        for x, stop in enumerate(self.stops):
            if stop.stop_name == self.hcr_pdc:
                self.stops[x].stop_name = self.pvs_pdc

    def sps_print(self, ws, row, column):

        print_list = []
        row_iter = row
        column_start = column
        pvs_name = self.pvs_name
        freq = self.freq_code

        if not self.effective_date:
            schedule_effective_date = ""
        else:
            schedule_effective_date = self.effective_date

        if not self.end_date:
            schedule_end_date = ""
        else:
            schedule_end_date = self.end_date

        schedule_run_number = ""
        transporter_type = self.vehicle_category

        for x, stop in enumerate(self.stops):
            print_list.append(pvs_name)
            print_list.append(self.schedule_name)
            print_list.append(schedule_effective_date)
            print_list.append(schedule_end_date)
            print_list.append(freq)
            print_list.append(schedule_run_number)
            # if stop.vehicle_type:
            #    print_list.append(stop.vehicle_type)
            # else:
            print_list.append(transporter_type)
            print_list.append(str(x + 1))

            if x == 0:
                print_list.append("")
                print_list.append("")
            else:
                previous_stop = self.stops[x - 1]
                print_list.append(previous_stop.stop_name)
                print_list.append(str(previous_stop.depart_time))
            print_list.append(str(stop.arrive_time))
            print_list.append(stop.stop_name)
            print_list.append(str(stop.depart_time))

            if x == len(self.stops) - 1:
                print_list.append("")
                print_list.append("")
                print_list.append("")
            else:
                next_stop = self.stops[x + 1]
                print_list.append(str(next_stop.arrive_time))
                print_list.append(next_stop.stop_name)
                print_list.append(str(next_stop.depart_time))
            print_list.append("")
            print_list.append("")
            print_list.append("")

            print_list.append(self.mileage)

            if len(self.flags) == 0:
                print_list.append("No flags")
            else:
                for line in self.flags:
                    print_list.append(line)
            for y, item in enumerate(print_list):
                ws.cell(row=row_iter, column=column_start + y).value = item

            row_iter = row_iter + 1
            print_list = []

    def add_hcr_freq_code_info(self, annual_trips, description):

        self.annual_trips = annual_trips
        self.freq_code_description = description

    def add_html_info(self, annual_miles, annual_hours, weekday_hours, sat_hours, sun_hours, hol_hours, night_diff,
                      paid_time, unassigned_time, lb):

        print(self.schedule_name)
        self.annual_miles = float(annual_miles)
        self.annual_hours = annual_hours
        self.weekday_hours = weekday_hours
        self.saturday_hours = sat_hours
        self.sunday_hours = sun_hours
        self.holiday_hours = hol_hours
        self.calculated_night_hours = night_diff
        self.paid_time = paid_time
        self.unassigned_time_listed = unassigned_time
        self.schedule_time = self.raw_duration()

        if self.annual_trips in (0, "0", None):
            if self.mileage > 0:
                self.annual_trips = self.annual_miles/self.mileage
            else:
                self.annual_trips = float(lb.find_code_info(self.freq_code)[1])

    def calc_other_info(self, night_diff):

        if not night_diff:
            time1 = datetime(100, 1, 1, 6, 0, 0).time()
            time2 = datetime(100, 1, 1, 18, 0, 0).time()
            night_diff = [time1, time2]

        if self.annual_trips:
            self.annual_miles = float(self.mileage) * float(self.annual_trips)
        else:
            self.annual_miles = 0
            self.annual_trips = 0

        if not self.paid_time:
            lunch = [x for x in self.stops if x.stop_name in ("LUNCH", "Lunch")]
            if lunch:
                self.paid_time = self.raw_duration() - lunch[0].layover()
            else:
                self.paid_time = self.raw_duration()

        if not self.annual_hours:
            self.annual_hours = self.paid_time * float(self.annual_trips)

        start_time = self.original_stops[0].arrive_time
        duration = self.original_duration

        night_minutes = night_differential(night_diff[0], night_diff[1], start_time, duration)
        if not self.calculated_night_hours:
            self.calculated_night_hours = (night_minutes * float(self.annual_trips)) / 60

        start_time = self.postalized_stops[0].arrive_time
        duration = self.postalized_duration

        night_minutes = night_differential(night_diff[0], night_diff[1], start_time, duration)
        self.postalized_night_hours = (night_minutes * float(self.annual_trips)) / 60

        minutes, next_day_minutes, following_day_minutes = day_minutes(start_time, duration)
        weekday_minutes = 0
        saturday_minutes = 0
        sunday_minutes = 0

        if self.mon:
            weekday_minutes += minutes + next_day_minutes + following_day_minutes
        if self.tue:
            weekday_minutes += minutes + next_day_minutes + following_day_minutes
        if self.wed:
            weekday_minutes += minutes + next_day_minutes + following_day_minutes
        if self.thu:
            weekday_minutes += minutes + next_day_minutes
            saturday_minutes += following_day_minutes
        if self.fri:
            weekday_minutes += minutes
            saturday_minutes += next_day_minutes
            sunday_minutes += following_day_minutes
        if self.sat:
            saturday_minutes += minutes
            sunday_minutes += next_day_minutes
            weekday_minutes += following_day_minutes
        if self.sun:
            sunday_minutes += minutes
            weekday_minutes += next_day_minutes + following_day_minutes

        # if not self.weekday_hours:
        self.weekday_hours = weekday_minutes/60
        # if not self.saturday_hours:
        self.saturday_hours = saturday_minutes/60
        # if not self.sunday_hours:
        self.sunday_hours = sunday_minutes/60

        self.unassigned_time_calculated = 0
        check_list = ["Standby", "STANDBY TIME", "STANDBY", "ASSIGNED TO OTHER DUTIES"]
        standby_stops = [stop for stop in self.stops if any(word in stop.stop_name for word in check_list)]
        if standby_stops:
            self.unassigned_time_calculated = sum(stop.layover() for stop in standby_stops)*self.annual_trips/60
        if self.unassigned_time_listed:
            self.unassigned_time_listed = self.unassigned_time_listed*self.annual_trips

    def insource_eligible_check(self, input_passer):

        eligible_mileage = input_passer.max_mileage
        max_duration = input_passer.max_duration

        duration_check = self.original_duration/60 <= max_duration
        mileage_check = self.original_mileage <= eligible_mileage

        if duration_check and mileage_check:
            return True
        if not duration_check:
            self.cant_eligible.append("Duration")
        if not mileage_check:
            self.cant_eligible.append("Mileage")
        return False

    def postalize(self, ip, address_book):

        if self.has_been_postalized:
            return

        if self.source == "JDA":
            self.multiple_load_fix(address_book)
        self.convert_to_pvs()
        self.correct_travel_times()
        self.add_pvs_and_pdc(ip, address_book)
        self.add_a_lunch(ip)
        self.set_cminutes(ip)
        self.round_tripify(ip)
        # add check for last stop here
        self.is_postalized = self.postal_compliance_check(ip)
        self.tried_to_postalize = True
        self.postalized_stops = copy.deepcopy(self.stops)
        self.postalized_duration = self.raw_duration()
        self.calc_other_info(None)
        # self.is_spotter_schedule = self.is_a_spotter_schedule(None)
        self.find_schedule_type()
        self.annual_postalized_duration = (self.postalized_duration * float(self.annual_trips)) / 60
        self.annual_calculated_duration = (self.original_duration * float(self.annual_trips)) / 60
        self.postalized_mileage = copy.copy(self.mileage)
        self.annual_postalized_mileage = self.postalized_mileage * float(self.annual_trips)
        self.annual_calculated_mileage = self.original_mileage * float(self.annual_trips)
        self.has_been_postalized = True

    def short_print_original(self, ws, row_start):

        row = row_start

        # num_to_print = self.trip_num
        # if not num_to_print:
        #     num_to_print = self.schedule_num

        for x, stop in enumerate(self.original_stops):
            ws["A" + str(row)].value = self.schedule_name
            ws["B" + str(row)].value = str(x + 1)
            ws["C" + str(row)].value = str(stop.arrive_time)
            ws["D" + str(row)].value = str(stop.depart_time)
            ws["E" + str(row)].value = stop.stop_name
            ws["F" + str(row)].value = stop.nass_code
            row += 1

    def short_print_postalized(self, ws, row_start):

        row = row_start

        # num_to_print = self.trip_num
        # if not num_to_print:
        #     num_to_print = self.schedule_num

        for x, stop in enumerate(self.postalized_stops):
            ws["A" + str(row)].value = self.schedule_name
            ws["B" + str(row)].value = str(x + 1)
            ws["C" + str(row)].value = str(stop.arrive_time)
            ws["D" + str(row)].value = str(stop.depart_time)
            ws["E" + str(row)].value = stop.stop_name
            ws["F" + str(row)].value = stop.nass_code
            row += 1

    def short_print_summary(self, ws, row):

        num_to_print = self.trip_num
        if not num_to_print:
            num_to_print = self.schedule_num

        plate_to_print = self.plate_num
        if not plate_to_print:
            plate_to_print = self.source_file

        ws["A" + str(row)].value = self.short_name
        ws["B" + str(row)].value = plate_to_print
        ws["C" + str(row)].value = num_to_print
        ws["D" + str(row)].value = str(self.is_postalized)
        ws["E" + str(row)].value = str(self.network_schedule)
        ws["F" + str(row)].value = str(not self.cross_state_lines)
        ws["G" + str(row)].value = self.freq_code
        ws["H" + str(row)].value = self.annual_trips
        ws["I" + str(row)].value = self.vehicle_category

        ws["J" + str(row)].value = self.original_mileage
        ws["K" + str(row)].value = self.postalized_mileage - self.original_mileage
        ws["L" + str(row)].value = self.postalized_mileage
        ws["M" + str(row)].value = self.annual_postalized_mileage

        ws["N" + str(row)].value = self.original_duration/60
        ws["O" + str(row)].value = self.postalized_duration/60
        ws["P" + str(row)].value = self.annual_postalized_duration
        ws["Q" + str(row)].value = self.sunday_hours * 51.61
        ws["R" + str(row)].value = self.postalized_night_hours
        ws["S" + str(row)].value = self.unassigned_time_listed
        ws["T" + str(row)].value = self.unassigned_time_calculated

    def round_tripify(self, ip):

        self.round_trips = []
        trimmed_stops = []
        new_stops = []
        for x, stop in enumerate(self.stops):
            if stop.stop_name == self.pvs_name and x in (0, len(self.stops)-1) and not self.same_name:
                continue
            if stop.stop_name in (self.hcr_pdc, self.pvs_pdc):
                if len(new_stops) == 0:
                    new_stops.append(copy.deepcopy(stop))
                else:
                    new_stops.append(copy.deepcopy(stop))
                    trimmed_stops.append(new_stops)
                    new_stops = [copy.deepcopy(stop)]
            else:
                new_stops.append(copy.deepcopy(stop))

        for x, row in enumerate(trimmed_stops):
            self.round_trips.append(RoundTrip(self.short_name, self.source_file, self.schedule_name, row,
                                              self.vehicle_category, self.freq_code, self.pvs_pdc, self.bin_string,
                                              self.holiday, x+1, ip.hours_wo_lunch, self.pvs_name))

    def find_schedule_type(self):

        # 1 = real trip, 2 = lunch trip, 3 = spotter trip, 4 = standby trip

        trip_types = [x.trip_type for x in self.round_trips]

        if 1 in trip_types:
            self.schedule_type = 1
        elif 3 in trip_types:
            self.schedule_type = 3
        else:
            self.schedule_type = 4

    def set_bin_string(self, lb):
        known_codes = lb.known_codes
        code_row = [x for x in known_codes if x[0] == self.freq_code]
        if len(code_row) == 0:
            if not self.annual_trips and "9" in self.freq_code:
                self.bin_string = "00000001"
            elif "9" in self.freq_code and self.annual_trips < 50:
                self.bin_string = "00000001"
            else:
                print("uhoh " + str(self.freq_code))
                self.bin_string = AddFrequencyCode(self.freq_code, self.annual_trips,
                                                   self.freq_code_description, lb).exec_()
        elif len(code_row) == 1:
            if self.annual_trips in (None, 0, "0"):
                self.annual_trips = code_row[0][1]
            self.freq_code_description = code_row[0][2]
            self.bin_string = code_row[0][3]
            self.set_days_of_week()
        else:
            print("more than one code found!!")
            print(self.freq_code)

    def is_holiday_schedule(self):

        if self.source == "HCR":
            # if "9" in self.freq_code:
            #     self.holiday = True
            if not self.annual_trips:
                self.annual_trips = 0
            if float(self.annual_trips) < 40:
                self.holiday = True
            if len(self.bin_string) == 8:
                if self.bin_string[7] in (1, "1"):
                    self.holiday = True

        if self.source == "PVS":
            if float(self.holiday_hours) > 0:
                self.holiday = True
            # if "9" in self.freq_code:
            #     self.holiday = True
            if self.bin_string:
                if len(self.bin_string) == 8:
                    if self.bin_string[7] in (1, "1"):
                        self.holiday = True

    def correct_travel_times(self):

        for x, stop in enumerate(self.stops[:-1]):
            next_stop = self.stops[x + 1]
            travel_time = dur(stop.depart_time, next_stop.arrive_time)
            if travel_time < 1:
                # print("shifting stops")
                for stop in self.stops[:x+1]:
                    stop.shift_stop(-1)

    # This method deals with read-in JDA schedules
    def multiple_load_fix(self, address_book):

        last_stop_indices = []

        for x, stop in enumerate(self.stops[:-1]):
            next_stop = self.stops[x+1]
            if stop.load_num != next_stop.load_num:
                last_stop_indices.append(x)

        for index in last_stop_indices:
            stop = self.stops[index]
            next_stop = self.stops[index + 1]
            current_travel_time = dur(stop.depart_time, next_stop.arrive_time)
            actual_travel_time = 0
            if stop.stop_name != next_stop.stop_name:
                if address_book.check_existing_trips(stop.stop_name, next_stop.stop_name):
                    travel_time = address_book.get_a_trip(stop.stop_name, next_stop.stop_name)
                    actual_travel_time = travel_time[3]
                else:
                    actual_travel_time, _ = address_book.add_a_trip(stop.stop_name, next_stop.stop_name,
                                                                    stop.address, next_stop.address)

            actual_travel_time = max(1, actual_travel_time)

            if current_travel_time > actual_travel_time:
                new_minutes = math.floor(current_travel_time - actual_travel_time)
                self.stops[index].change_depart_time(new_minutes)

    # These methods are for post-optimization schedules
    def set_new_schedule_name(self):

        start_minute = self.stops[0].arrive_cminute

        # if not 1200 < start_minute < 2640 + 15:
        #     print("hmmmm start time, check " + str(self.schedule_num))

        if start_minute < 1680 + 15:
            self.tour = "1"
        elif start_minute < 2161 + 15:
            self.tour = "2"
        elif start_minute < 2641 + 15:
            self.tour = "3"
        else:
            # here the schedule should be pushed to tour one of the next day
            self.tour = "1"
            self.adjust = True

        if self.vehicle_type in ("Single", "SINGLE"):
            vehicle_num = "1"
        elif self.vehicle_type in ("11-TON", "11-Ton"):
            vehicle_num = "3"
        else:
            vehicle_num = "5"

        read_in_index = str(self.read_in_index).zfill(3)
        new_schedule_num = self.tour + vehicle_num + read_in_index
        self.schedule_name = new_schedule_num

    def clean_optimized_schedule(self, ip):

        self.pre_cleaned_duration = int(self.raw_duration())
        self.pre_cleaned_stops = copy.deepcopy(self.stops)

        pvs_time = 14
        pvs_to_pdc_time = 1
        pdc_time = 10
        hours_wo_lunch = 6
        round_minute = 15

        if ip:
            pvs_time = ip.pvs_time
            pvs_to_pdc_time = ip.pvs_to_pdc
            pdc_time = ip.pdc_time
            hours_wo_lunch = ip.hours_wo_lunch
            round_minute = ip.round_time

        # attach first pvs
        first_arrive_time = self.stops[0].arrive_time
        pvs_depart_time = time_and_timedelta(first_arrive_time, -pvs_to_pdc_time)
        pvs_arrive_time = time_and_timedelta(pvs_depart_time, -pvs_time)
        # add in rounding to nearest 15 minutes
        # current_minute = pvs_arrive_time.minute % round_minute
        # if current_minute <= 10:
        #     self.fluff_minutes += current_minute
        #     pvs_arrive_time = time_and_timedelta(pvs_arrive_time, -current_minute)
        # else:
        #     self.fluff_minutes -= (15-current_minute)
        #     pvs_arrive_time = time_and_timedelta(pvs_arrive_time, 15 - current_minute)
        # done
        pvs_start = Stop(stop_name=self.pvs_name, arrive_time=pvs_arrive_time, depart_time=pvs_depart_time)
        self.stops.insert(0, pvs_start)

        # attach end pvs
        last_depart_time = self.stops[-1].depart_time
        pvs_arrive_time = time_and_timedelta(last_depart_time, pvs_to_pdc_time)
        pvs_depart_time = time_and_timedelta(pvs_arrive_time, pvs_time)
        pvs_stop = Stop(stop_name=self.pvs_name, arrive_time=pvs_arrive_time, depart_time=pvs_depart_time)
        self.stops.append(pvs_stop)

        # smoosh P&DCs
        for x, stop in enumerate(self.stops[:-1]):
            if stop.stop_name == self.pvs_pdc:
                if self.stops[x+1].stop_name == self.pvs_pdc:
                    stop.to_remove = True
                    self.stops[x+1].arrive_time = stop.arrive_time
                    self.stops[x+1].arrive_cminute = stop.arrive_cminute

        self.stops = [x for x in self.stops if not x.to_remove]

        for stop in self.stops:
            stop.is_post_office_location()

        # add lunch if necessary
        if self.raw_duration() >= hours_wo_lunch*60:
            lunch_stop = [x for x in self.stops if x.stop_name in ("LUNCH", "Lunch", "lunch")]
            if len(lunch_stop) > 1:
                print("MULTIPLE LUNCHES!!")
                return
            if len(lunch_stop) == 1:
                self.is_postalized = self.postal_compliance_check(ip)
                return
            else:
                self.add_a_lunch(ip)
                self.is_postalized = self.postal_compliance_check(ip)
        else:
            self.is_postalized = self.postal_compliance_check(ip)
            return

        self.is_postalized = self.postal_compliance_check(ip)

    def set_cminutes(self, ip):

        tour_one = cminute(ip.tour_one_time)

        original_first_minute = cminute(self.original_stops[0].arrive_time)
        postalized_first_minute = cminute(self.stops[0].arrive_time)

        if original_first_minute < postalized_first_minute:
            original_first_minute += 1440

        if original_first_minute < tour_one or postalized_first_minute < tour_one:
            original_first_minute += 1440
            postalized_first_minute += 1440

        source_check = self.source in ("HCR", "JDA")
        not_adjusted_check = self.freq_code == self.original_freq_code

        if original_first_minute < 1440 and postalized_first_minute > 1200 and source_check and not_adjusted_check:
            print(self.schedule_name, original_first_minute, postalized_first_minute)
            self.adjust_frequency_code()

        self.stops[0].arrive_cminute = postalized_first_minute
        self.stops[0].depart_cminute = postalized_first_minute + self.stops[0].layover()

        for x, stop in enumerate(self.stops[1:]):
            last_stop = self.stops[x]
            travel_time = dur(last_stop.depart_time, stop.arrive_time)
            stop.arrive_cminute = last_stop.depart_cminute + travel_time
            stop.depart_cminute = stop.arrive_cminute + stop.layover()

    def adjust_frequency_code(self):

        self.freq_code = create_adjust_code(self.freq_code)
        print(self.schedule_name, self.original_freq_code, self.freq_code)

        if len(self.bin_string) == 7:
            self.bin_string = self.bin_string[6] + self.bin_string[:6]
            self.set_days_of_week()
        elif len(self.bin_string) == 8:
            self.bin_string = self.bin_string[6] + self.bin_string[:6] + self.bin_string[7]
            self.set_days_of_week()

    def staffing_print(self, ws, row, dow):

        base_minute = 1440 * (dow-1)
        ws["A" + str(row)].value = dow
        ws["B" + str(row)].value = self.schedule_name

        start_minute = self.stops[0].arrive_cminute + base_minute
        ws["C" + str(row)].value = start_minute

        stop_minute = self.stops[-1].depart_cminute + base_minute
        ws["D" + str(row)].value = stop_minute

        ws["E" + str(row)].value = self.raw_duration()

        lunch_stops = [x for x in self.stops if x.stop_name in ("LUNCH", "Lunch", "lunch")]
        lunch_minutes = 0
        lunch_start = None
        lunch_end = None

        if len(lunch_stops) > 1:
            print("MULTIPLE LUNCHES FOUND!")
            print(self.schedule_name)
        elif lunch_stops:
            lunch_minutes += lunch_stops[0].layover()
            lunch_start = lunch_stops[0].arrive_cminute + base_minute
            lunch_end = lunch_stops[0].depart_cminute + base_minute

        ws["F" + str(row)].value = lunch_minutes
        ws["G" + str(row)].value = lunch_start
        ws["H" + str(row)].value = lunch_end

    def vehicle_print(self, ws, row, dow):

        ws["A" + str(row)].value = row - 8
        ws["B" + str(row)].value = self.schedule_name
        ws["C" + str(row)].value = row - 8
        ws["D" + str(row)].value = 0

        base_minute = dow * 1440

        start_minute = self.stops[0].arrive_cminute + base_minute
        stop_minute = self.stops[-1].depart_cminute + base_minute

        ws["E" + str(row)].value = start_minute
        ws["F" + str(row)].value = stop_minute
        ws["G" + str(row)].value = (stop_minute - start_minute)

    def excel_print_summary(self, ws, row):

        ws["A" + str(row)].value = self.short_name
        ws["B" + str(row)].value = self.schedule_name
        ws["C" + str(row)].value = self.freq_code
        ws["D" + str(row)].value = self.vehicle_type
        ws["E" + str(row)].value = len(self.stops)
        ws["F" + str(row)].value = self.stops[0].arrive_time
        ws["G" + str(row)].value = self.stops[-1].depart_time
        ws["H" + str(row)].value = int(self.raw_duration())
        ws["I" + str(row)].value = self.schedule_num

    def excel_print_stops(self, ws, row):

        for x, stop in enumerate(self.stops):
            ws["A" + str(row + x)].value = self.schedule_name
            ws["B" + str(row + x)].value = x + 1
            ws["C" + str(row + x)].value = stop.stop_name
            ws["D" + str(row + x)].value = stop.arrive_time
            ws["E" + str(row + x)].value = stop.depart_time
            ws["F" + str(row + x)].value = stop.original_schedule_name
            ws["G" + str(row + x)].value = stop.nass_code

    def is_network_schedule(self, network_sites):

        network_nass_codes = list(set([x[0] for x in network_sites]))
        stop_nass_codes = list(set(stop.nass_code for stop in self.stops))
        network_count = len([x for x in stop_nass_codes if x in network_nass_codes])

        if network_count > 1:
            self.network_schedule = True

    def crosses_state_lines(self):

        states = list(set([x.stop_name[-2:] for x in self.stops]))
        all_states = ['AL', 'AK', 'AZ', 'AR', 'CA', 'CO', 'CT', 'DE', 'DC', 'FL', 'GA', 'HI', 'ID', 'IL',
                      'IN', 'IA', 'KS', 'KY', 'LA', 'ME', 'MD', 'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE',
                      'NV', 'NH', 'NJ', 'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'PR', 'RI', 'SC',
                      'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'VI', 'WA', 'WV', 'WI', 'WY']
        states_in_all_states = [state for state in states if state in all_states]
        if len(states_in_all_states) > 1:
            self.cross_state_lines = True


class RoundTrip:

    def __init__(self, short_name, source_file, schedule_name, stops, vehicle_type, freq_code, pdc_name,
                 bin_string, holiday, trip_num, hours_wo_lunch, pvs_name):

        self.holiday = holiday
        self.source_file = source_file
        self.trip_num = trip_num
        self.optimizer_trip_num = None
        self.short_name = short_name
        self.schedule_name = schedule_name
        self.schedule_num = schedule_name.split()[-1]
        self.plate_name = self.schedule_name[:-len(self.schedule_num)]
        self.hours_wo_lunch = hours_wo_lunch

        self.stops = stops
        self.vehicle_type = vehicle_type
        self.freq_code = freq_code
        self.bin_string = bin_string

        self.duration = self.raw_duration()

        self.pvs_pdc = pdc_name
        self.pvs_name = pvs_name
        self.shrink_pdcs()
        self.contains_lunch = self.contains_a_lunch()
        self.trip_type = self.find_trip_type()
        # 1 = real trip; 2 = lunch trip; 3 = spotter trip; 4 = standby trip
        if self.trip_type == 1:
            self.trim_stops()
        self.contains_lunch = self.contains_a_lunch()
        # self.remove_lunch()
        # self.remove_spotter()

        self.start_minute = self.stops[0].arrive_cminute
        if not self.start_minute:
            print("here")
            print(self.trip_num)
            print(self.schedule_name)
            self.start_minute = cminute(self.stops[0].arrive_time)

        self.duration = self.raw_duration()
        self.end_minute = self.start_minute + self.duration
        self.end_minute = self.start_minute + self.duration
        self.is_selected = True

    def raw_duration(self):

        duration = 0
        for x, stop in enumerate(self.stops[:-1]):
            next_stop = self.stops[x + 1]
            duration += stop.layover()
            duration += dur(stop.depart_time, next_stop.arrive_time)

        duration += self.stops[-1].layover()
        return duration

    def shrink_pdcs(self):

        if self.stops[0].stop_name != self.pvs_pdc or self.stops[-1].stop_name != self.pvs_pdc:
            print("ERROR WITH ROUND TRIPS")

        if self.stops[0].layover() != 10:
            self.stops[0].change_arrive_time(self.stops[0].layover() - 10)

        if self.stops[-1].layover() != 10:
            self.stops[-1].change_depart_time(10 - self.stops[-1].layover())

    def contains_a_lunch(self):

        lunch_stops = [stop for stop in self.stops if stop.stop_name in ("LUNCH", "Lunch", "lunch")]

        if lunch_stops:
            return True
        else:
            return False

    def find_trip_type(self):

        standby_list = ["Standby", "STANDBY TIME", "STANDBY", "ASSIGNED TO OTHER DUTIES", self.pvs_name]
        spotter_list = ["Spotter", "SPOTTER", "spotter"]
        lunch_list = ["LUNCH", "Lunch", "lunch"]
        stops_categorized = []
        for stop in self.stops[1:-1]:
            if any(word in stop.stop_name for word in standby_list):
                stop.category = 4
                stops_categorized.append(4)
            elif any(word in stop.stop_name for word in lunch_list):
                stop.category = 2
                stops_categorized.append(2)
            elif any(word in stop.stop_name for word in spotter_list):
                stop.category = 3
                stops_categorized.append(3)
            else:
                stop.category = 1
                stops_categorized.append(1)

        stops_categorized = list(set(stops_categorized))
        if 1 in stops_categorized:
            return 1
        elif 3 in stops_categorized:
            return 3
        elif 2 in stops_categorized:
            return 2
        else:
            return 4

    def remove_lunch(self):

        if not self.contains_lunch:
            return
        if self.trip_type == 2:
            return

        lunch_stop = next(stop for stop in self.stops if stop.stop_name in ("Lunch", "LUNCH", "lunch"))
        lunch_stop_index = self.stops.index(lunch_stop)
        lunch_duration = lunch_stop.layover()
        max_allowed_duration = (self.hours_wo_lunch * 60) + lunch_duration

        if self.duration > max_allowed_duration:
            return

        if lunch_stop_index == 1:
            # lunch is at start P&DC, push start P&DC later
            del self.stops[1]
            self.stops[0].shift_stop(lunch_duration)
            self.contains_lunch = False
        elif lunch_stop_index == len(self.stops) - 2:
            del self.stops[-2]
            self.stops[-1].shift_stop(-lunch_duration)
            self.contains_lunch = False

    def spotter_standby_stop(self):

        check_list = ["Spotter", "SPOTTER", "spotter", "Standby", "STANDBY TIME", "STANDBY",
                      "ASSIGNED TO OTHER DUTIES"]

        stops = [stop for stop in self.stops if any(word in stop.stop_name for word in check_list)]

        if stops:
            return True
        else:
            return False

    def remove_spotter(self):

        if self.trip_type != 1:
            return

        check_list = ["Spotter", "SPOTTER", "spotter", "Standby", "STANDBY TIME", "STANDBY",
                      "ASSIGNED TO OTHER DUTIES"]

        while self.spotter_standby_stop():
            spotter_standby_stop = next(
                stop for stop in self.stops if any(word in stop.stop_name for word in check_list))
            spotter_standby_index = self.stops.index(spotter_standby_stop)
            spotter_standby_time = spotter_standby_stop.layover()

            if spotter_standby_index == 1:
                del self.stops[1]
                self.stops[0].shift_stop(spotter_standby_time)
            elif spotter_standby_index == len(self.stops) - 2:
                del self.stops[-2]
                self.stops[-1].shift_stop(-spotter_standby_time)
            else:
                self.remove_lunch()
                self.remove_spotter()

    def optidata_row(self, day_num):

        start_minute = self.start_minute + (day_num * 1440)
        end_minute = self.end_minute + (day_num * 1440)

        x = 0
        if self.contains_lunch:
            x = 1

        return [self.schedule_name, self.optimizer_trip_num, x, start_minute, end_minute, self.duration]

    def print_optidata_data(self, ws, row):

        name = self.schedule_name
        num = self.optimizer_trip_num
        veh = self.vehicle_type
        lunch = 0
        if self.contains_lunch:
            lunch = 1

        minutes_passed = self.start_minute

        for x, stop in enumerate(self.stops[:-1]):
            next_stop = self.stops[x + 1]
            stop.arrive_cminute = minutes_passed
            minutes_passed += stop.layover()
            stop.depart_cminute = minutes_passed
            minutes_passed += dur(stop.depart_time, next_stop.arrive_time)

        self.stops[-1].arrive_cminute = minutes_passed
        minutes_passed += self.stops[-1].layover()
        self.stops[-1].depart_cminute = minutes_passed

        test_duration = minutes_passed - self.start_minute
        if test_duration != self.duration:
            print("Duration error")

        for x, stop in enumerate(self.stops):
            ws["A" + str(x + row)].value = name
            ws["B" + str(x + row)].value = num
            ws["C" + str(x + row)].value = (x + 1)
            ws["D" + str(x + row)].value = stop.stop_name
            ws["E" + str(x + row)].value = stop.arrive_time
            ws["F" + str(x + row)].value = stop.depart_time
            ws["G" + str(x + row)].value = stop.arrive_cminute
            ws["H" + str(x + row)].value = stop.depart_cminute
            ws["I" + str(x + row)].value = lunch
            ws["J" + str(x + row)].value = veh
            ws["K" + str(x + row)].value = self.freq_code

    def trim_stops(self):

        real_stops = [x for x in self.stops if x.category == 1]
        first_index = self.stops.index(real_stops[0])
        last_index = self.stops.index(real_stops[-1])

        start_time = self.stops[first_index - 1].depart_time
        move_time = dur(self.stops[0].depart_time, start_time)

        while self.stops[1].category != 1:
            del self.stops[1]
            last_index -= 1

        self.stops[0].shift_stop(move_time)

        end_time = self.stops[last_index].depart_time
        move_time = dur(end_time, self.stops[-2].depart_time)

        while self.stops[-2].category != 1:
            del self.stops[-2]

        self.stops[-1].shift_stop(-move_time)