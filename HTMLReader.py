
from bs4 import BeautifulSoup
from ScheduleClasses import Schedule, Stop
from datetime import datetime
from LookupReader import FacilityNameEntry


def to_normal_hours(hours):
    numbers = hours.split(":")
    return int(numbers[0]) + int(numbers[1]) / 60


def html_to_time(var):
    return datetime.strptime(var, '%H:%M').time()


class HTMLReader:

    def __init__(self, file_name, lb):

        self.lb = lb
        self.soup = BeautifulSoup(open(file_name), "html.parser")

        self.source = "PVS"
        self.source_type = "HTML"

        self.short_file_name = file_name.split("/")[-1][:-4]
        self.source_file = self.short_file_name

        self.pages = []
        self.break_out_pages()

        # Find facility names
        self.short_name = None
        self.facility_name = self.pages[0][1][1]
        # these will be used in postalized schedules
        self.pvs_name = None
        self.pvs_pdc = None
        # these are found on HTML
        self.alternate_pvs_name = None
        self.alternate_pdc_name = None

        self.get_facility_names()

        self.schedules = []
        self.schedule_nums = []

        self.break_down_pages()
        self.set_bin_strings()

    def break_out_pages(self):

        main_table = self.soup.findAll('table')[0]
        text_in_rows = []

        for row in main_table.findAll('tr'):
            row_text = []
            for div in row.findAll('div'):
                if div.getText():
                    row_text.append(div.getText())
            text_in_rows.append(row_text)

        text_in_rows = [x for x in text_in_rows if x != []]

        first_list = []
        for x, row in enumerate(text_in_rows):
            if row[0] == "U.S. Postal Service":
                first_list.append(x)
        first_list.append(len(text_in_rows) - 1)

        x = 0
        start_list = []
        end_list = []

        while x < len(first_list) - 2:
            x += 1
            start_list.append(first_list[x])
            x += 1
            end_list.append(first_list[x])

        pages = []
        for x, start in enumerate(start_list):
            pages.append(text_in_rows[start:end_list[x]])

        self.pages = pages

    def get_facility_names(self):

        if self.facility_name not in self.lb.postal_facility_names:
            FacilityNameEntry(source_type="pvs_postal_facility", lb=self.lb,
                              file_name=self.short_file_name, html_postal_facility=self.facility_name).exec_()

        index = self.lb.postal_facility_names.index(self.facility_name)
        self.pvs_pdc = self.lb.pvs_pdcs[index]
        self.pvs_name = self.lb.pvs_names[index]
        self.short_name = self.lb.short_names[index]
        self.alternate_pdc_name = self.lb.alternate_pdc_names[index]
        self.alternate_pvs_name = self.lb.alternate_pvs_names[index]
        return

    def break_down_pages(self):

        for page in self.pages:
            self.break_down_a_page(page)

    def break_down_a_page(self, page):

        # find vehicles
        vehicles = []
        num_vehicles = len(page[7])
        for x in range(0, num_vehicles):
            vehicles.append(page[7][x])

        # remove empty data
        page = [x for x in page if len(x) > 2]

        # Break down stops
        try:
            stop_line = page.index(next(line for line in page if line[0] == 'Stop'))
        except:
            print("Couldn't break down page")
            return

        stop_lines = page[stop_line + 1:]
        stops = []
        stop_nums = []

        for stop in stop_lines:
            stop_num = stop[0]

            if stop_num in stop_nums:
                continue

            stop_nums.append(stop_num)

            stop_name = stop[3].replace('\n', ' ')
            if self.alternate_pvs_name in stop_name:
                stop_name = self.pvs_name
            elif self.alternate_pdc_name in stop_name:
                stop_name = self.pvs_pdc

            arrive_time = html_to_time(stop[4])
            depart_time = html_to_time(stop[5])

            new_stop = Stop(stop_name=stop_name, arrive_time=arrive_time, depart_time=depart_time)
            stops.append(new_stop)

        # break down header
        header_lines = page[:stop_line]

        h1 = header_lines[1]
        pvs_site = h1[0]
        postal_facility = h1[1]
        schedule_type = h1[2]
        tour = h1[3]
        run_num = h1[4]
        mileage = h1[5]
        freq_code = h1[6]
        tractor = h1[7]
        schedule_num = h1[8]

        h2 = header_lines[3]
        effective_date = h2[0]
        end_date = h2[1]
        unassigned_num = h2[3]
        unassigned_time = to_normal_hours(h2[4])
        spotter_num = h2[5]
        spotter_time = to_normal_hours(h2[6])
        paid_time = to_normal_hours(h2[7])

        try:
            lunch_hours = h2[9]
        except:
            lunch_hours = None

        if schedule_num in self.schedule_nums:
            print("Duplicate schedule found: " + schedule_num)
            return

        if tractor == "X":
            vehicle_type = "Single"
        else:
            vehicle_type = "11-Ton"

        vehicle_two = None
        vehicle_three = None

        h3 = header_lines[4]
        vehicle_one = vehicles[0]
        if len(vehicles) > 1:
            vehicle_two = vehicles[1]
        if len(vehicles) > 2:
            vehicle_three = vehicles[2]

        x = 0
        if header_lines[6][0] == "TIMES Route":
            x = 1

        h4 = header_lines[5 + x]
        times_trip = h4[2]

        h5 = header_lines[6 + x]
        times_route = h5[0]
        total_miles = h5[1]
        total_hours = to_normal_hours(h5[2])
        weekday_hours = to_normal_hours(h5[3])
        saturday_hours = to_normal_hours(h5[4])
        sunday_hours = to_normal_hours(h5[5])
        holiday_hours = to_normal_hours(h5[6])
        night_diff_hours = to_normal_hours(h5[7])

        # other_info = [pvs_site, postal_facility, schedule_type, tour, schedule_num, unassigned_num,
        #               unassigned_time, spotter_num, spotter_time, paid_time, lunch_hours, vehicle_one,
        #               vehicle_two, vehicle_three, times_route, times_trip, total_miles, total_hours,
        #               weekday_hours, saturday_hours, sunday_hours, holiday_hours, night_diff_hours]

        new_schedule = Schedule(source_type=self.source_type, source=self.source, schedule_num=schedule_num,
                                run_num=run_num, tractor=tractor, freq_code=freq_code, mileage=mileage,
                                pvs_pdc=self.pvs_pdc, pvs_name=self.pvs_name, stops=stops,
                                effective_date=effective_date, end_date=end_date, tour=tour,
                                vehicle_type=vehicle_type, vehicle_category=vehicle_type, short_name=self.short_name,
                                source_file=self.short_file_name)

        new_schedule.add_html_info(total_miles, total_hours, weekday_hours, saturday_hours, sunday_hours,
                                   holiday_hours, night_diff_hours, paid_time, unassigned_time, self.lb)

        new_schedule.is_holiday_schedule()

        self.schedules.append(new_schedule)
        self.schedule_nums.append(schedule_num)

    def set_bin_strings(self):

        for schedule in self.schedules:
            schedule.set_bin_string(self.lb)
            schedule.is_holiday_schedule()
