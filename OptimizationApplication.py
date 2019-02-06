# This is the inward facing app, that does our process and runs the optimizer
import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QHBoxLayout, QVBoxLayout, QLabel, QPushButton, QInputDialog,
                             QFileDialog, QMainWindow, QDialog, QTabWidget, QRadioButton, QProgressBar, QComboBox,
                             QLineEdit, QGridLayout)
from LookupReader import LookupBook
from GeneralMethods import InputPasser, qtime_to_time
# from PlateReader import HCRReader
from HCRPlateReaderV2 import HCRReader, CustomHCRReaderPopUp
from HTMLReader import HTMLReader
from OptimizerTable import (OptimizerView, PostalizersToPandas, PopUpView, PopUpToPandas, SchedulePopUpView,
                            SchedulePopUpToPandas, RoundTripView, RoundTripsToPandas, SingleRoundTripToPandas)
from ScheduleCompilation import Postalizer
from PostalizerInputTable import PostalizerInputsView, PostalizerInputsToLists
from OptimizedScheduleProcessor import (Process, OptimizedScheduleView, OptimizedSchedulesToPandas,
                                        OneOptScheduleView, OneOptScheduleToPandas)
from PostStaffingOptimization import PostStaffingProcess, StaffingOutputView, ReadInHuasOutput
from PostVehicleOptimization import VehicleSummarizer
from ExcelReader import JDAExcelReader
from StaffingManipulation import StaffingManipulationGUI


class App(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Optimizing Mail Model")
        self.setGeometry(100, 100, 1000, 500)

        self.lb = LookupBook()

        self.tabs = QTabWidget()

        self.tab1 = PostalizerPrinterGUI(self.lb)
        self.tabs.addTab(self.tab1, "Home Page")

        self.tab2 = PostCPLEXGUI(self.lb, self.tab1.ip)
        self.tabs.addTab(self.tab2, "Post Scheduler")

        self.tab3 = PostStaffingGUI(self.tab1.ip)
        self.tabs.addTab(self.tab3, "Optimized Vehicles")

        self.tab4 = StaffingManipulationGUI(self.tab1.ip)
        self.tabs.addTab(self.tab4, "Staffing Manipulation")

        self.input_tab = PostalizerInputsGUI()
        self.tabs.addTab(self.input_tab, "Change Inputs")
        self.input_tab.submit_button.clicked[bool].connect(self.set_ip)

        self.setCentralWidget(self.tabs)
        self.show()

    def set_ip(self):

        self.tab1.ip = self.input_tab.ip
        self.tab2.ip = self.input_tab.ip
        self.tab3.ip = self.input_tab.ip
        self.tab4.ip = self.input_tab.ip
        self.tab1.deep_clear()


class PostalizerPrinterGUI(QWidget):

    def __init__(self, lb):
        super().__init__()

        self.ip = InputPasser(source="Optimizer")
        self.hcr_check_list = [self.ip.min_check_time, self.ip.max_check_time, self.ip.max_combined_time]
        self.strict_merge = False

        self.lb = lb

        self.load_plates_button = None
        self.custom_load_plates_button = None
        self.load_html_button = None
        self.load_excel_button = None
        self.print_cplex_input_button = None
        self.print_schedules_button = None
        self.clear_button = None
        self.main_table = None

        self.show_schedules = None
        self.show_trips = None
        self.case = 1

        self.file_names = []
        self.readers = []

        self.site_names = []
        self.postalizers = []
        self.initUI()

    def initUI(self):

        main_layout = QHBoxLayout()
        left_layout = QVBoxLayout()

        self.main_table = OptimizerView()
        self.main_table.doubleClicked.connect(self.pop_up_site)

        self.load_plates_button = QPushButton("Load HCR Plates", self)
        self.custom_load_plates_button = QPushButton("Custom load HCR Plates", self)
        self.load_html_button = QPushButton("Load HTML File", self)
        self.load_excel_button = QPushButton("Load JDA (Excel) File", self)
        self.print_cplex_input_button = QPushButton("Print CPLEX input for selected row", self)
        self.print_schedules_button = QPushButton("Print schedules from selected row", self)
        self.clear_button = QPushButton("Clear All", self)
        self.show_schedules = QRadioButton("Schedules")
        self.show_trips = QRadioButton("Round Trips")

        self.load_plates_button.clicked[bool].connect(self.load_plates)
        self.custom_load_plates_button.clicked[bool].connect(self.custom_load_plates)
        self.load_html_button.clicked[bool].connect(self.load_html_files)
        self.load_excel_button.clicked[bool].connect(self.load_excel)
        self.print_cplex_input_button.clicked[bool].connect(self.print_selected_round_trips)
        self.print_schedules_button.clicked[bool].connect(self.print_selected_schedules)
        self.clear_button.clicked[bool].connect(self.deep_clear)

        self.show_schedules.case = 1
        self.show_trips.case = 2

        self.show_schedules.toggled.connect(self.change_case)
        self.show_trips.toggled.connect(self.change_case)

        self.show_schedules.setChecked(True)

        left_layout.addWidget(self.load_plates_button)
        left_layout.addWidget(self.custom_load_plates_button)
        left_layout.addWidget(self.load_html_button)
        left_layout.addWidget(self.load_excel_button)
        left_layout.addWidget(self.print_cplex_input_button)
        left_layout.addWidget(self.print_schedules_button)
        left_layout.addWidget(self.clear_button)
        left_layout.addStretch(1)

        pop_up_label = QLabel("Double click to show: ")
        left_layout.addWidget(pop_up_label)
        left_layout.addWidget(self.show_schedules)
        left_layout.addWidget(self.show_trips)
        left_layout.addStretch(5)

        main_layout.addLayout(left_layout)
        main_layout.addWidget(self.main_table)

        self.setLayout(main_layout)

    def change_case(self):

        radiobutton = self.sender()

        if radiobutton.isChecked():
            self.case = radiobutton.case

    def load_plates(self):

        self.hcr_check_list = [self.ip.min_check_time, self.ip.max_check_time, self.ip.max_combined_time]

        files, _ = QFileDialog.getOpenFileNames(self, "Select HCR Plate(s)",
                                                "", "PDF Files (*.pdf);;PDF Files (*.pdf)")

        if files:

            progress_bar = ProgressBarWindow("Progress Reading Plates", len(files))
            QApplication.processEvents()

            for x, pdf in enumerate(files):

                file_name = pdf.split('/')[-1][:5]
                if file_name in self.file_names:
                    print("duplicate not added")
                    continue

                self.file_names.append(file_name)

                progress_bar.update(x)

                reader = HCRReader(pdf, self.lb, "Optimizer", self.hcr_check_list, self.strict_merge)
                if reader.is_readable:
                    self.readers.append(reader)
                else:
                    self.readers.append(None)
                    print("error reading " + pdf[:5])
        else:
            print("No files selected")
            return

        self.update_table()
        self.postalize_schedules()

    def custom_load_plates(self):

        self.hcr_check_list = [self.ip.min_check_time, self.ip.max_check_time, self.ip.max_combined_time]

        files, _ = QFileDialog.getOpenFileNames(self, "Select HCR Plate(s)",
                                                "", "PDF Files (*.pdf);;PDF Files (*.pdf)")

        if not files:
            print("Nothing selected")
            return

        for x, pdf in enumerate(files):

            file_name = pdf.split('/')[-1][:5]
            if file_name in self.file_names:
                print("duplicate not added")
                continue

            self.file_names.append(file_name)

            test_reader = HCRReader(pdf, self.lb, "Test", self.hcr_check_list, self.strict_merge)
            possible_names, possible_addresses = test_reader.get_potential_pdcs_and_addresses()

            # given_name, given_address = CustomHCRReaderPopUp(possible_names, possible_addresses, self.lb).exec_()
            given_name, given_address = CustomHCRReaderPopUp(possible_names, possible_addresses, self.lb).exec_()

            reader = HCRReader(pdf, self.lb, "Optimizer", self.hcr_check_list, self.strict_merge, custom_pdc_name=given_name,
                               custom_pdc_address=given_address)

            if reader.is_readable:
                self.readers.append(reader)
            else:
                self.readers.append(None)
                print("error reading " + pdf[:5])

        self.update_table()
        self.postalize_schedules()

    def load_html_files(self):

        files, _ = QFileDialog.getOpenFileNames(self, "Select HTML File(s)",
                                                "", "HTML Files (*.htm);;HTML Files (*.htm)")

        if files:
            progress_bar = ProgressBarWindow("Progress Reading HTMLs", len(files))
            QApplication.processEvents()

            for x, html in enumerate(files):

                file_name = str(html)
                if file_name in self.file_names:
                    print("duplicate not added")
                    continue

                self.file_names.append(file_name)

                progress_bar.update(x)

                reader = HTMLReader(html, self.lb)
                self.file_names.append(str(html))
                self.readers.append(reader)
        else:
            print("No files selected")
            return

        self.update_table()
        self.postalize_schedules()

    def load_excel(self):

        file, _ = QFileDialog.getOpenFileName(self, "Select JDA to SPS file", "",
                                              "Excel Files (*.xlsx);;Excel Files (*.xlsx)")

        if file:

            file_name = file.split("/")[-1]
            if file_name not in self.file_names:
                self.file_names.append(file_name)
                reader = JDAExcelReader(file, self.lb)
                self.readers.append(reader)
            else:
                print("Duplicate not added")
        else:
            print("No file selected")
            return

        self.update_table()
        self.postalize_schedules()

    def update_table(self):

        self.clear_all()
        self.site_names = []

        for reader in self.readers:
            if reader:
                site_name = reader.pvs_name
            else:
                site_name = "Unread plates"

            if site_name not in self.site_names:
                self.site_names.append(site_name)
                self.postalizers.append(Postalizer(self.ip, site_name))

            postalizer = next(x for x in self.postalizers if x.pvs_name == site_name)
            postalizer.add_reader(reader)

        compiled_df = PostalizersToPandas(self.postalizers).df
        self.main_table.set_model(compiled_df)
        self.main_table.update()

    def postalize_schedules(self):

        for postalizer in self.postalizers:
            postalizer.postalize_schedules()

        self.update_table()

    def clear_all(self):

        del self.file_names
        del self.postalizers
        del self.site_names

        self.site_names = []
        self.file_names = []
        self.postalizers = []
        compiled_df = PostalizersToPandas(self.postalizers).df
        self.main_table.set_model(compiled_df)
        self.main_table.update()

    def deep_clear(self):

        del self.readers
        self.readers = []

        self.clear_all()

    def pop_up_site(self):

        index = self.main_table.selectedIndexes()[0]
        index = self.main_table.model().mapToSource(index)

        site_name = self.site_names[index.row()]
        postalizer = self.postalizers[index.row()]

        if postalizer:
            round_trips = postalizer.round_trips

        if self.case == 1:
            PlatePopUpGUI(site_name, postalizer).exec_()
        elif self.case == 2:
            if round_trips:
                to_print = RoundTripsGUI(site_name, round_trips).exec_()
                if to_print:
                    self.print_selected_round_trips()
        else:
            print("WTF!")

    def print_selected_round_trips(self):

        if len(self.postalizers) == 0:
            return

        index = self.main_table.selectedIndexes()
        if not index:
            print("nothing selected")
            return

        row = self.main_table.model().mapToSource(index[0]).row()

        postalizer_to_print = self.postalizers[row]
        # print(postalizer_to_print.short_name)

        if not postalizer_to_print:
            print("nothing selected!")
            return

        postalizer_to_print.print_cplex_scheduler_input()

    def print_selected_schedules(self):

        if len(self.postalizers) == 0:
            return

        index = self.main_table.selectedIndexes()
        if not index:
            print("nothing selected")
            return

        row = self.main_table.model().mapToSource(index[0]).row()

        postalizer_to_print = self.postalizers[row]
        # print(postalizer_to_print.short_name)

        if not postalizer_to_print:
            print("nothing selected!")
            return

        postalizer_to_print.print_schedules()


class PlatePopUpGUI(QDialog):

    def __init__(self, site_name, postalizer, parent=None):
        super(PlatePopUpGUI, self).__init__(parent)

        self.site_name = site_name
        self.setWindowTitle(self.site_name)

        if site_name == "Unread plates" or not postalizer:
            self.no_reader()
            return

        self.schedules = postalizer.schedules
        self.close_button = None
        self.central_table = None
        self.initUI()

    def initUI(self):

        vbox = QVBoxLayout()
        button_box = QHBoxLayout()

        self.setGeometry(200, 200, 900, 500)
        self.central_table = PopUpView()
        self.central_table.doubleClicked.connect(self.pop_up_schedule)
        self.central_table.set_model(PopUpToPandas(self.schedules).df)

        self.close_button = QPushButton("Close", self)
        self.close_button.clicked[bool].connect(self.accept)

        vbox.addWidget(self.central_table)
        vbox.addWidget(self.close_button)

        self.setLayout(vbox)
        self.show()

    def no_reader(self):

        vbox = QVBoxLayout()
        vbox.addWidget(QLabel("Postalizer not read in correctly"))
        close_button = QPushButton("Close", self)
        close_button.clicked[bool].connect(self.accept)
        vbox.addWidget(close_button)
        self.setLayout(vbox)

    def pop_up_schedule(self):

        index = self.central_table.selectedIndexes()[0]
        index = self.central_table.model().mapToSource(index)
        schedule = self.schedules[index.row()]
        SchedulePopUpGUI(schedule).exec_()


class RoundTripsGUI(QDialog):

    def __init__(self, site_name, round_trips, parent=None):
        super(RoundTripsGUI, self).__init__(parent)

        self.trips = round_trips
        self.site_name = site_name
        self.setWindowTitle(self.site_name)
        self.setGeometry(150, 100, 950, 550)
        self.close_button = None
        self.round_trip_table = None
        self.vehicle_filter_combo = None
        self.select_all_button = None
        self.all_selected = True
        self.to_print = False
        self.trip_filter_combo = None
        self.print_selection_button = None

        self.all_trips = None
        self.real_trips = None
        self.spotter_trips = None

        self.vehicle_filter = ""
        self.trip_filter = ""

        self.initUI()

    def initUI(self):

        vbox = QVBoxLayout()
        top_hbox = QHBoxLayout()
        bottom_bar = QHBoxLayout()

        self.vehicle_filter_combo = QComboBox()
        self.vehicle_filter_combo.addItem("All Vehicles")
        self.vehicle_filter_combo.addItem("11-Ton")
        self.vehicle_filter_combo.addItem("Single")
        self.vehicle_filter_combo.setCurrentIndex(0)
        self.vehicle_filter_combo.currentIndexChanged.connect(self.vehicle_filter_change)

        self.trip_filter_combo = QComboBox()
        self.trip_filter_combo.addItem("All Trip Types")
        self.trip_filter_combo.addItem("Real trip")
        self.trip_filter_combo.addItem("Spotter trip")
        self.trip_filter_combo.addItem("Standby trip")
        self.trip_filter_combo.addItem("Lunch trip")
        self.trip_filter_combo.currentIndexChanged.connect(self.trip_filter_change)

        self.select_all_button = QPushButton("Select/Unselect All", self)
        self.select_all_button.clicked[bool].connect(self.select_all_click)

        self.all_trips = QRadioButton("All Trips")
        self.real_trips = QRadioButton("Real Trips")
        self.spotter_trips = QRadioButton("Spotter Trips")
        self.all_trips.setChecked(True)

        self.round_trip_table = RoundTripView()
        self.set_table()

        self.print_selection_button = QPushButton("Print CPlex Input for Selection", self)
        self.print_selection_button.clicked[bool].connect(self.print_selection)

        self.close_button = QPushButton("Close", self)
        self.close_button.clicked[bool].connect(self.accept)

        filter_label = QLabel("Filter Trips:")
        top_hbox.addWidget(filter_label)
        top_hbox.addStretch(2)
        top_hbox.addWidget(self.vehicle_filter_combo)
        top_hbox.addWidget(self.trip_filter_combo)
        top_hbox.addStretch(1)
        top_hbox.addWidget(self.select_all_button)

        vbox.addLayout(top_hbox)
        vbox.addWidget(self.round_trip_table)
        bottom_bar.addWidget(self.print_selection_button)
        bottom_bar.addWidget(self.close_button)
        vbox.addLayout(bottom_bar)

        self.setLayout(vbox)
        self.show()

    def vehicle_filter_change(self):

        self.vehicle_filter = self.vehicle_filter_combo.currentText()
        if self.vehicle_filter == "All Vehicles":
            self.vehicle_filter = ""

        self.filter_table()

    def trip_filter_change(self):

        self.trip_filter = self.trip_filter_combo.currentText()
        if self.trip_filter == "All Trip Types":
            self.trip_filter = ""

        self.filter_table()

    def filter_table(self):

        if self.round_trip_table.sorter:
            self.round_trip_table.sorter.setFilterByColumn(self.vehicle_filter, 8)
            self.round_trip_table.sorter.setFilterByColumn(self.trip_filter, 9)
        else:
            return

    def select_all_click(self):

        if self.all_selected:
            self.unselect_all_round_trips()
            self.all_selected = False
        else:
            self.select_all_round_trips()
            self.all_selected = True

    def select_round_trips(self, item):

        index = self.round_trip_table.model().mapToSource(item)
        if index.column() != 19:
            return

        row = index.row()
        col = index.column()

        new_val = self.round_trip_table.model().sourceModel()._data[row][col]

        if new_val:
            self.trips[row].is_selected = False
        else:
            self.trips[row].is_selected = True

        self.set_table()

    def select_all_round_trips(self):

        for trip in self.trips:
            trip.is_selected = True

        self.set_table()

    def unselect_all_round_trips(self):

        for trip in self.trips:
            trip.is_selected = False

        self.set_table()

    def set_table(self):

        self.round_trip_table.set_model(RoundTripsToPandas(self.trips).df)
        self.round_trip_table.setSortingEnabled(True)
        self.round_trip_table.doubleClicked.connect(self.pop_up_round_trip)
        self.round_trip_table.model().dataChanged.connect(self.select_round_trips)

    def print_selection(self):

        self.to_print = True
        self.accept()

    def pop_up_round_trip(self):

        try:
            index = self.round_trip_table.selectedIndexes()[0]
            index = self.round_trip_table.model().mapToSource(index)
            round_trip = self.trips[index.row()]
            RoundTripPopUpGUI(round_trip).exec_()
        except:
            return

    def exec_(self):

        super().exec_()
        self.close()

        return self.to_print


class SchedulePopUpGUI(QDialog):

    def __init__(self, schedule, parent=None):
        super(SchedulePopUpGUI, self).__init__(parent)

        self.schedule = schedule
        self.setWindowTitle(schedule.schedule_name)
        self.setGeometry(250, 100, 800, 550)
        self.close_button = None
        self.left_table = None
        self.right_table = None
        self.round_trip_table = None
        self.initUI()

    def initUI(self):

        vbox = QVBoxLayout()
        left_vbox = QVBoxLayout()
        right_vbox = QVBoxLayout()
        round_trip_vbox = QVBoxLayout()
        hbox = QHBoxLayout()

        left_label = QLabel("Schedule as read:")
        self.left_table = SchedulePopUpView()
        self.left_table.set_model(SchedulePopUpToPandas(self.schedule, 1).df)

        if self.schedule.is_postalized:
            right_label = "Postalized Schedule: Compliant"
        else:
            right_label = "Postalized Schedule: Non-Compliant"

        right_label = QLabel(right_label)
        self.right_table = SchedulePopUpView()
        self.right_table.set_model(SchedulePopUpToPandas(self.schedule, 2).df)

        round_trip_label = QLabel("Round Trips:")
        self.round_trip_table = RoundTripView()
        self.round_trip_table.set_model(RoundTripsToPandas(self.schedule.round_trips).df)
        self.round_trip_table.setSortingEnabled(False)
        self.round_trip_table.doubleClicked.connect(self.pop_up_round_trip)

        self.close_button = QPushButton("Close", self)
        self.close_button.clicked[bool].connect(self.accept)

        left_vbox.addWidget(left_label)
        left_vbox.addWidget(self.left_table)
        right_vbox.addWidget(right_label)
        right_vbox.addWidget(self.right_table)
        round_trip_vbox.addWidget(round_trip_label)
        round_trip_vbox.addWidget(self.round_trip_table)

        hbox.addLayout(left_vbox)
        hbox.addLayout(right_vbox)
        # hbox.addLayout(round_trip_vbox)
        vbox.addLayout(hbox)
        vbox.addLayout(round_trip_vbox)
        vbox.addWidget(self.close_button)

        self.setLayout(vbox)
        self.show()

    def pop_up_round_trip(self):

        index = self.round_trip_table.selectedIndexes()[0]
        index = self.round_trip_table.model().mapToSource(index)
        round_trip = self.schedule.round_trips[index.row()]
        RoundTripPopUpGUI(round_trip).exec_()


class RoundTripPopUpGUI(QDialog):

    def __init__(self, round_trip, parent=None):
        super(RoundTripPopUpGUI, self).__init__(parent)

        self.setWindowTitle(round_trip.schedule_name + " " + str(round_trip.trip_num))
        self.setGeometry(300, 300, 500, 300)
        self.round_trip = round_trip
        self.central_table = SchedulePopUpView()
        self.central_table.set_model(SingleRoundTripToPandas(round_trip).df)

        self.close_button = QPushButton("Close", self)
        self.close_button.clicked[bool].connect(self.accept)

        vbox = QVBoxLayout()
        label = QLabel("Round Trip: ")
        vbox.addWidget(label)
        vbox.addWidget(self.central_table)
        vbox.addWidget(self.close_button)
        self.setLayout(vbox)
        self.show()


class PostalizerInputsGUI(QWidget):

    def __init__(self):
        super().__init__()

        self.inputs_table = None
        self.ip = None

        # trip inputs
        self.pvs_to_pdc = None
        self.pvs_time = None
        self.pdc_time = None

        # lunch inputs
        self.lunch_duration = None
        self.hours_wo_lunch = None
        self.allow_non_postal = None

        # other requirements
        self.max_duration = None
        self.max_mileage = None
        self.tour_one = None
        self.nd_pm = None
        self.nd_am = None

        # u_part combination sensitivity
        self.min_check_minutes = None
        self.max_check_minutes = None
        self.max_combined_minutes = None
        self.use_check_codes = None

        self.submit_button = None
        self.restore_defaults_button = None
        self.initUI()

    def initUI(self):

        hbox = QHBoxLayout()
        button_hbox = QHBoxLayout()
        left_vbox = QVBoxLayout()

        self.restore_defaults_button = QPushButton("Restore Defaults", self)
        self.submit_button = QPushButton("Submit Changes", self)

        self.restore_defaults_button.clicked[bool].connect(self.restore_defaults)
        self.submit_button.clicked[bool].connect(self.save_changes)

        button_hbox.addWidget(self.restore_defaults_button)
        button_hbox.addWidget(self.submit_button)

        self.inputs_table = PostalizerInputsView()
        inputs, headers = PostalizerInputsToLists().get_lists()
        self.inputs_table.set_model(inputs, headers)
        left_vbox.addWidget(self.inputs_table)
        left_vbox.addLayout(button_hbox)

        hbox.addLayout(left_vbox, 1)
        self.setLayout(hbox)

    def restore_defaults(self):

        inputs, headers = PostalizerInputsToLists().get_lists()
        self.inputs_table.set_model(inputs, headers)
        self.inputs_table.update()

    def save_changes(self):

        data = self.inputs_table.model()._data
        pvs_time = data[1][1]
        pvs_to_pdc = data[2][1]
        pdc_time = data[3][1]
        lunch_minutes = data[5][1]
        hours_wo_lunch = data[6][1]
        allow_non_postal = data[7][1]
        wash_up_time = data[8][1]
        max_working_time = data[11][1]
        tour_one = qtime_to_time(data[12][1])
        nd_pm = qtime_to_time(data[13][1])
        nd_am = qtime_to_time(data[14][1])
        min_check_time = data[16][1]
        max_check_time = data[17][1]
        combined_max_time = data[18][1]
        new_ip = InputPasser(source="Optimizer", pvs_to_pdc=pvs_to_pdc, pvs_time=pvs_time, pdc_time=pdc_time,
                             lunch_duration=lunch_minutes, hours_wo_lunch=hours_wo_lunch,
                             allow_non_postal=allow_non_postal, max_duration=max_working_time,tour_one_time=tour_one,
                             nd_pm=nd_pm, nd_am=nd_am, min_check_time=min_check_time, max_check_time=max_check_time,
                             max_combined_time=combined_max_time, wash_up_time=wash_up_time)

        self.ip = new_ip


class ProgressBarWindow(QWidget):

    def __init__(self, title, length):
        super().__init__()
        self.setWindowTitle(title)
        vbox = QVBoxLayout()
        self.length = length
        self.pbar = QProgressBar(self)
        self.pbar.setMaximum(length)
        self.pbar.setGeometry(10, 10, 200, 50)
        vbox.addWidget(self.pbar)
        self.label = QLabel("Reading: x/" + str(length))
        vbox.addWidget(self.label)
        self.setLayout(vbox)
        self.setGeometry(300, 300, 300, 50)
        self.show()

    def update(self, value):

        self.pbar.setValue(value)
        self.label.setText("Reading Plate: " + str(value) + "/" + str(self.length))


class PostCPLEXGUI(QWidget):

    def __init__(self, lb, ip):
        super().__init__()

        self.process_button = None
        self.print_button = None
        self.processor = None
        self.schedules = None
        self.main_table = None
        self.lb = lb
        self.ip = ip

        self.initUI()

    def initUI(self):

        left_vbox = QVBoxLayout()
        hbox = QHBoxLayout()

        self.process_button = QPushButton("Process a Workbook", self)
        self.process_button.clicked[bool].connect(self.process_a_book)

        self.print_button = QPushButton("Print to Excel", self)
        self.print_button.clicked[bool].connect(self.print_processor)

        left_vbox.addWidget(self.process_button)
        left_vbox.addWidget(self.print_button)
        left_vbox.addStretch(1)

        self.main_table = OptimizedScheduleView()
        self.main_table.doubleClicked.connect(self.pop_up_a_schedule)

        hbox.addLayout(left_vbox)
        hbox.addWidget(self.main_table)

        self.setLayout(hbox)

    def process_a_book(self):

        file, _ = QFileDialog.getOpenFileName(self, "Select Post-Cplex File", "",
                                              "Excel Files (*.xlsx);;Excel Files (*.xlsx)")

        if file:
            self.processor = Process(file, self.lb, self.ip)
            self.schedules = self.processor.schedules
            self.update_table()

    def update_table(self):

        if not self.schedules:
            return

        else:
            self.main_table.set_model(OptimizedSchedulesToPandas(self.schedules).df)

    def pop_up_a_schedule(self):

        index = self.main_table.selectedIndexes()[0]
        index = self.main_table.model().mapToSource(index)

        schedule = self.schedules[index.row()]

        OptSchedulePopUp(schedule).exec_()

    def print_processor(self):

        if not self.processor:
            return

        self.processor.print_all()


class OptSchedulePopUp(QDialog):

    def __init__(self, schedule, parent=None):

        super(OptSchedulePopUp, self).__init__(parent)

        self.setWindowTitle(schedule.schedule_name)
        self.setGeometry(300, 200, 800, 400)

        vbox = QVBoxLayout()
        table_hbox = QHBoxLayout()

        self.left_table = OneOptScheduleView()
        self.left_table.set_model(OneOptScheduleToPandas(schedule, 2).df)

        self.right_table = OneOptScheduleView()
        self.right_table.set_model(OneOptScheduleToPandas(schedule, 1).df)

        self.close_button = QPushButton("Close", self)
        self.close_button.clicked[bool].connect(self.accept)

        table_hbox.addWidget(self.left_table)
        table_hbox.addWidget(self.right_table)

        vbox.addLayout(table_hbox)
        vbox.addWidget(self.close_button)

        self.setLayout(vbox)
        self.show()


class PostStaffingGUI(QWidget):

    def __init__(self, ip):
        super().__init__()

        self.ip = ip
        # self.process_button = None
        self.process_vehicles_button = None
        self.processor = None
        self.main_table = None

        self.initUI()

    def initUI(self):

        outer_layout = QHBoxLayout()
        left_vbox = QVBoxLayout()

        # self.process_button = QPushButton("Select Hua's output/Post-Scheduling File", self)
        # self.process_button.clicked[bool].connect(self.process_staffing)

        self.process_vehicles_button = QPushButton("Select 11-Ton and Single Solutions", self)
        self.process_vehicles_button.clicked[bool].connect(self.process_vehicles)

        self.main_table = StaffingOutputView()

        # left_vbox.addWidget(self.process_button)
        left_vbox.addWidget(self.process_vehicles_button)
        left_vbox.addStretch(1)
        outer_layout.addLayout(left_vbox)
        outer_layout.addWidget(self.main_table)

        self.setLayout(outer_layout)

    # def process_staffing(self):
    #
    #     hua_file, _ = QFileDialog.getOpenFileName(self, "Select Hua's File", "",
    #                                               "Excel Files (*.xls*);;Excel Files (*.xls*)")
    #
    #     file, _ = QFileDialog.getOpenFileName(self, "Select Staffing Input Data File", "",
    #                                           "Excel Files (*.xlsx);;Excel Files (*.xlsx)")
    #
    #     if not (hua_file and file):
    #         return
    #
    #     hua_processor = ReadInHuasOutput(hua_file, file)
    #     self.processor = PostStaffingProcess(file, self.ip)

    def process_vehicles(self):

        site_name, _ = QInputDialog.getText(self, "Input Site Name", "Site Name:", QLineEdit.Normal, "")

        if not site_name:
            site_name = "None Entered"

        eleven_ton_file, _ = QFileDialog.getOpenFileName(self, "Select Eleven-Ton File", "",
                                                         "Excel Files (*.xlsx);;Excel Files (*.xlsx)")

        single_file, _ = QFileDialog.getOpenFileName(self, "Select Single File", "",
                                                     "Excel Files (*.xlsx);;Excel Files (*.xlsx)")

        summaries = VehicleSummarizer(eleven_ton_file, single_file, site_name)


class CustomHCRReader(QDialog):

    def __init__(self, possible_names, possible_addresses, lb, parent=None):
        super(CustomHCRReader, self).__init__(parent)

        self.setWindowTitle("Select P&DC For Plate")

        self.read_names = possible_names
        self.read_addresses = possible_addresses

        self.lb = lb
        self.sorted = sorted(lb.hcr_pdcs)

        self.found_on_plate_selection = None
        self.found_in_lb_selection = None

        self.accept_button = None
        self.update_lookup_book_button = None

        self.line_edits = []
        for x in range(0, 8):
            self.line_edits.append(QLineEdit())

        self.labels = ["HCR P&DC Name", "PVS P&DC Name", "PVS Name", "Postal Facility Name", "Alt. PVS Name",
                       "Alt. P&DC Name", "Short Name", "Address"]

        self.initUI()

    def initUI(self):

        outer_layout = QVBoxLayout()
        top_grid = QGridLayout()
        bottom_grid = QGridLayout()

        self.found_on_plate_selection = QComboBox()
        self.found_on_plate_selection.addItem("None")
        for name in self.read_names:
            self.found_on_plate_selection.addItem(name)

        self.found_on_plate_selection.activated.connect(self.update_from_plate)

        self.found_in_lb_selection = QComboBox()
        self.found_in_lb_selection.addItem("None")
        for name in self.sorted:
            self.found_in_lb_selection.addItem(name)

        self.found_in_lb_selection.activated.connect(self.update_from_lb)

        top_grid.addWidget(QLabel("Found on Plate: "), 0, 0)
        top_grid.addWidget(QLabel("Found in Lookups: "), 1, 0)

        top_grid.addWidget(self.found_on_plate_selection, 0, 1)
        top_grid.addWidget(self.found_in_lb_selection, 1, 1)

        outer_layout.addWidget(QLabel("Select P&DC by Name: "))
        outer_layout.addLayout(top_grid)
        outer_layout.addStretch(1)

        for x in range(0, 8):
            bottom_grid.addWidget(QLabel(self.labels[x]), x, 0)
            bottom_grid.addWidget(self.line_edits[x], x, 1)

        outer_layout.addWidget(QLabel("Edit Information: "))
        outer_layout.addLayout(bottom_grid)
        outer_layout.addStretch(1)

        self.accept_button = QPushButton("Set Information", self)
        self.accept_button.clicked[bool].connect(self.accept)

        self.update_lookup_book_button = QPushButton("Update Lookups", self)
        self.update_lookup_book_button.clicked[bool].connect(self.update_lookups)

        outer_layout.addWidget(self.accept_button)
        outer_layout.addWidget(self.update_lookup_book_button)
        self.setLayout(outer_layout)

    def update_from_plate(self):

        name = self.found_on_plate_selection.currentText()

        if name in self.sorted:
            index = self.sorted.index(name)
        elif name in [x[:-4] for x in self.sorted]:
            index = [x[:-4] for x in self.sorted].index(name)
        else:
            index = -1

        self.found_in_lb_selection.setCurrentIndex(index + 1)

    def update_from_lb(self):

        self.found_on_plate_selection.setCurrentIndex(0)

    def update_lookups(self):

        pass

    def exec(self):

        super().exec_()
        self.close()

        return None


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
