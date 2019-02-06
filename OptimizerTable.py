
import pandas as pd
from PyQt5.QtCore import QAbstractTableModel, Qt, QVariant, QSortFilterProxyModel
from PyQt5.QtWidgets import QTableView
from PyQt5.QtGui import QFont
from GeneralMethods import number, minutes_to_hours_and_minutes, quick_color


class OptimizerModel(QAbstractTableModel):

    def __init__(self, data, parent=None):
        QAbstractTableModel.__init__(self, parent)
        self._data = data.values
        self.header_data = data.columns

    def rowCount(self, parent=None):
        if not self._data[0][0]:
            return 0
        return len(self._data)

    def columnCount(self, parent=None):
        return len(self._data[0])

    def data(self, index, role):
        if not index.isValid():
            return None
        elif role == Qt.BackgroundRole:
            row=index.row()
            # if self._data[row][4]:
            #     return QBrush(QColor(51, 255, 51))
            if not self._data[row][4]:
                return quick_color("red")
        elif role == Qt.DisplayRole:
            value = self._data[index.row()][index.column()]
            return str(value)
        else:
            return None

    def headerData(self, col, orientation, role):

        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.header_data[col]

        return QVariant()

    def flags(self, index):

        return Qt.ItemIsEnabled | Qt.ItemIsSelectable


class OptimizerView(QTableView):

    def __init__(self, parent=None):
        super().__init__()
        self.sorter = None

    def set_model(self, data):

        self.sorter = QSortFilterProxyModel()
        self.sorter.setDynamicSortFilter(True)
        self.sorter.setSortRole(Qt.EditRole)
        self.sorter.setSourceModel(OptimizerModel(data))
        self.setModel(self.sorter)
        self.resizeColumnsToContents()
        self.setSortingEnabled(True)
        # self.setSelectionBehavior(QAbstractItemView.SelectRows)


class PostalizersToPandas:

    def __init__(self, postalizers):

        labels = ["PVS Name", "Source(s)", "Plate Number(s)", "# Schedules", "Postalizable", "# Postalizable Schedules",
                  "# Non-Postalizable Schedules", "# Trips"]

        compiled = []

        for postalizer in postalizers:
            if postalizer:
                compiled.append(postalizer.output_list)
            else:
                compiled.append([None, None, 0, 0, "", 0, 0, 0])
        if len(postalizers) == 0:
            compiled.append([None, None, 0, 0, "", 0, 0, 0])

        self.df = pd.DataFrame(compiled, columns=labels)


class PopUpModel(QAbstractTableModel):

    def __init__(self, data, parent=None):
        QAbstractTableModel.__init__(self, parent)
        self._data = data.values
        self.header_data = data.columns

    def rowCount(self, parent=None):
        try:
            if not self._data[0][0]:
                return 0
            return len(self._data)
        except:
            return 0

    def columnCount(self, parent=None):
        try:
            return len(self._data[0])
        except:
            return 0

    def data(self, index, role):
        if not index.isValid():
            return None
        elif role == Qt.TextAlignmentRole:
            return Qt.AlignCenter
        elif role == Qt.BackgroundRole:
            row = index.row()
            if not self._data[row][5] and index.column() in (5, 6, 7, 8):
                # return QBrush(QColor(255, 145, 164))
                return quick_color("red")
            elif self._data[row][5] and index.column() in (5, 6, 7, 8):
                # return QBrush(QColor(151, 233, 110))
                return quick_color("green")
        elif role == Qt.DisplayRole:
            # if index.column() in (2, 4):
            #     value = number(float((self._data[index.row()][index.column()])))
            if index.column() in (3, 7):
                value = int((self._data[index.row()][index.column()]))
                value = minutes_to_hours_and_minutes(value)
            else:
                value = str(self._data[index.row()][index.column()])
            return value
        elif role == Qt.EditRole:
            if index.column() in (2, 3, 4, 7):
                value = float(self._data[index.row()][index.column()])
            elif index.column() == 1:
                value = int(str(self._data[index.row()][index.column()]).split()[-1])
            else:
                value = str(self._data[index.row()][index.column()])
            return value
        elif role == Qt.FontRole:
            if index.column() in (0, 1):
                font = QFont()
                font.setBold(True)
                return font
        else:
            return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.header_data[col]
        elif role == Qt.FontRole:
            if col in (0, 1):
                font = QFont()
                font.setBold(True)
                return font
        return QVariant()

    def flags(self, index):
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable


class PopUpView(QTableView):

    def __init__(self, parent=None):
        super().__init__()
        self.sorter = None

    def set_model(self, data):

        self.sorter = QSortFilterProxyModel()
        self.sorter.setDynamicSortFilter(True)
        self.sorter.setSortRole(Qt.EditRole)
        self.sorter.setSourceModel(PopUpModel(data))
        self.setModel(self.sorter)
        self.resizeColumnsToContents()
        self.setSortingEnabled(True)


class PopUpToPandas:

    def __init__(self, schedules):

        labels = ["Source", "Sched No.", "O. Stops", "O. Duration", "O. Miles", "Postalized", "Stops",
                  "Duration", "Miles", "Trips", "Schedule Type", "Freq Code"]

        compiled = []
        for schedule in schedules:
            num = schedule.schedule_name
            trip_num = num.split()[-1]
            plate_num = schedule.source_file
            num_stops = len(schedule.original_stops)
            duration = float(schedule.original_duration)
            mileage = round(schedule.original_mileage, 1)
            num_stops2 = len(schedule.postalized_stops)
            duration2 = float(schedule.postalized_duration)
            mileage2 = round(schedule.postalized_mileage, 1)
            postalized = schedule.is_postalized
            num_round_trips = len(schedule.round_trips)
            type = schedule.schedule_type
            if type == 1:
                spotter = "Real"
            elif type == 2:
                spotter = "Lunch"
            elif type == 3:
                spotter = "Spotter"
            elif type == 4:
                spotter = "Standby"
            else:
                spotter = "?????"
            freq_code = schedule.freq_code
            compiled.append([plate_num, trip_num, num_stops, duration, mileage, postalized, num_stops2, duration2,
                             mileage2, num_round_trips, spotter, freq_code])

        self.df = pd.DataFrame(compiled, columns=labels)


class SchedulePopUpModel(QAbstractTableModel):

    def __init__(self, data, parent=None):
        QAbstractTableModel.__init__(self, parent)
        self._data = data.values
        self.header_data = data.columns

    def rowCount(self, parent=None):
        if not self._data[0][0]:
            return 0
        return len(self._data)

    def columnCount(self, parent=None):
        return len(self._data[0])

    def data(self, index, role):
        if not index.isValid():
            return None
        elif role == Qt.BackgroundRole:
            row=index.row()
            if not self._data[row][1] or not self._data[row][2]:
                return quick_color("red")
        elif role != Qt.DisplayRole:
            return None

        value = str(self._data[index.row()][index.column()])

        return value

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.header_data[col]
        return QVariant()

    def flags(self, index):
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable


class SchedulePopUpView(QTableView):

    def __init__(self, parent=None):
        super().__init__()
        self.sorter = None

    def set_model(self, data):
        self.setModel(SchedulePopUpModel(data))
        self.resizeColumnsToContents()


class SchedulePopUpToPandas:

    def __init__(self, schedule, stop_choice):

        labels = ["Stop", "Stop Name", "Arrive Time", "Depart Time"]

        if stop_choice == 1:
            stops = schedule.original_stops
        elif stop_choice == 2:
            stops = schedule.postalized_stops
        else:
            stops = schedule.stops

        compiled = []
        for x, stop in enumerate(stops):
            stop_number = str(x+1)
            stop_name = stop.stop_name
            arrive_time = stop.arrive_time
            depart_time = stop.depart_time
            compiled.append([stop_number, stop_name, arrive_time, depart_time])

        self.df = pd.DataFrame(compiled, columns=labels)


class SingleRoundTripToPandas:

    def __init__(self, round_trip):

        labels = ["Stop", "Stop Name", "Arrive Time", "Depart Time"]
        compiled = []

        for x, stop in enumerate(round_trip.stops):
            stop_number = str(x+1)
            stop_name = stop.stop_name
            arrive_time = stop.arrive_time
            depart_time = stop.depart_time
            compiled.append([stop_number, stop_name, arrive_time, depart_time])

        self.df = pd.DataFrame(compiled, columns=labels)


class RoundTripModel(QAbstractTableModel):

    def __init__(self, data, parent=None):
        QAbstractTableModel.__init__(self, parent)
        self._data = data.values
        self.header_data = data.columns

    def rowCount(self, parent=None):
        if not self._data[0][0]:
            return 0
        return len(self._data)

    def columnCount(self, parent=None):
        return len(self._data[0])

    def data(self, index, role):
        if not index.isValid():
            return None
        else:
            row = index.row()
            col = index.column()
            val = self._data[row][col]

        if role == Qt.DisplayRole:
            if col == 6:
                return minutes_to_hours_and_minutes(val)
            if col == 19:
                return None
            if col in (11, 12, 13, 14, 15, 16, 17):
                if val:
                    return "X"
                else:
                    return ""
        elif role == Qt.TextAlignmentRole:
            if col in (11, 12, 13, 14, 15, 16, 17):
                return Qt.AlignCenter
            else:
                return None
        elif role == Qt.BackgroundRole:
            check = self._data[row][9]
            hol = self._data[row][18]
            if check != "Real trip":
                return quick_color("grey")
            if hol:
                return quick_color("grey")
            else:
                return None
        elif col == 19:
            if role == Qt.CheckStateRole:
                if val:
                    return Qt.Checked
                else:
                    return Qt.Unchecked
            else:
                return None

        elif role != Qt.DisplayRole:
            return None

        value = str(self._data[index.row()][index.column()])
        return value

    def setData(self, index, value, role):

        row = index.row()
        col = index.column()

        if col == 19:
            if role == Qt.CheckStateRole:
                self.dataChanged.emit(index, index, [])

        if role == Qt.EditRole:
            self._data[row][col] = value
            self.dataChanged.emit(index, index, [])

        super().setData(index, value, role)
        return True

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.header_data[col]
        return QVariant()

    def flags(self, index):
        if index.column() == 19:
            return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable


class RoundTripFilter(QSortFilterProxyModel):

    def __init__(self, *args, **kwargs):
        QSortFilterProxyModel.__init__(self, *args, **kwargs)
        self.filters = {}

    def setFilterByColumn(self, regex, column):
        self.filters[column] = regex
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        all_bools = []
        for key, regex in self.filters.items():
            ix = self.sourceModel().index(source_row, key, source_parent)
            if ix.isValid():
                text = self.sourceModel().data(ix, role=Qt.DisplayRole)
                if regex == "":
                    all_bools.append(True)
                elif not text == regex:
                    all_bools.append(False)
        return False not in all_bools


class RoundTripView(QTableView):

    def __init__(self, parent=None):
        super().__init__()
        self.sorter = None

    def set_model(self, data):

        self.sorter = RoundTripFilter()
        self.sorter.setDynamicSortFilter(True)
        self.sorter.setSourceModel(RoundTripModel(data))
        self.setModel(self.sorter)
        self.resizeColumnsToContents()


class RoundTripsToPandas:

    def __init__(self, round_trips):

        labels = ["Source", "Sched #", "Trip", "Stops", "Start", "End", "Duration", "Has Lunch", "Vehicle", "Trip Type",
                  "Frequency", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", "Hol", "Include"]

        compiled = []

        for x, round_trip in enumerate(round_trips):
            plate_name = round_trip.source_file
            schedule_num = round_trip.schedule_num
            trip = round_trip.trip_num
            stops = len(round_trip.stops)
            start = round_trip.stops[0].arrive_time
            end = round_trip.stops[-1].depart_time
            duration = round_trip.duration
            has_lunch = round_trip.contains_lunch
            trip_type = round_trip.trip_type
            trip_list = ["Real trip", "Lunch trip", "Spotter trip", "Standby trip"]
            removable = trip_list[trip_type - 1]
            vehicle = round_trip.vehicle_type
            freq_code = round_trip.freq_code
            bin_string = round_trip.bin_string
            mon = bin_string[0] in (1, "1")
            tue = bin_string[1] in (1, "1")
            wed = bin_string[2] in (1, "1")
            thu = bin_string[3] in (1, "1")
            fri = bin_string[4] in (1, "1")
            sat = bin_string[5] in (1, "1")
            sun = bin_string[6] in (1, "1")
            holiday = round_trip.holiday
            selected = round_trip.is_selected
            compiled.append([plate_name, schedule_num, trip, stops, start, end, duration, has_lunch, vehicle,
                             removable, freq_code, mon, tue, wed, thu, fri, sat, sun, holiday, selected])
        if len(round_trips) == 0:
            compiled.append([None, None, None, None, None, None, None, None, None, None, None,
                             None, None, None, None, None, None, None, None, None])

        self.df = pd.DataFrame(compiled, columns=labels)
