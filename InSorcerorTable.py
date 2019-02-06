
import pandas as pd
from PyQt5.QtCore import QAbstractTableModel, Qt, QVariant, QSortFilterProxyModel
from PyQt5.QtWidgets import QTableView, QAbstractItemView
from PyQt5.QtGui import QFont
from GeneralMethods import big_money, number, minutes_to_hours_and_minutes, quick_color, quick_bold


class CompilerModel(QAbstractTableModel):

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
            if self._data[row][4] in (-1, "-1") or self._data[row][0] in ("None", None, "U/K"):
                return quick_color("grey")
            elif not self._data[row][2]:
                return quick_color("red")
            elif not self._data[row][3]:
                return quick_color("yellow")
        elif role == Qt.DisplayRole:
            value = self._data[index.row()][index.column()]
            if index.column() in (9, 10, 11, 12):
                if value in (-1, "-1"):
                    return "n/a"
                if index.column() in (10, 11, 12) and self._data[index.row()][0] in (None, "None", "U/K"):
                    return "U/K"
                if index.column() in (10, 11, 12):
                    val1 = float(self._data[index.row()][9])
                    val2 = float(value)
                    diff = big_money(val2 - val1)
                    return big_money(val2) + "; " + diff
                value = big_money(value)
            if index.column() == 1:
                value = str(value)[:5]
            if index.column() == 0:
                if value in (None, "None"):
                    return "U/K"
            if index.column() in (2, 3):
                if self._data[index.row()][0] in (None, "None", "U/K"):
                    return "U/K"
            return str(value)
        elif role == Qt.EditRole:
            value = self._data[index.row()][index.column()]
            if index.column() == 4:
                value = float(value)
            if index.column() in (10, 11, 12):
                val1 = float(self._data[index.row()][4])
                value = float(value)
                return value - val1
            else:
                value = str(value)
            return value
        elif role == Qt.FontRole:
            if index.column() in (0, 1):
                return quick_bold()
        elif role == Qt.TextAlignmentRole:
            if index.column() in (0, 1):
                return Qt.AlignCenter
        elif role == Qt.ForegroundRole:
            if index.column() in (10, 11, 12):
                val1 = float(self._data[index.row()][4])
                value = float(self._data[index.row()][index.column()])
                if value < val1:
                    return quick_color("dark red")
        else:
            return None

    def headerData(self, col, orientation, role):

        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.header_data[col]
        if orientation == Qt.Horizontal and role == Qt.FontRole:
            font = QFont()
            font.setBold(True)
            return font
        if role == Qt.ToolTipRole:
            if col == 0:
                return "PVS site associated with this contract"
            if col == 1:
                return "HCR Plate Number, Contract Number, HCR ID"
            if col == 2:
                return "Checks each schedule on plate against maximum duration and maximum mileage, " \
                       "and plate against minimum cost threshold"
            if col == 3:
                return "Checks if each schedule on plate can be made to comply with postal rules"
            if col == 9:
                return "Current annual contract rate"
            if col == 10:
                return "Estimated in-sourced cost with full acquisition cost of new fleet"
            if col == 11:
                return "Estimated in-sourced cost with depreciated cost of new fleet"
            if col == 12:
                return "Estimated in-sourced cost with leased fleet"

        return QVariant()

    def flags(self, index):
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable


class CompilerView(QTableView):

    def __init__(self, parent=None):
        super().__init__()
        self.sorter = None

    def set_model(self, data):

        self.sorter = QSortFilterProxyModel()
        self.sorter.setDynamicSortFilter(True)
        self.sorter.setSortRole(Qt.EditRole)
        self.sorter.setFilterKeyColumn(1)
        self.sorter.setSourceModel(CompilerModel(data))
        self.setModel(self.sorter)
        self.resizeColumnsToContents()
        self.setColumnWidth(0, max(100, self.columnWidth(0)))
        self.setColumnWidth(1, max(100, self.columnWidth(1)))
        for x in range (4, 8):
            self.setColumnWidth(x, max(self.columnWidth(x), 100))
        self.setSortingEnabled(True)
        self.setSelectionBehavior(QAbstractItemView.SelectRows)


class CompilersToPandas:

    def __init__(self, compilers, case):

        labels = ["PVS Site", "Plate Num", "Eligible", "Postalizable", "# Trips", "# Insourceable",
                  "# Postalizable", "# Network Trips", "# One-State", "HCR Cost", "Cost w/o Dep",
                  "Cost w/ Dep", "Cost w/ Lease"]

        compiled = []

        for compiler in compilers:
            if case == 1:
                compiled.append(compiler.output_list)
            if case == 2:
                compiled.append(compiler.output_list_postalized)
        if len(compilers) == 0:
            compiled.append([None, None, "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0])

        self.df = pd.DataFrame(compiled, columns=labels)


class PopUpModel(QAbstractTableModel):

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
            if not self._data[row][4]:
                return quick_color("red")
            elif not self._data[row][6]:
                return quick_color("yellow")
        elif role == Qt.DisplayRole:
            if index.column() == (1, 3):
                value = number(float((self._data[index.row()][index.column()])))
            elif index.column() in (2, 6):
                value = minutes_to_hours_and_minutes(int(self._data[index.row()][index.column()]))
            elif index.column() == 0:
                value = str(self._data[index.row()][index.column()])
                plate_num = value[:5]
                schedule_num = value.split()[-1]
                value = plate_num + " " + schedule_num
            else:
                value = str(self._data[index.row()][index.column()])
            return value
        elif role == Qt.EditRole:
            if index.column() in (1, 2, 3, 9):
                value = float(self._data[index.row()][index.column()])
            elif index.column() == 0:
                value = int(str(self._data[index.row()][index.column()]).split()[-1])
            else:
                value = str(self._data[index.row()][index.column()])
            return value
        elif role == Qt.TextAlignmentRole:
            if index.column() == 0:
                return Qt.AlignCenter
        elif role == Qt.FontRole:
            if index.column() == 0:
                return quick_bold()
        else:
            return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.header_data[col]
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
        self.setSelectionBehavior(QAbstractItemView.SelectRows)


class PopUpToPandas:

    def __init__(self, schedules):

        labels = ["Schedule Number", "# Stops", "Duration", "Mileage", "Insource eligible", "Reason(s)",
                  "Postal Duration", "Postal Mileage", "Postalizable", "Reason(s)"]

        compiled = []
        for schedule in schedules:
            num = schedule.schedule_name
            num_stops = len(schedule.original_stops)
            duration = float(schedule.original_duration)
            mileage = round(schedule.original_mileage, 1)
            insource_eligible = schedule.is_eligible
            reasons1 = schedule.cant_eligible
            duration2 = float(schedule.postalized_duration)
            mileage2 = round(schedule.postalized_mileage, 1)
            reason1_string = ""
            if len(reasons1) > 0:
                for reason in reasons1:
                    reason1_string += reason + ", "
                reason1_string = reason1_string[:-2]
            postalizable = schedule.can_postalize
            reasons = schedule.cant_postalize
            reason_string = ""
            if len(reasons) > 0:
                for reason in reasons:
                    reason_string += reason + ", "
                reason_string = reason_string[:-2]
            compiled.append([num, num_stops, duration, mileage, insource_eligible, reason1_string,
                             duration2, mileage2, postalizable, reason_string])

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
