from PyQt5.QtCore import QAbstractTableModel, Qt, QTime
from PyQt5.QtWidgets import QTableView
from PyQt5.QtGui import QBrush
from GeneralMethods import quick_color, minutes_to_hours_and_minutes, quick_bold, quick_italic


class PostalizerInputsModel(QAbstractTableModel):

    def __init__(self, data, headers, parent=None):
        QAbstractTableModel.__init__(self, parent)
        self._data = data
        self.header_data = headers

    def rowCount(self, parent=None):
        return len(self._data)

    def columnCount(self, parent=None):
        return len(self._data[0])

    def data(self, index, role):
        row = index.row()
        col = index.column()

        if not index.isValid():
            return None
        else:
            val = self._data[row][col]

        if role == Qt.DisplayRole:
            if val == None:
                return ""
            elif col == 0 and row not in (0, 4, 10, 15):
                return "    " + val
            elif row in (17, 18, 16) and col == 1:
                return minutes_to_hours_and_minutes(float(val))
            elif row in (13, 14, 12) and col == 1:
                return val
            else:
                return str(val)
        if role == Qt.EditRole:
            if row in (17, 18, 16) and col == 1:
                return int(val)
            if row in (13, 14, 12) and col == 1:
                return val
            if row == 7:
                return bool(val)
            if row in (1, 2, 3, 5, 6, 8, 9, 11):
                return int(val)
        if role == Qt.FontRole:
            if row in (0, 4, 10, 15):
                return quick_bold()
            elif col == 2:
                return quick_italic()
        if role == Qt.ForegroundRole:
            if col == 1:
                return QBrush(quick_color("blue"))
        if role == Qt.TextAlignmentRole:
            if col == 1:
                return Qt.AlignCenter

        return None

    def setData(self, index, value, role):
        row = index.row()
        col = index.column()

        if role == Qt.EditRole:
            self._data[row][col] = value

        return True

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.header_data[col]
        return None

    def flags(self, index):
        if index.column() == 1 and self._data[index.row()][index.column()] != None:
            return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable


class PostalizerInputsView(QTableView):

    def __init__(self):
        super().__init__()

    def set_model(self, data, headers):

        self.setModel(PostalizerInputsModel(data, headers))
        self.resizeColumnsToContents()


class PostalizerInputsToLists:

    def __init__(self):
        # trip inputs
        self.pvs_to_pdc = None
        self.pvs_time = None
        self.pdc_time = None

        # lunch inputs
        self.lunch_duration = None
        self.hours_wo_lunch = None
        self.allow_non_postal = None
        self.lunch_travel_time = None
        self.lunch_buffer_time = None

        # other requirements
        self.max_working_time = None
        self.tour_one = None
        self.nd_pm = None
        self.nd_am = None

        # u_part combination sensitivity
        self.min_check_minutes = None
        self.max_check_minutes = None
        self.max_combined_minutes = None
        self.use_check_codes = None

        self.inputs = []
        self.inputs_headers = ["Variable", "Value", "Description"]
        self.set_defaults()
        self.set_inputs()

    def set_defaults(self):

        # trip inputs
        self.pvs_to_pdc = 1
        self.pvs_time = 14
        self.pdc_time = 10

        # lunch inputs
        self.lunch_duration = 30
        self.hours_wo_lunch = 6
        self.allow_non_postal = True
        self.lunch_travel_time = 5
        self.lunch_buffer_time = 10

        # other requirements
        self.max_working_time = 8
        self.tour_one = QTime(20, 0, 0)
        self.nd_pm = QTime(18, 0, 0)
        self.nd_am = QTime(6, 0, 0)

        # u_part combination sensitivity
        self.min_check_minutes = 60
        self.max_check_minutes = 480
        self.max_combined_minutes = 1440
        self.use_check_codes = True

    def set_inputs(self):

        rows = []

        desc1 = "Minutes spent at PVS site at beginning and end of schedules."
        desc2 = "Minutes spent en route between PVS site and P&DC site."
        desc3 = "Minutes spent at P&DC site at beginning and end of round trips."
        desc4 = "Duration of lunch in minutes."
        desc5 = "Consecutive working hours allowed without lunch break."
        desc6 = "Waive requirement that lunch breaks occur at USPS locations."
        desc61 = "Travel time required before and after lunch (paid)."
        desc62 = "Buffer time required before and after lunch (paid)."
        desc7 = "Maximum time allowed in a single schedule (not including lunch)."
        desc8 = "Start time of USPS work days."
        desc9 = "Start time of night differential pay."
        desc10 = "End time of night differential pay."
        desc11 = "HCR 'U-Parts' with a layover lower than this limit will be combined."
        desc12 = "HCR 'U-Parts' with a layover greater than this limit will not be combined."
        desc13 = "HCR 'U-Parts' with a combined duration greater than this limit will not be combined"

        rows.append(["Trip Inputs", None, None])
        rows.append(["Minutes at PVS", self.pvs_time, desc1])
        rows.append(["Minutes from PVS to P&DC", self.pvs_to_pdc, desc2])
        rows.append(["Minutes at P&DC", self.pdc_time, desc3])
        rows.append(["Lunch Inputs", None, None])
        rows.append(["Minutes for lunch", self.lunch_duration, desc4])
        rows.append(["Hours w/o lunch", self.hours_wo_lunch, desc5])
        rows.append(["Allow at non-postal sites", self.allow_non_postal, desc6])
        rows.append(["Lunch travel minutes", self.lunch_travel_time, desc61])
        rows.append(["Lunch buffer minutes", self.lunch_buffer_time, desc62])
        rows.append(["Other Inputs", None, None])
        rows.append(["Maximum working hours", self.max_working_time, desc7])
        rows.append(["Tour one start time", self.tour_one, desc8])
        rows.append(["Night differential start", self.nd_pm, desc9])
        rows.append(["Night differential end", self.nd_am, desc10])
        rows.append(["U-Part merger sensitivity", None, None])
        rows.append(["Layover minutes minimum", self.min_check_minutes, desc11])
        rows.append(["Layover minutes maximum", self.max_check_minutes, desc12])
        rows.append(["Combined maximum duration", self.max_combined_minutes, desc13])

        self.inputs = rows

    def get_lists(self):

        return self.inputs, self.inputs_headers