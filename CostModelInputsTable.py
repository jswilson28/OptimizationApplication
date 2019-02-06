from PyQt5.QtCore import QAbstractTableModel, Qt, QVariant, QSortFilterProxyModel
from PyQt5.QtWidgets import QTableView, QAbstractItemView
from PyQt5.QtGui import QBrush, QColor, QFont
from GeneralMethods import money, quick_color, percent, number


def read_a_rectangle(ws, x_start, x_end, y_start, y_end):
    return_list = []
    for x in range(x_start, x_end):
            row = []
            for y in range(y_start, y_end):
                row.append(ws.cell(row=x, column=y).value)
            return_list.append(row)
    return return_list


def read_a_rectangle_transpose(ws, y_start, y_end, x_start, x_end):
    return_list = []
    for y in range(y_start, y_end):
        row = []
        for x in range(x_start, x_end):
            row.append(ws.cell(row=x, column=y).value)
        return_list.append(row)
    return return_list


class CostInputsModel(QAbstractTableModel):

    def __init__(self, data, headers, table_num, parent=None):
        QAbstractTableModel.__init__(self, parent)
        self._data = data
        self.header_data = headers
        self.table_num = table_num

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

        if self.table_num == 1:
            if col > 0:
                if role == Qt.DisplayRole:
                    return money(val)
                elif role == Qt.ForegroundRole:
                    return QBrush(quick_color("blue"))
                elif role == Qt.EditRole:
                    return float(val)
            elif role == Qt.DisplayRole:
                return str(val)
        if self.table_num == 2:
            if col > 0:
                if role == Qt.DisplayRole:
                    return percent(val)
                elif role == Qt.ForegroundRole:
                    return QBrush(quick_color("blue"))
                elif role == Qt.EditRole:
                    return float(val)
            elif role == Qt.DisplayRole:
                return str(val)
        if self.table_num == 3:
            if col > 0:
                if role == Qt.DisplayRole:
                    if row in (0, 2, 9):
                        return money(val)
                    if row in (1, 3, 4, 5, 6):
                        return percent(val)
                    if row in (7, 8):
                        return number(val)
                    if row in (10, 11):
                        return int(val)
                if role == Qt.ForegroundRole:
                    return QBrush(quick_color("blue"))
                if role == Qt.EditRole:
                    if row in (7, 8, 10, 11):
                        return int(val)
                    else:
                        return float(val)
            elif role == Qt.DisplayRole:
                return str(val)
        if self.table_num == 4:
            if col > 0:
                if role == Qt.DisplayRole:
                    if col == 1:
                        return money(val)
                    if col == 2:
                        if not val:
                            return "n/a"
                        else:
                            return int(val)
                    if col == 3:
                        if not val:
                            return "n/a"
                        else:
                            return money(val)
                if role == Qt.ForegroundRole:
                    if col in (1, 2):
                        return QBrush(quick_color("blue"))
                if role == Qt.EditRole:
                    if col == 1:
                        return float(val)
                    if col == 2:
                        if not val:
                            return "n/a"
                        else:
                            return int(val)
                    if col == 3:
                        if not val:
                            return "n/a"
                        else:
                            return float(val)
            elif role == Qt.DisplayRole:
                return str(val)
        if self.table_num == 5:
            if col > 0:
                if role == Qt.ForegroundRole:
                    return QBrush(quick_color("blue"))
                if role == Qt.DisplayRole:
                    if col in (1, 2):
                        return money(val)
                    if col == 3:
                        if val != "n/a":
                            return number(val)
                        else:
                            return "n/a"
                    if col == 4:
                        if val != "n/a":
                            return money(val)
                        else:
                            return "n/a"
                if role == Qt.EditRole:
                    if col in (1, 2):
                        return float(val)
                    if col == 3:
                        if val != "n/a":
                            return int(val)
                        else:
                            return "n/a"
                    if col == 4:
                        if val != "n/a":
                            return float(val)
                        else:
                            return "n/a"
            elif role == Qt.DisplayRole:
                return str(val)
        if self.table_num in (6, 7, 8):
            if col > 0:
                if role == Qt.ForegroundRole:
                    return QBrush(quick_color("blue"))
                if role == Qt.DisplayRole:
                    return money(val)
                if role == Qt.EditRole:
                    return float(val)
            elif role == Qt.DisplayRole:
                return str(val)
        elif role == Qt.DisplayRole:
            return str(self._data[index.row()][index.column()])
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.header_data[col]
        return None

    def flags(self, index):

        if self._data[index.row()][index.column()] not in  (None, "n/a"):
            if self.table_num != 4:
                if index.column() > 0:
                    return Qt.ItemIsEnabled | Qt.ItemIsSelectable  #| Qt.ItemIsEditable
            else:
                if index.column() in (1, 2):
                    return Qt.ItemIsEnabled | Qt.ItemIsSelectable  #| Qt.ItemIsEditable
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable


class CostInputsView(QTableView):

    def __init__(self):
        super().__init__()

    def set_model(self, data, headers, table_num):
        self.setModel(CostInputsModel(data, headers, table_num))
        self.resizeColumnsToContents()


class CostInputsToLists:

    def __init__(self, ws):

        self.ws = ws

        self.hourly_wages = []
        self.hourly_wages_headers = []
        self.labor_splits = []
        self.labor_splits_headers = []
        self.other_inputs = []
        self.other_inputs_headers = []
        self.vehicle_acquisition_costs = []
        self.vehicle_acquisition_costs_headers= []
        self.vehicle_lease_costs = []
        self.vehicle_lease_costs_headers = []
        self.regional_cpm = []
        self.regional_cpm_headers = []
        self.trailer_maint = []
        self.trailer_maint_headers = []
        self.fuel_inputs = []
        self.fuel_inputs_headers = []

        self.set_headers()
        self.read_in_data()

    def set_headers(self):

        self.hourly_wages_headers = ["Employee Type", "Hourly Wage Rate", "Fully Loaded Wage Rate",
                                     "Night Differential"]
        self.labor_splits_headers = ["Labor Inputs", "Hours Split"]
        self.other_inputs_headers = ["Input", "Value"]
        self.vehicle_acquisition_costs_headers = ["Vehicle Type", "Total Acquisition Cost",
                                                  "Years Until Fully Depreciated", "Annual Acquisition Cost"]
        self.vehicle_lease_costs_headers = ["Vehicle Type", "1-Year EBuy Monthly Cost", "Annual Leasing Cost",
                                            "Maximum Mileage", "Extra Mileage Charge"]
        self.regional_cpm_headers = ["Region", "Cargo Van", "Tractor (SA)", "Tractor (TA)"]
        self.trailer_maint_headers = ["Region", "Trailer"]
        self.fuel_inputs_headers = ["Vehicle Type", "Annual Value"]

    def read_in_data(self):

        ws = self.ws

        self.hourly_wages = read_a_rectangle(ws, 8, 12, 2, 6)
        self.labor_splits = read_a_rectangle(ws, 14, 17, 2, 4)
        self.other_inputs = read_a_rectangle(ws, 19, 31, 2, 4)
        self.vehicle_acquisition_costs = read_a_rectangle(ws, 33, 37, 2, 5)
        for row in self.vehicle_acquisition_costs:
            if row[2]:
                row.append(row[1]/row[2])
            else:
                row.append(None)
        self.vehicle_lease_costs = read_a_rectangle(ws, 39, 43, 2, 7)
        self.regional_cpm = read_a_rectangle_transpose(ws, 3, 11, 45, 49)
        self.trailer_maint = read_a_rectangle_transpose(ws, 3, 11, 51, 53)
        self.fuel_inputs = read_a_rectangle(ws, 55, 58, 2, 4)

    def get_a_table(self, table_num):

        if table_num == 1:
            return self.hourly_wages, self.hourly_wages_headers
        elif table_num == 2:
            return self.labor_splits, self.labor_splits_headers
        elif table_num == 3:
            return self.other_inputs, self.other_inputs_headers
        elif table_num == 4:
            return self.vehicle_acquisition_costs, self.vehicle_acquisition_costs_headers
        elif table_num == 5:
            return self.vehicle_lease_costs, self.vehicle_lease_costs_headers
        elif table_num == 6:
            return self.regional_cpm, self.regional_cpm_headers
        elif table_num == 7:
            return self.trailer_maint, self.trailer_maint_headers
        elif table_num == 8:
            return self.fuel_inputs, self.fuel_inputs_headers
        else:
            return [None], [None]