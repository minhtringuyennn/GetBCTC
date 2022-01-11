# Import Python library
import csv, io
from datetime import date, datetime
from unidecode import unidecode
from PyQt5 import QtWidgets, QtCore, QtGui, uic
from PyQt5.QtCore import QAbstractTableModel, Qt

# Import custom modules
from Handle import FetchData

# Import pyqt GUI
from GUI.GUI import Ui_Client

class AlignDelegate(QtWidgets.QStyledItemDelegate):
    def initStyleOption(self, option, index):
        super(AlignDelegate, self).initStyleOption(option, index)
        option.displayAlignment = (QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)

class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        QAbstractTableModel.__init__(self)
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parnet=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[col]
        return None

# Handle Client UI
class ClientUI(QtWidgets.QDialog):
    # Initalize UI
    def __init__(self):
        super(ClientUI, self).__init__()
        
        self.ui = Ui_Client()
        self.ui.setupUi(self)
        self.setWindowFlags(
            QtCore.Qt.WindowMinimizeButtonHint |
            QtCore.Qt.WindowMaximizeButtonHint |
            QtCore.Qt.WindowCloseButtonHint
           )

        self.LoadFunction()
        self.Data = None
        self.ui.dataTable.installEventFilter(self)
        self.show()

    # Load UI function
    def LoadFunction(self):
        self.ui.searchButton.clicked.connect(self.QueryType)
        
        now = str(date.today()).split("-")
        self.ui.getDateButton.setDateTime(QtCore.QDateTime(
            QtCore.QDate(int(now[0]), int(now[1]), int(now[2])), QtCore.QTime(0, 0, 0)))
    
    # add event filter
    def eventFilter(self, source, event):
        if (event.type() == QtCore.QEvent.KeyPress and
            event.matches(QtGui.QKeySequence.Copy)):
            self.copySelection()
            return True
        return super(ClientUI, self).eventFilter(source, event)

    # add copy method
    def copySelection(self):
        selection = self.ui.dataTable.selectedIndexes()
        if selection:
            rows = sorted(index.row() for index in selection)
            columns = sorted(index.column() for index in selection)
            rowcount = rows[-1] - rows[0] + 1
            colcount = columns[-1] - columns[0] + 1
            table = [[''] * colcount for _ in range(rowcount)]
            for index in selection:
                row = index.row() - rows[0]
                column = index.column() - columns[0]
                table[row][column] = index.data()
            stream = io.StringIO()
            csv.writer(stream).writerows(table)
            QtWidgets.qApp.clipboard().setText(stream.getvalue())

    def handleButton(self):
        filters = (
            'CSV files (*.csv *.txt)',
            'Excel Files (*.xls *.xml *.xlsx *.xlsm)',
            )
        path, filter = QtWidgets.QFileDialog.getOpenFileName(
            self, 'Open File', '', ';;'.join(filters))
        if path:
            csv = filter.startswith('CSV')
            if csv:
                dataframe = pd.read_csv(path)
            else:
                dataframe = pd.read_excel(path)
            self.model.setRowCount(0)
            dateformat = '%m/%d/%Y'
            rows, columns = dataframe.shape
            for row in range(rows):
                items = []
                for column in range(columns):
                    field = dataframe.iat[row, column]
                    if csv and isinstance(field, str):
                        try:
                            field = pd.to_datetime(field, format=dateformat)
                        except ValueError:
                            pass
                    if isinstance(field, pd.Timestamp):
                        text = field.strftime(dateformat)
                        data = field.timestamp()
                    else:
                        text = str(field)
                        if isinstance(field, np.number):
                            data = field.item()
                        else:
                            data = text
                    item = QtGui.QStandardItem(text)
                    item.setData(data, QtCore.Qt.UserRole)
                    items.append(item)
                self.model.appendRow(items)
    
    def QueryType(self):
        searchField     = unidecode(self.ui.searchField.text().upper())
        typeFinField    = self.ui.typeFinanceField.currentText().lower()
        getDate         = self.ui.getDateButton.date().toPyDate()
        yearCheckBox    = self.ui.yearCheckBox.isChecked()
        yearQuarter     = self.ui.yearQuarterField.value()
        
        if len(self.ui.searchField.text()) < 3:
            return
        
        res = FetchData(searchField, typeFinField, getDate, yearCheckBox, yearQuarter)

        self.Data = res.fetchBCTC()
        print(self.Data)
        
        if self.Data is not None:
            label = f"Bảng {typeFinField} của {searchField}. Đơn vị: VNĐ."
            self.ui.financeLabel.setText(label)
            self.ui.dataTable.setModel(PandasModel(self.Data))
            self.ui.dataTable.setAlternatingRowColors(True)
            self.ui.dataTable.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
            
            delegate = AlignDelegate(self.ui.dataTable)
            for i in range(1, yearQuarter + 1):
                self.ui.dataTable.setItemDelegateForColumn(i, delegate)
        else:
            QtWidgets.QMessageBox.critical(None, "Lỗi", "Không tìm thấy dữ liệu!")