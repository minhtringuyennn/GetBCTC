# Import Python library
import csv, io, math
from datetime import date, datetime
from PyQt5 import QtWidgets, QtCore, QtGui, uic
from PyQt5.QtCore import QAbstractTableModel, Qt

# Import custom modules
import Utils
from Handle import FetchData

# Import pyqt GUI
from GUI.GUI import Ui_Client

# Align cells
class AlignDelegate(QtWidgets.QStyledItemDelegate):
    def initStyleOption(self, option, index):
        super(AlignDelegate, self).initStyleOption(option, index)
        option.displayAlignment = (QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)

# Convert pandas dataframe to pyqt table
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
        self.ui.financeStatementTable.installEventFilter(self)
        self.show()

    # Load UI function
    def LoadFunction(self):
        self.ui.updateButton.clicked.connect(self.QueryBCTC)
        self.ui.exportButton.clicked.connect(self.ExportBCTC)
        
        now = str(date.today()).split("-")
        self.ui.getDateButton.setDateTime(QtCore.QDateTime(
            QtCore.QDate(int(now[0]), int(now[1]), int(now[2])), QtCore.QTime(0, 0, 0)))
    
    # Add custom event filter
    def eventFilter(self, source, event):
        if (event.type() == QtCore.QEvent.KeyPress and
            event.matches(QtGui.QKeySequence.Copy)):
            self.copySelection()
            return True
        return super(ClientUI, self).eventFilter(source, event)

    # Add custom copy method
    def copySelection(self):
        selection = self.ui.financeStatementTable.selectedIndexes()
        
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
            csv.writer(stream, delimiter='\t').writerows(table)
            QtWidgets.qApp.clipboard().setText(stream.getvalue())

    def QueryBCTC(self):
        searchField     = Utils.no_accent_vietnamese(self.ui.searchSymbolField.text().upper())
        typeFinField    = self.ui.typeFinanceField.currentText().lower()
        typeCurrField   = self.ui.typeCurrencyField.currentText().lower()
        getDate         = self.ui.getDateButton.date().toPyDate()
        yearCheckBox    = self.ui.isYearCheckBox.isChecked()
        numYearQuarter  = self.ui.numberOfYearQuarterrField.value()
        
        if len(self.ui.searchSymbolField.text()) < 3:
            return
        
        res = FetchData(searchField, typeFinField, typeCurrField, getDate, yearCheckBox, numYearQuarter)

        self.Data = res.fetchBCTC()
        # print(self.Data)
        
        if self.Data is not None:
            label = f"Bảng {typeFinField} của {searchField}. Đơn vị: {typeCurrField}."
            self.ui.financeStatementLabel.setText(label)
            self.ui.financeStatementTable.setModel(PandasModel(self.Data))
            self.ui.financeStatementTable.setAlternatingRowColors(True)
            self.ui.financeStatementTable.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
            
            delegate = AlignDelegate(self.ui.financeStatementTable)
            for i in range(1, numYearQuarter + 1):
                self.ui.financeStatementTable.setItemDelegateForColumn(i, delegate)
        else:
            QtWidgets.QMessageBox.critical(None, "Lỗi", "Không tìm thấy dữ liệu!")
            
    def ExportBCTC(self):
        symbol      = Utils.no_accent_vietnamese(self.ui.searchSymbolField.text().upper())
        typeFin     = Utils.no_accent_vietnamese(self.ui.typeFinanceField.currentText()).replace(" ", "")
        getDate     = self.ui.getDateButton.date().toPyDate()
        getDate     = str(getDate)
        bctcYear    = getDate[:-6]
        bctcQuarter = math.ceil(int(getDate[5:-3])/3)
        
        fileName = f"{symbol}_{typeFin}_{bctcYear}Q{bctcQuarter}.xlsx"
        
        self.QueryBCTC()
        
        if self.Data is not None:
            self.Data.to_excel(fileName)
            QtWidgets.QMessageBox.about(None, "Thông báo", "Đã xuất file thành công tại thư mục chạy .exe hiện thời")