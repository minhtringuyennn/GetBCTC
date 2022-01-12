import sys 
from PyQt5 import QtWidgets
from UI import ClientUI

if __name__ == "__main__":
    application = QtWidgets.QApplication(sys.argv)
    windows = ClientUI()
    sys.exit(application.exec_())