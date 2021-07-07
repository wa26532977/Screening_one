import os
from PyQt5.QtWidgets import QDialog
from PyQt5 import QtWidgets
from PyQt5.uic import loadUi
import pandas as pd
import sys
pd.options.display.max_columns = 999
pd.options.display.max_rows = 999
pd.set_option("display.precision", 6)


class SetPathWithFunction(QDialog):
    def __init__(self):
        super(SetPathWithFunction, self).__init__()
        loadUi("RGui_SetPath.ui", self)
        self.p_path = os.path.dirname(sys.argv[0])
        self.label.setText(self.p_path + r"/Screening_Template")
        self.label_2.setText(self.p_path + r"/Screening_Data")
        self.label_3.setText(self.p_path + r"/Final_Report")
        self.label_4.setText(self.p_path + r"/Fast_data_collection")
        self.toolButton.clicked.connect(self.open_template)
        self.toolButton_2.clicked.connect(self.open_data)
        self.toolButton_3.clicked.connect(self.open_report)
        self.toolButton_4.clicked.connect(self.open_fast_data)

    def open_template(self):
        os.startfile(self.p_path + r"\\Screening_Template")

    def open_data(self):
        os.startfile(self.p_path + r"/Screening_Data")

    def open_report(self):
        os.startfile(self.p_path + r"/Final_Report")

    def open_fast_data(self):
        os.startfile(self.p_path + r"/Fast_data_collection")


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    qt_app = SetPathWithFunction()
    qt_app.show()
    app.exec_()