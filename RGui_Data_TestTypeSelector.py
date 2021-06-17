import sys
import os
import pandas as pd
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QDialog
from PyQt5.uic import loadUi
from Screening_System_PyQt5 import RGui_Screening_DataCollection_WithFunctions

pd.options.display.max_columns = 999
pd.options.display.max_rows = 999
pd.set_option("display.precision", 6)


class DataTestTypeSelector(QDialog):
    def __init__(self, test_name):
        super(DataTestTypeSelector, self).__init__()
        loadUi("RGui_Type_Select.ui", self)
        dir_path = os.path.dirname(sys.argv[0])
        self.temp_file = pd.read_csv(dir_path + r"\\Screening_Template\\" + test_name, sep='\t')
        self.test_name = test_name
        self.bare_cell = False
        self.label.setText(f"Test Name: {test_name}")
        self.populate_radio_button()
        self.pushButton_3.clicked.connect(self.ok_pressed)
        self.pushButton.clicked.connect(self.cancel_pressed)

    def cancel_pressed(self):
        self.close()

    def ok_pressed(self):
        ui = RGui_Screening_DataCollection_WithFunctions.Screening_DataCollection_WithFunction()
        if self.radioButton_1.isChecked():
            print(self.test_name)
            ui.getTestNumber(self.test_name, self.bare_cell)
        else:
            ui.getTestNumber("Post" + self.test_name, self.bare_cell)
        self.close()
        ui.show()
        ui.exec_()

    def populate_radio_button(self):
        # this for tabbing
        if self.temp_file["Tabbed?"][0] == 'Tabbed':
            self.radioButton_1.setText("Pre Tab")
            self.radioButton_2.setText("Post Tab")
        elif self.temp_file["Profile No."][0] == 2:
            self.radioButton_1.setText("Profile one")
            self.radioButton_2.setText("Profile two")
        elif self.temp_file["Section No."][0] == 2:
            self.radioButton_1.setText("Section one")
            self.radioButton_2.setText("Section two")
        else:
            self.radioButton_1.setText("Bare cell")
            self.radioButton_2.setEnabled(False)
            self.radioButton_1.setChecked(True)
            self.bare_cell = True


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    qt_app = DataTestTypeSelector("14872-00.txt")
    qt_app.show()
    app.exec_()