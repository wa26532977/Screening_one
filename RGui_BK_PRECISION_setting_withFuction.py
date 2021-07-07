import os
import glob
import sys
from PyQt5.QtWidgets import QDialog
from PyQt5 import QtWidgets
from PyQt5.uic import loadUi
import pandas as pd
pd.options.display.max_columns = 999
pd.options.display.max_rows = 999
pd.set_option("display.precision", 6)


class BKPrecisionSettingWithFunction(QDialog):
    def __init__(self):
        super(BKPrecisionSettingWithFunction, self).__init__()
        loadUi("RGui_BK_PRECISION_setting.ui", self)
        self.setting_file_path = os.path.dirname(sys.argv[0]) + r'\\BK_Prescions_setting.csv'
        self.setting_file = pd.read_csv(self.setting_file_path, sep=',')
        self.lineEdit.setText(self.setting_file['com_port'].values[0])
        self.lineEdit_2.setText(str(self.setting_file['baurate'].values[0]))
        self.lineEdit_4.setText(str(self.setting_file['model'].values[0]))
        self.pushButton.clicked.connect(self.cancel_clicked)
        self.pushButton_3.clicked.connect(self.set_clicked)
        self.pushButton_2.clicked.connect(self.help_clicked)

    def help_clicked(self):
        os.startfile(os.path.dirname(sys.argv[0]) + r'\\Setup for BK8500.docx')

    def set_clicked(self):
        self.setting_file.iloc[0, 0] = self.lineEdit.text()
        self.setting_file.iloc[0, 1] = self.lineEdit_2.text()
        self.setting_file.iloc[0, 2] = self.lineEdit_4.text()
        self.setting_file.to_csv(self.setting_file_path, sep=',', index=False)
        self.close()

    def cancel_clicked(self):
        self.close()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    qt_app = BKPrecisionSettingWithFunction()
    qt_app.show()
    app.exec_()
