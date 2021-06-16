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


class AnalysisSearchWithFunction(QDialog):
    def __init__(self):
        super(AnalysisSearchWithFunction, self).__init__()
        loadUi("RGui_Anaylsis_search.ui", self)
        self.pushButton_3.clicked.connect(self.cancel_button)
        self.pushButton_4.clicked.connect(self.search_button)
        self.pushButton_2.clicked.connect(self.clear_button)
        self.pushButton.clicked.connect(self.go_to_file_button)

    def go_to_file_button(self):
        while self.listWidget.currentItem() is None:
            msgbox = QtWidgets.QMessageBox(self)
            msgbox.setText("No item is selected, please select an item.")
            msgbox.exec()
            return
        os.system(os.path.dirname(sys.argv[0]) + r"\\Screening_Data\\" + self.listWidget.currentItem().text() + '.txt')

    def clear_button(self):
        self.listWidget.clear()
        self.lineEdit.clear()

    def cancel_button(self):
        self.close()

    def search_button(self):
        self.listWidget.clear()
        search_term = self.lineEdit.text()
        search_dir = os.path.dirname(sys.argv[0]) + r"\\Screening_Data\\"
        # create the screening data files name
        all_files = list(filter(os.path.isfile, glob.glob(search_dir + "*")))
        all_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

        for file in all_files:
            path_datafile = file
            data_file = pd.read_csv(path_datafile, sep="\t")
            if float(search_term) in data_file["Barcode"].values:
                self.listWidget.addItem(file[file.rfind("\\")+1:-4])
        if self.listWidget.count() == 0:
            msgbox = QtWidgets.QMessageBox(self)
            msgbox.setText(f"The Barcode {self.lineEdit.text()} does not exit in our database.")
            msgbox.exec()
            return


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    qt_app = AnalysisSearchWithFunction()
    qt_app.show()
    app.exec_()
