import os
import pandas as pd
import sys
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog, QMessageBox
from PyQt5 import QtCore, QtGui, QtWidgets
from Screening_System_PyQt5 import RGui_Graph_FastData_withFunctions
from PyQt5.uic import loadUi

pd.options.display.max_columns = 999
pd.options.display.max_rows = 999
pd.set_option("display.precision", 6)


class FastDataWithFunction(QDialog):
    def __init__(self):
        super(FastDataWithFunction, self).__init__()
        loadUi("RGui_FastData_select.ui", self)
        self.populate_list_widget()
        self.lineEdit.textChanged.connect(self.search_list)
        self.listWidget.itemClicked.connect(self.item_clicked)
        self.pushButton_3.clicked.connect(self.select_clicked)
        self.pushButton_2.clicked.connect(self.close_clicked)

    def select_clicked(self):
        # the first gui is for statistic table
        while self.listWidget.currentItem() is None:
            msgbox = QtWidgets.QMessageBox(self)
            msgbox.setText("No item is selected, please select an item.")
            msgbox.exec()
            return
        ui = RGui_Graph_FastData_withFunctions.GraphFastDataWithFunction(self.listWidget.currentItem().text())
        ui.show()
        self.close()
        ui.exec_()

    def item_clicked(self):
        self.lineEdit.setText(self.listWidget.currentItem().text())

    def close_clicked(self):
        self.close()

    def populate_list_widget(self):
        self.listWidget.clear()
        dir_path = os.path.dirname(sys.argv[0])
        path_data = dir_path + r"\\Fast_data_collection"

        for r, d, f in os.walk(path_data):
            for file in f:
                if ".txt" in file:
                    self.listWidget.addItem(file)

    def search_list(self):
        search = self.lineEdit.text()
        self.listWidget.clear()
        dir_path = os.path.dirname(sys.argv[0])
        path_data = dir_path + r"\\Fast_data_collection"
        if search == "":
            self.populate_list_widget()
        else:
            for r, d, f in os.walk(path_data):
                for file in f:
                    if search in file:
                        self.listWidget.addItem(file)
        # highlight the item when list size is 1
        if self.listWidget.count() == 1:
            self.listWidget.setCurrentRow(0)
            print(self.listWidget.currentItem().text())



if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    qt_app = FastDataWithFunction()
    qt_app.show()

    def exception_hook(exctype, value, traceback):
        print(exctype, value, traceback)
        sys.excepthook(exctype, value, traceback)
        sys.exit(1)

    sys.excepthook = exception_hook
    app.exec_()