import os
import pyodbc
import datetime
import pandas as pd
import numpy as np
import sys
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog, QMessageBox
from PyQt5 import QtCore, QtGui, QtWidgets
from Screening_System_PyQt5 import RGui_Report_StatisticsTable_WithFunctions
from Screening_System_PyQt5 import RGui_Report_RawData_WithFunction
from Screening_System_PyQt5 import RGui_Report_FrontPage_WithFunction
from Screening_System_PyQt5 import RGui_Graph_Pre_WithFunction
from PyQt5.uic import loadUi

pd.options.display.max_columns = 999
pd.options.display.max_rows = 999
pd.set_option("display.precision", 6)


class Report_Open_WithFunctions(QDialog):
    def __init__(self):
        super(Report_Open_WithFunctions, self).__init__()
        loadUi("RGui_Report_Open.ui", self)
        self.populate_list_widget()
        self.lineEdit.textChanged.connect(self.search_In_List)
        self.listWidget.itemClicked.connect(self.item_Clicked)
        self.pushButton_2.clicked.connect(self.OK_Pressed)

    def item_Clicked(self):
        search = self.listWidget.currentItem().text()
        self.lineEdit.setText(search)


    def OK_Pressed(self):
        # the first gui is for statistic table
        ui = RGui_Report_StatisticsTable_WithFunctions.Report_StatisticsTable_WithFunctions()
        while self.listWidget.currentItem() is None:
            msgbox = QtWidgets.QMessageBox(self)
            msgbox.setText("No item is selected, please select an item.")
            msgbox.exec()
            return
        ui.getTestNumber2(self.listWidget.currentItem().text())
        self.close()

        # the second GUI is for RawData Table
        ui2 = RGui_Report_RawData_WithFunction.Report_RawData_WithFunction()
        ui2.getTestNumber3(self.listWidget.currentItem().text())
        # The frontPage of the report
        ui3 = RGui_Report_FrontPage_WithFunction.Report_FrontPage_WithFunction()
        ui3.getTestNumber4(self.listWidget.currentItem().text())
        # The PreGraph of the report
        ui4 = RGui_Graph_Pre_WithFunction.Graph_Pre_WithFunction()
        ui4.getTestNumber5(self.listWidget.currentItem().text(), "Pre")

        ui.show()
        ui2.show()
        ui3.show()
        ui4.show()

        # The PostGraph of the Report
        # first check if there is any postData
        path_datafile = os.getcwd() + r"\\Screening_Data\\" + self.listWidget.currentItem().text()
        data_file = pd.read_csv(path_datafile, sep="\t")
        if not data_file["Post-OCV"].isnull().all():
            ui5 = RGui_Graph_Pre_WithFunction.Graph_Pre_WithFunction()
            ui5.getTestNumber5(self.listWidget.currentItem().text(), "Post")
            ui5.show()

        ui.exec_()




    def search_In_List(self):
        search = self.lineEdit.text()
        self.listWidget.clear()
        #dir_path = os.path.dirname(os.path.realpath(__file__))
        dir_path = os.path.dirname(sys.argv[0])
        path_data = dir_path + r"\\Screening_Data"
        if search == "":
            self.populate_list_widget()
        else:
            for r, d, f in os.walk(path_data):
                for file in f:
                    if search in file:
                        self.listWidget.addItem(file)
        # hightlight the item when list size is 1
        if self.listWidget.count() == 1:
            self.listWidget.setCurrentRow(0)
            print(self.listWidget.currentItem().text())

    def populate_list_widget(self):
        self.listWidget.clear()
        #dir_path = os.path.dirname(os.path.realpath(__file__))
        dir_path = os.path.dirname(sys.argv[0])
        path_data = dir_path + r"\\Screening_Data"

        for r, d, f in os.walk(path_data):
            for file in f:
                if ".txt" in file:
                    self.listWidget.addItem(file)
