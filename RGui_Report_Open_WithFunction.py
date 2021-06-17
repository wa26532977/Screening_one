import os
import pandas as pd
import sys
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog, QMessageBox
from PyQt5 import QtCore, QtGui, QtWidgets
from Screening_System_PyQt5 import RGui_Report_StatisticsTable_WithFunctions, RGui_Report_Open_WithFunction_2, \
    RGui_Report_RawData_WithFunction, RGui_Report_FrontPage_WithFunction, RGui_Graph_Pre_WithFunction, \
    RGui_Graph_FastData_withFunctions
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
        self.show_report = False

    def item_Clicked(self):
        search = self.listWidget.currentItem().text()
        self.lineEdit.setText(search)

    def show_report_toggle(self):
        self.show_report = True

    def OK_Pressed(self):
        # the first gui is for statistic table
        ui = RGui_Report_StatisticsTable_WithFunctions.Report_StatisticsTable_WithFunctions()

        while self.listWidget.currentItem() is None:
            msgbox = QtWidgets.QMessageBox(self)
            msgbox.setText("No item is selected, please select an item.")
            msgbox.exec()
            return
        ui.getTestNumber2(self.listWidget.currentItem().text())
        print(f"pete looking {self.listWidget.currentItem().text()}")
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
        # show the fast data collection
        dir_path = os.path.dirname(sys.argv[0])
        fast_data_dir = dir_path + r"\Fast_data_collection\\" + self.listWidget.currentItem().text()

        ui.show()
        ui2.show()
        ui3.show()
        ui4.show()

        # if fast_data exit show it
        if os.path.exists(fast_data_dir):
            ui6 = RGui_Graph_FastData_withFunctions.GraphFastDataWithFunction(self.listWidget.currentItem().text())
            ui6.show()

        # The PostGraph of the Report
        # first check if there is any postData
        path_datafile = os.path.dirname(sys.argv[0]) + r"\\Screening_Data\\" + self.listWidget.currentItem().text()
        data_file = pd.read_csv(path_datafile, sep="\t")
        if not data_file["Post-OCV"].isnull().all():
            ui5 = RGui_Graph_Pre_WithFunction.Graph_Pre_WithFunction()
            ui5.getTestNumber5(self.listWidget.currentItem().text(), "Post")
            ui5.show()

        if self.show_report:
            ui7 = RGui_Report_Open_WithFunction_2.Report_Open_WithFunctions()
            ui7.OK_Pressed(self.listWidget.currentItem().text())

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
