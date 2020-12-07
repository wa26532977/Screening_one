import os
import pyodbc
import datetime
import pandas as pd
import math
import numpy as np
import sys
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog, QMessageBox
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.uic import loadUi
from Screening_System_PyQt5 import client
from PyQt5.QtCore import Qt
from PyQt5.uic import loadUi

pd.options.display.max_columns = 999
pd.options.display.max_rows = 999
pd.set_option("display.precision", 6)


class Report_FrontPage_WithFunction(QDialog):

    def __init__(self):
        super(Report_FrontPage_WithFunction, self).__init__()
        loadUi("RGui_Report_FrontPage.ui", self)
        self.tabWidget.setCurrentIndex(0)

    def getTestNumber4(self, x):
        dir_path = os.path.dirname(sys.argv[0])
        path_datafile = dir_path + r"\\Screening_Data\\" + x
        data_file = pd.read_csv(path_datafile, sep="\t")
        # print(data_file)

        # Load the template info for report
        path_template = dir_path + r"\\Screening_Template\\" + x
        data_file1 = pd.read_csv(path_template, sep="\t")
        columns_1 = data_file1.loc[0].fillna("N/A")
        testNumber = x[:-4]
        # set the Test Number
        self.lineEdit.setText(testNumber)
        # set the Report Date
        self.lineEdit_27.setText(str(datetime.date.today()))
        # set the Cell Name
        self.lineEdit_2.setText(str(columns_1["Cell Name"]))
        # set the Cell Chemistry
        self.lineEdit_3.setText(str(columns_1["Chemistry"]))
        # set the Request number
        self.lineEdit_4.setText(str(columns_1["Request Number"]))
        # set the Request Title
        self.lineEdit_5.setText(str(columns_1["Task Number"]))
        # set the Lot Number
        self.lineEdit_6.setText(str(columns_1["Lot No"]))
        # set the Nom.Capacity(Ah)
        self.lineEdit_7.setText(str(columns_1["Capacity"]))
        # set the Project Engineer
        self.lineEdit_8.setText(str(columns_1["Tech POC."]))
        # set the Purpose
        self.lineEdit_9.setText(str(columns_1["Test Purpose"]))
        # set the pre/post OCV/CCV min Check to see if the cell is being tabbed
        if columns_1["Tabbed?"] == "Not Tabbed":
            print("peter  printed?")
            self.lineEdit_12.setText(str(columns_1["Profile One OCV Min"])+" V")
            self.lineEdit_14.setText(str(columns_1["Profile One CCV Min"]) + " V")
            self.lineEdit_15.setText(str(columns_1["Profile Two OCV Min"]) + " V")
            self.lineEdit_17.setText(str(columns_1["Profile Two CCV Min"]) + " V")
        else:
            self.lineEdit_12.setText(str(columns_1["Pre-Tab OCV"]) + " V")
            self.lineEdit_14.setText("N/A")
            self.lineEdit_15.setText(str(columns_1["Post-Tab OCV"]) + " V")
            self.lineEdit_17.setText(str(columns_1["Post-Tab CCV"]) + " V")

        # Cell Dimension(mm)
        self.lineEdit_10.setText(str(columns_1["Dimension (mm)"]))
        # Date of Manufacture
        self.lineEdit_19.setText(str(columns_1["Mfg Date"]))
        # Date of Received
        self.lineEdit_20.setText(str(columns_1["Date Received"]))
        # Profile one
        if columns_1["Tabbed?"] == "Not Tabbed":
            if columns_1["Profile One Type"] == "Constant Current":
                self.lineEdit_21.setText(
                    str(columns_1["Profile One Values"]) + " mA for " + str((columns_1["Profile One Timer"]))+" Seconds")
            else:
                self.lineEdit_21.setText(
                    str(columns_1["Profile One Values"]) + " mV for " + str((columns_1["Profile One Timer"])) + " Seconds")
        else:
            self.label_24.setText("Pre-Tab")
            self.lineEdit_21.setText("N/A")
        # Profile two
        if columns_1["Tabbed?"] == "Not Tabbed":
            if columns_1["Profile Two Type"] == "Constant Current":
                self.lineEdit_22.setText(
                    str(columns_1["Profile Two Value"]) + " mA for " + str((columns_1["Profile Two Timer"]))+" Seconds")
            else:
                self.lineEdit_22.setText(
                    str(columns_1["Profile Two Value"]) + " mV for " + str((columns_1["Profile Two Timer"])) + " Seconds")
        else:
            self.label_25.setText("Post-Tab")
            if columns_1["Profile One Type"] == "Constant Current":
                self.lineEdit_22.setText(
                    str(columns_1["Profile One Values"]) + " mA for " + str((columns_1["Profile One Timer"]))+" Seconds")
            else:
                self.lineEdit_22.setText(
                    str(columns_1["Profile One Values"]) + " mV for " + str((columns_1["Profile One Timer"])) + " Seconds")

        # set Discharge Temp
        self.lineEdit_23.setText(str(columns_1["Screening Temp(°C)"]))
        # Number of sample
        self.lineEdit_24.setText(str(columns_1["Total Sample Count"]))
        # Date Started
        self.lineEdit_25.setText(str(columns_1["Begin Date"]))
        # Date completed
        self.lineEdit_26.setText(str(columns_1["Finished Date"]))
        # set the Note
        self.plainTextEdit.setPlainText(columns_1["Note"])














