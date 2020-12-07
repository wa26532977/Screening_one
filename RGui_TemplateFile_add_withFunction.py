import os
import pyodbc
import datetime
import pandas as pd
import numpy as np
import sys
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog, QMessageBox
from PyQt5 import QtCore, QtGui, QtWidgets

from PyQt5.uic import loadUi

pd.options.display.max_columns = 999
pd.options.display.max_rows = 999
pd.set_option("display.precision", 6)

class Lift2Coding(QDialog):
    def __init__(self):
        super(Lift2Coding, self).__init__()
        loadUi("RGui_TemplateFile_add.ui", self)
        self.radioButton_25.setChecked(True)
        self.radioButton_27.setChecked(True)
        self.radioButton_4.setChecked(True)
        self.radioButton_12.setChecked(True)
        self.radioButton_4.setChecked(True)
        self.frame_3.setEnabled(False)
        self.tabWidget.setCurrentIndex(0)
        self.pushButton.setAutoDefault(False)
        self.pushButton_2.setAutoDefault(False)
        self.pushButton_3.setAutoDefault(False)
        self.pushButton_4.setAutoDefault(False)
        self.lineEdit_1.returnPressed.connect(self.searchRedely)
        self.radioButton_11.clicked.connect(self.tabbedCellsYes)
        self.radioButton_12.clicked.connect(self.tabbedCellsNo)
        self.radioButton_26.clicked.connect(self.twoProfiles)
        self.radioButton_25.clicked.connect(self.oneProfile)
        self.radioButton_28.clicked.connect(self.twoSection)
        self.comboBox_3.currentIndexChanged.connect(self.ProfileOneTypeCC)
        self.comboBox_4.currentIndexChanged.connect(self.ProfileOneTypeCC1)
        self.radioButton_3.clicked.connect(self.HighRateYes)
        self.radioButton_4.clicked.connect(self.HighRateNo)
        self.pushButton_2.clicked.connect(self.SaveButton)
        self.pushButton_4.clicked.connect(self.SaveButton)
        self.pushButton.clicked.connect(self.CancelButton)
        self.pushButton_3.clicked.connect(self.CancelButton)

    def CancelButton(self):
        self.close()

    def duplicateNumber(self, x):
        currentName = x
        self.label_18.setText("Duplicating Template")
        self.getTestNumber(currentName)

    def getTestNumber(self, x):
        #setting all the template data to GUI
        #dir_path = os.path.dirname(os.path.realpath(__file__))
        dir_path = os.path.dirname(sys.argv[0])
        path_data = dir_path + r"\\Screening_Template\\"+x

        data_file = pd.read_csv(path_data, sep="\t")
        print(data_file)

        columns_1 = data_file.loc[0].fillna("")

        print(columns_1[1])
        for i in range(1, 32):
            y = str(columns_1[i-1])
            x = "self.lineEdit_" + str(i) + ".setText(\'"+y+"\')"
            exec(x)

        if columns_1[31] == "Cell":
            print("Cell")
            self.radioButton.setChecked(True)
        else:
            print("Battery")
            self.radioButton_2.setChecked(True)

        if columns_1[32] == "Not Tabbed":
            print("not Tabbed")
            self.radioButton_12.setChecked(True)
        else:
            print("Tabbed")
            self.frame_7.setEnabled(True)
            self.radioButton_11.setChecked(True)

        if columns_1[33] == 1:
            print("p1")
            self.radioButton_25.setChecked(True)
        else:
            print("p2")
            self.frame_9.setEnabled(True)
            self.radioButton_26.setChecked(True)

        if columns_1[34] == 1:
            print("S1")
            self.radioButton_27.setChecked(True)
        else:
            print("s2")
            self.radioButton_28.setChecked(True)

        if columns_1[35] == "No":
            print("No High Rate")
            self.radioButton_4.setChecked(True)
        else:
            print("High Rate")
            self.frame_3.setEnabled(True)
            self.radioButton_3.setChecked(True)

        if columns_1[36] == "Constant Current":
            print("Constant Current")
            self.comboBox_3.setCurrentIndex(0)
            self.label_27.setText("mA")
        else:
            print("Constant Resistor")
            self.comboBox_3.setCurrentIndex(1)
            self.label_27.setText("Ω")

        if columns_1[37] == "Constant Current":
            print("Constant Current")
            self.comboBox_4.setCurrentIndex(0)
            self.label_29.setText("mA")
        else:
            print("Constant Resistor")
            self.comboBox_4.setCurrentIndex(1)
            self.label_29.setText("Ω")

        self.plainTextEdit.setPlainText(columns_1[38])

    def SaveButton(self):
        #checking it has a Test no.
        if str(self.lineEdit_2.text()) == "":
            msgbox = QtWidgets.QMessageBox(self)
            msgbox.setText("A Test No. is Required")
            msgbox.exec()
            return
        # for duplicateing template the name need to be a new one
        if str(self.label_18.text()) == "Duplicating Template":
            #dir_path = os.path.dirname(os.path.realpath(__file__))
            dir_path = os.path.dirname(sys.argv[0])
            path_template = dir_path + r"\\Screening_Template\\" + str(self.lineEdit_2.text()) + ".txt"
            if os.path.exists(path_template):
                msgbox = QtWidgets.QMessageBox(self)
                msgbox.setText("Please change this duplicated template Name")
                msgbox.exec()
                return

        columns_1 = []
        for i in range(1, 32):
            x = "self.lineEdit_" + str(i) + ".text()"
            columns_1.append(eval(x))
        #checking all criteria is all number
        res = []
        for val in columns_1[16:31]:
            if val != "":
                try:
                    res.append(float(val))
                except ValueError:
                    msgbox = QtWidgets.QMessageBox(self)
                    msgbox.setText("Please check all your Criteria, make sure they are all numbers")
                    msgbox.exec()
                    return
        #battery or Cell
        if self.radioButton.isChecked():
            columns_1.append("Battery")
        else:
            columns_1.append("Cell")
        #Bing Tabbed?
        if self.radioButton_11.isChecked():
            columns_1.append("Tabbed")
        else:
            columns_1.append("Not Tabbed")
        #Profile Number
        if self.radioButton_25.isChecked():
            columns_1.append(1)
        else:
            columns_1.append(2)
        # Section number
        if self.radioButton_27.isChecked():
            columns_1.append(1)
        else:
            columns_1.append(2)
        # HighRate or not
        if self.radioButton_3.isChecked():
            columns_1.append("Yes")
        else:
            columns_1.append("No")
        #Profileone type
        if self.comboBox_3.currentIndex() == 0:
            columns_1.append("Constant Current")
        else:
            columns_1.append("Constant Resistor")
        # Profiletwo type
        if self.comboBox_4.currentIndex() == 0:
            columns_1.append("Constant Current")
        else:
            columns_1.append("Constant Resistor")
        note = self.plainTextEdit.toPlainText()
        columns_1.append(note)
        #Created the datafram
        df = pd.DataFrame(columns=["Lot No", "Test No", "Cell Name", "Chemistry", "Form Factor", "Capacity",
                                   "Request Number", "Task Number", "Tech POC.", "Test Purpose", "Mfg Date",
                                   "Date Received", "Dimension (mm)", "Screening Temp(°C)", "Begin Date",
                                   "Finished Date", "Total Sample Count", "Profile One Values", "Profile One Timer",
                                   "Profile One OCV Min", "Profile One CCV Min", "High Rate Min(Y-axis)",
                                   "High Rate Max (Y-axis)", "Pre-Tab OCV", "Post-Tab OCV", "Post-Tab CCV",
                                   "OCV Tab Tolerance", "Profile Two Value", "Profile Two Timer", "Profile Two OCV Min",
                                   "Profile Two CCV Min", "Battery/Cell", "Tabbed?", "Profile No.", "Section No.",
                                   "High Rate?", "Profile One Type", "Profile Two Type", "Note"])


        df.loc[0] = columns_1
        #dir_path = os.path.dirname(os.path.realpath(__file__))
        dir_path = os.path.dirname(sys.argv[0])
        path_template = dir_path+r"\\Screening_Template\\"+str(self.lineEdit_2.text())+".txt"
        df.to_csv(path_template, sep="\t", index=False)

        TestingRecoder = pd.read_csv(dir_path+r"\\TestingRecoder.txt", sep="\t")
        print(TestingRecoder["Lot No"].last_valid_index())
        TestingRecoder.loc[TestingRecoder["Lot No"].last_valid_index()+1] = columns_1
        TestingRecoder.to_csv(dir_path+r"\\TestingRecoder.txt", sep="\t", index=False)
        note = self.plainTextEdit.toPlainText()
        self.close()

    def HighRateYes(self):
        self.frame_3.setEnabled(True)

    def HighRateNo(self):
        self.frame_3.setEnabled(False)

    def ProfileOneTypeCC1(self):
        if self.comboBox_4.currentIndex() == 0:
            self.label_29.setText("mA")

        if self.comboBox_4.currentIndex() == 1:
            self.label_29.setText("Ω")

    def ProfileOneTypeCC(self):
        if self.comboBox_3.currentIndex() == 0:
            self.label_27.setText("mA")

        if self.comboBox_3.currentIndex() == 1:
            self.label_27.setText("Ω")

    def twoSection(self):
        self.radioButton_25.setChecked(True)
        self.radioButton_12.setChecked(True)
        self.frame_9.setEnabled(False)
        self.frame_7.setEnabled(False)
        self.lineEdit_20.setEnabled(True)
        self.lineEdit_21.setEnabled(True)

    def oneProfile(self):
        self.frame_9.setEnabled(False)
        self.lineEdit_20.setEnabled(True)
        self.lineEdit_21.setEnabled(True)

    def twoProfiles(self):
        self.frame_9.setEnabled(True)
        self.frame_7.setEnabled(False)
        self.radioButton_12.setChecked(True)
        self.radioButton_27.setChecked(True)
        self.lineEdit_20.setEnabled(True)
        self.lineEdit_21.setEnabled(True)

    def tabbedCellsYes(self):
        self.frame_7.setEnabled(True)
        self.radioButton_25.setChecked(True)
        self.radioButton_27.setChecked(True)
        self.frame_9.setEnabled(False)
        self.lineEdit_20.setEnabled(False)
        self.lineEdit_21.setEnabled(False)

    def tabbedCellsNo(self):
        self.frame_7.setEnabled(False)
        self.lineEdit_20.setEnabled(True)
        self.lineEdit_21.setEnabled(True)

    def searchRedely(self):
        constr = 'DRIVER={SQL Server};SERVER=BTC-SQL2016\\BTCSQL2016;DATABASE=RADLEY_PROD;Trusted_Connection=Yes'
        con = pyodbc.connect(constr)
        print(con)
        query = "SELECT XLSM_LOT as [Lot], " + \
                "XIM_ITEM as [Item], " + \
                "XEF_USER_CHAR1 as [Manufacturer], " + \
                "XIC_NAME as [Chemistry], " + \
                "XPT_NAME as [Form Factor], " + \
                "XIT_NAME as [Type], " + \
                "XIM_GRADE as [Nominal Capacity], " + \
                "XIM_UPC_CODE as [Nominal Voltage], " + \
                "XIM_HEIGHT as [Height], " + \
                "XIM_WIDTH as [Width], " + \
                "XIM_DEPTH as [Depth], " + \
                "XIM_WEIGHT as [Weight], " + \
                "XIM_DIAMETER as [Diameter], " + \
                "XLSM_RECEIPT_DATE as [Receipt Date], " + \
                "XLSM_MFG_DATE as [Mfg Date] " + \
                "FROM XINV_LOT_SERIAL_MASTER " + \
                "JOIN XINV_ITEM_MASTER ON XLSM_ITEM_ID = XIM_ID " + \
                "JOIN XINV_ITEM_CLASS ON XIM_CLASS_ID = XIC_ID " + \
                "JOIN XAS_EXTENDED_FIELD ON XIM_EXTENDED_FIELD_ID = XEF_ID " + \
                "LEFT JOIN XINV_PACKAGE_TYPE ON XIM_PACKAGE_TYPE_ID = XPT_ID " + \
                "LEFT JOIN XINV_ITEM_TYPE ON XIM_TYPE_ID = XIT_ID "
        query += "WHERE XLSM_LOT = '" + str(self.lineEdit_1.text()) + "' "
        print(str(self.lineEdit_1.text()))
        df = pd.read_sql_query(query, con)
        print("DF = ")
        print(df)
        if df.empty:
            msgbox = QtWidgets.QMessageBox(self)
            msgbox.setText("Can't find this lot number, please check.")
            msgbox.exec()
            return
        self.lineEdit_3.setText(str(df.iloc[0]["Manufacturer"]) + " " + str(df.iloc[0]["Item"]))
        self.lineEdit_4.setText(str(df.iloc[0]["Chemistry"]))
        self.lineEdit_5.setText(str(df.iloc[0]["Form Factor"]) + " " + str(df.iloc[0]["Type"]))
        self.lineEdit_6.setText(str(df.iloc[0]["Nominal Capacity"]))
        self.lineEdit_11.setText(str(df.iloc[0]["Mfg Date"])[0:10])
        self.lineEdit_12.setText(str(df.iloc[0]["Receipt Date"])[0:10])
        if df.iloc[0]["Diameter"] is not None:
            self.lineEdit_13.setText(str(df.iloc[0]["Height"]) + "(H) X " + str(df.iloc[0]["Diameter"]) + "(D)")
        elif df.iloc[0]["Width"] is not None:
            self.lineEdit_13.setText(str(df.iloc[0]["Height"]) + "(H) X" + str(df.iloc[0]["Width"]) + "(W) X" + str(
                df.iloc[0]["Depth"]) + "(D)")
        self.lineEdit_15.setText(str(datetime.date.today()))
        self.lineEdit_16.setText(str(datetime.date.today()))
        self.lineEdit_14.setText("24")