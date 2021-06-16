import os
import pyodbc
import datetime
import pandas as pd
import numpy as np
import sys
from PyQt5.QtWidgets import QDialog, QTableWidgetItem
from PyQt5 import QtWidgets
from PyQt5.uic import loadUi
from scipy import stats
from docx import Document
import docx

pd.options.display.max_columns = 999
pd.options.display.max_rows = 999
pd.set_option("display.precision", 6)


class Report_RawData_WithFunction(QDialog):
    def __init__(self):
        super(Report_RawData_WithFunction, self).__init__()
        loadUi("RGui_Report_RawData.ui", self)

    def findSampleOutlier(self, data):
        iqr = stats.iqr(data)
        q1, q3 = np.percentile(data, [25, 75])

        min = round(q1 - (1.5 * iqr), 3)
        max = round(q3 + (1.5 * iqr), 3)

        #outlier = len(data) - len(data[data.between(min, max)])
        return min, max

    def outlier_report_generator(self, x):
        path_data = os.path.dirname(sys.argv[0]) + r"\\Report_word_Outlier\\" + x[0: -4] + r"outlier.txt"
        df = pd.DataFrame(columns=["Barcode", "Pre-OCV", "Pre-CCV", "Post-OCV", "Post-CCV", "Outlier"])
        df.to_csv(path_data, sep="\t", index=False)
        i = 0
        while i < self.tableWidget.rowCount():
            df.loc[i, "Barcode"] = self.tableWidget.item(i, 0).text()
            df.loc[i, "Pre-OCV"] = self.tableWidget.item(i, 3).text()
            if self.tableWidget.item(i, 4) is not None:
                df.loc[i, "Pre-CCV"] = self.tableWidget.item(i, 4).text()
            if self.tableWidget.item(i, 5) is not None:
                df.loc[i, "Post-OCV"] = self.tableWidget.item(i, 5).text()
            if self.tableWidget.item(i, 6) is not None:
                df.loc[i, "Post-CCV"] = self.tableWidget.item(i, 6).text()
            df.loc[i, "Outlier"] = self.tableWidget.item(i, 7).text()
            i += 1
        df.to_csv(path_data, sep="\t", index=False)

    def getTestNumber3(self, x):
        # Remove the .txt from x
        testNumber = x[:-4]
        self.label_4.setText(testNumber)

        # Load the template for criteria
        path_template = os.path.dirname(sys.argv[0]) + r"\\Screening_Template\\" + x
        data_file1 = pd.read_csv(path_template, sep="\t")
        columns_1 = data_file1.loc[0].fillna("")

        # get the MFR date
        MFR_Date = columns_1["Mfg Date"]
        # set the criteria
        # Get the Load info for Pre-screening
        if columns_1["Tabbed?"] == "Not Tabbed":
            if columns_1[36] == "Constant Resistor":
                loadinfo = str(columns_1[17]) + " Ohms for " + str(columns_1[18]) + " Seconds."
                self.label_5.setText(loadinfo)
            elif columns_1[36] == "Constant Current":
                loadinfo = str(columns_1[17]) + " mA for " + str(columns_1[18]) + " Seconds."
                self.label_5.setText(loadinfo)
        else:
            self.label_5.setText("Pre-Tab")

        # Get the load info for Post-Screening
        # Change to Post Table
        if columns_1["Tabbed?"] == "Not Tabbed":
            if not columns_1[27] == "":
                if columns_1[37] == "Constant Resistor":
                    loadinfo = str(columns_1[27]) + " Ohms for " + str(columns_1[28]) + " Seconds."
                    self.label_6.setText(loadinfo)
                elif columns_1[37] == "Constant Current":
                    loadinfo = str(columns_1[27]) + " mA for " + str(columns_1[28]) + " Seconds."
                    self.label_6.setText(loadinfo)
        else:
            if columns_1[36] == "Constant Resistor":
                loadinfo = str(columns_1[17]) + " Ohms for " + str(columns_1[18]) + " Seconds."
                self.label_6.setText(loadinfo)
            elif columns_1[36] == "Constant Current":
                loadinfo = str(columns_1[17]) + " mA for " + str(columns_1[18]) + " Seconds."
                self.label_6.setText(loadinfo)

        # load the data file

        path_datafile = os.path.dirname(sys.argv[0]) + r"\\Screening_Data\\" + x
        data_file = pd.read_csv(path_datafile, sep="\t")
        # set the rowcount for the tablewidget
        self.tableWidget.setRowCount(len(data_file))
        i = 0
        fall_cell = 0
        inspection = ""

        while i < len(data_file):
            # Inspection
            if pd.isna(data_file.iloc[i]["Comments"]) is False:
                # self.tableWidget.setItem(i, 7, QTableWidgetItem(str(data_file.iloc[i]["Comments"])))
                self.tableWidget.setItem(i, 7, QTableWidgetItem("OK"))
            else:
                self.tableWidget.setItem(i, 7, QTableWidgetItem("OK"))

            # barcode
            self.tableWidget.setItem(i, 0, QTableWidgetItem(str(data_file.iloc[i]["Barcode"])))
            # MFR. SN
            if pd.isna(data_file.iloc[i]["Serial#"]) is False:
                self.tableWidget.setItem(i, 1, QTableWidgetItem(str(data_file.iloc[i]["Serial#"])))
            else:
                self.tableWidget.setItem(i, 1, QTableWidgetItem(str("")))
            # MFR. SN
            if MFR_Date != "None":
                self.tableWidget.setItem(i, 2, QTableWidgetItem(str(MFR_Date)))
            # Pre-OCV
            if not pd.isna(data_file.iloc[i]["Pre-OCV"]):
                dropNan = data_file["Pre-OCV"].dropna()
                checkOutlier = self.findSampleOutlier(dropNan)

                # check if tab for different criteria
                if columns_1["Tabbed?"] == "Tabbed":
                    criteria = columns_1["Pre-Tab OCV"]
                else:
                    criteria = columns_1["Profile One OCV Min"]

                # check if the data failed the criteria
                if data_file.iloc[i]["Pre-OCV"] < criteria:
                    self.tableWidget.setItem(i, 3, QTableWidgetItem("%.3f" % (round(data_file.iloc[i]["Pre-OCV"], 3)) + " !"))
                    inspection += " F"
                # check the min of the outlier
                elif data_file.iloc[i]["Pre-OCV"] < checkOutlier[0]:
                    self.tableWidget.setItem(i, 3, QTableWidgetItem("%.3f" % (round(data_file.iloc[i]["Pre-OCV"], 3))+" *"))
                    inspection += " OL"
                elif data_file.iloc[i]["Pre-OCV"] > checkOutlier[1]:
                    self.tableWidget.setItem(i, 3, QTableWidgetItem("%.3f" % (round(data_file.iloc[i]["Pre-OCV"], 3)) + " *"))
                    inspection += " OH"
                else:
                    self.tableWidget.setItem(i, 3, QTableWidgetItem("%.3f" % (round(data_file.iloc[i]["Pre-OCV"], 3))))
            # Pre-CCV
            # Check if the data failed the criteria
            if not pd.isna(data_file.iloc[i]["Pre-CCV"]):
                dropNan = data_file["Pre-CCV"].dropna()
                checkOutlier = self.findSampleOutlier(dropNan)
                # Check to see if Pre-CCV fail
                if data_file.iloc[i]["Pre-CCV"] < columns_1["Profile One CCV Min"]:
                    self.tableWidget.setItem(i, 4,
                                             QTableWidgetItem("%.3f" % (round(data_file.iloc[i]["Pre-CCV"], 3)) + " !"))
                    inspection += " F"
                # check the min of the outliter
                elif data_file.iloc[i]["Pre-CCV"] < checkOutlier[0]:
                    self.tableWidget.setItem(i, 4, QTableWidgetItem("%.3f" % (round(data_file.iloc[i]["Pre-CCV"], 3))+" *"))
                    inspection += " OL"
                elif data_file.iloc[i]["Pre-CCV"] > checkOutlier[1]:
                    self.tableWidget.setItem(i, 4, QTableWidgetItem("%.3f" % (round(data_file.iloc[i]["Pre-CCV"], 3)) + " *"))
                    inspection += " OH"
                else:
                    self.tableWidget.setItem(i, 4, QTableWidgetItem("%.3f" % (round(data_file.iloc[i]["Pre-CCV"], 3))))

            # Post-OCV
            if not pd.isna(data_file.iloc[i]["Post-OCV"]):
                dropNan = data_file["Post-OCV"].dropna()
                checkOutlier = self.findSampleOutlier(dropNan)
                # check the min of the outliter

                # check if tab for different criteria
                if columns_1["Tabbed?"] == "Tabbed":
                    criteria = columns_1["Post-Tab OCV"]
                elif columns_1["Section No."] == 2:
                    criteria = columns_1["Profile One OCV Min"]
                else:
                    criteria = columns_1["Profile Two OCV Min"]

                # check if the data failed the criteria
                if columns_1["Tabbed?"] == "Tabbed" and \
                        round(data_file.iloc[i]["Pre-OCV"], 3) - round(data_file.iloc[i]["Post-OCV"], 3) > (columns_1["OCV Tab Tolerance"] + 0.0001):
                    self.tableWidget.setItem(i, 5,
                                             QTableWidgetItem("%.3f" % (round(data_file.iloc[i]["Post-OCV"], 3)) + " ^"))
                    inspection = inspection + ' T'
                elif data_file.iloc[i]["Post-OCV"] < criteria:
                    self.tableWidget.setItem(i, 5, QTableWidgetItem("%.3f" % (round(data_file.iloc[i]["Post-OCV"], 3)) + " !"))
                    inspection = inspection + ' F'
                elif data_file.iloc[i]["Post-OCV"] < checkOutlier[0]:
                    self.tableWidget.setItem(i, 5, QTableWidgetItem("%.3f" % (round(data_file.iloc[i]["Post-OCV"], 3))+" *"))
                    inspection = inspection + ' OL'
                elif data_file.iloc[i]["Post-OCV"] > checkOutlier[1]:
                    self.tableWidget.setItem(i, 5, QTableWidgetItem("%.3f" % (round(data_file.iloc[i]["Post-OCV"], 3)) + " *"))
                    inspection = inspection + ' OH'
                else:
                    self.tableWidget.setItem(i, 5, QTableWidgetItem("%.3f" % (round(data_file.iloc[i]["Post-OCV"], 3))))
            # Post-CCV
            if not pd.isna(data_file.iloc[i]["Post-CCV"]):
                dropNan = data_file["Post-CCV"].dropna()
                checkOutlier = self.findSampleOutlier(dropNan)
                # check the min of the outliter

                # check if tab for different criteria
                if columns_1["Tabbed?"] == "Tabbed":
                    criteria = columns_1["Post-Tab CCV"]
                elif columns_1["Section No."] == 2:
                    criteria = columns_1["Profile One CCV Min"]
                else:
                    criteria = columns_1["Profile Two CCV Min"]

                # check if the data failed the criteria
                if data_file.iloc[i]["Post-CCV"] < criteria:
                    self.tableWidget.setItem(i, 6,
                                             QTableWidgetItem("%.3f" % (round(data_file.iloc[i]["Post-CCV"], 3)) + " !"))
                    inspection += ' F'
                elif data_file.iloc[i]["Post-CCV"] < checkOutlier[0]:
                    self.tableWidget.setItem(i, 6, QTableWidgetItem("%.3f" % (round(data_file.iloc[i]["Post-CCV"], 3))+" *"))
                    inspection += ' OL'
                elif data_file.iloc[i]["Post-CCV"] > checkOutlier[1]:
                    self.tableWidget.setItem(i, 6, QTableWidgetItem("%.3f" % (round(data_file.iloc[i]["Post-CCV"], 3)) + " *"))
                    inspection += ' OH'
                else:
                    self.tableWidget.setItem(i, 6, QTableWidgetItem("%.3f" % (round(data_file.iloc[i]["Post-CCV"], 3))))
                if inspection == "":
                    inspection = "OK"
                self.tableWidget.setItem(i, 7, QTableWidgetItem(inspection.lstrip()))
                inspection = ""
            # getting the failed count number
            #if data_file.iloc[i]["Pre-screen pass"] == "Pass" and data_file.iloc[i]["Post-screen pass"] == "Pass":
            if "T" in self.tableWidget.item(i, 7).text() or "F" in self.tableWidget.item(i, 7).text():
                fall_cell = fall_cell + 1
            i = i + 1

        # total number of sample tested
        self.label_10.setText(str(len(data_file)))
        # total number of sample failed
        self.label_8.setText(str(fall_cell))
        # set the column header
        if columns_1["Tabbed?"] == "Tabbed":
            self.tableWidget.setHorizontalHeaderLabels(
                ["Barcode", "MFR.SN", "MFR.Date", "Pre-Tab OCV", "", "Post-Tab OCV", "Post-Tab CCV", "Inspection"])
            self.groupBox.setTitle("Pre-Tab")
            self.groupBox_2.setTitle("Post-Tab")
        elif columns_1["Profile No."] == 2:
            self.tableWidget.setHorizontalHeaderLabels(
                ["Barcode", "MFR.SN", "MFR.Date", "Profile One OCV", "Profile One CCV", "Profile Two OCV", "Profile Two CCV", "Inspection"])
            self.groupBox.setTitle("Profile one")
            self.groupBox_2.setTitle("Profile Two")
        elif columns_1["Section No."] == 2:
            self.tableWidget.setHorizontalHeaderLabels(
                ["Barcode", "MFR.SN", "MFR.Date", "Section One OCV", "Section One CCV", "Section Two OCV", "Section Two CCV", "Inspection"])
            self.groupBox.setTitle("Section one")
            self.groupBox_2.setTitle("Section Two")
            if columns_1[36] == "Constant Resistor":
                loadinfo = str(columns_1["Profile One Values"]) + " Ohms for " + str(
                    columns_1["Profile One Timer"]) + " Seconds."
                self.label_6.setText(loadinfo)
            elif columns_1[36] == "Constant Current":
                loadinfo = str(columns_1["Profile One Values"]) + " mA for " + str(
                    columns_1["Profile One Timer"]) + " Seconds."
                self.label_6.setText(loadinfo)
        else:
            self.groupBox.setTitle("Profile")
            self.groupBox_2.setTitle("")
            self.tableWidget.setHorizontalHeaderLabels(
                ["Barcode", "MFR.SN", "MFR.Date", "OCV", "CCV", "", "", "Inspection"])
        # creat the outlier.txt
        self.outlier_report_generator(x)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    qt_app = Report_RawData_WithFunction()
    qt_app.getTestNumber3("14575A00.txt")
    qt_app.show()
    app.exec_()
