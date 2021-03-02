import os
import pyodbc
import datetime
import pandas as pd
import math
import numpy as np
import sys
from PyQt5.QtGui import *
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog, QMessageBox, QTableWidgetItem
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.uic import loadUi
from scipy import stats
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker


pd.options.display.max_columns = 999
pd.options.display.max_rows = 999
pd.set_option("display.precision", 6)


class Graph_Pre_WithFunction(QDialog):
    x = "something"
    PrePost = ""

    def __init__(self):
        super(Graph_Pre_WithFunction, self).__init__()
        loadUi("RGui_Graph_Pre.ui", self)
        self.pushButton_16.clicked.connect(self.WithoutCriteriaOCV)
        self.pushButton_17.clicked.connect(self.WithCriteriaOCV)
        self.pushButton_18.clicked.connect(self.WithoutCriteriaCCV)
        self.pushButton_19.clicked.connect(self.WithCriteriaCCV)
        self.pushButton_20.clicked.connect(self.WithoutCriteriaOCV_CCV)
        self.pushButton_21.clicked.connect(self.WithCriteriaOCV_CCV)
        self.pushButton_12.clicked.connect(self.OCVConfidence)
        self.pushButton_15.clicked.connect(self.CCVConfidence)

    def mean_confidence_interval(self, data, confidence=0.95):
        a = 1.0 * np.array(data)
        n = len(a)
        m, se = np.mean(a), stats.sem(a)
        h = se * stats.t.ppf((1 + confidence) / 2., n - 1)
        # m is mean, m-h is min, m+h is max
        return m, m - h, m + h

    def CCVConfidence(self):
        if self.PrePost == "Pre":
            pixmap = QPixmap("Pre CCV 95% confidence interval.png")
        if self.PrePost == "Post":
            pixmap = QPixmap("Post CCV 95% confidence interval.png")
        self.label_12.setPixmap(pixmap)

    def OCVConfidence(self):
        if self.PrePost == "Pre":
            pixmap = QPixmap("Pre OCV 95% confidence interval.png")
        if self.PrePost == "Post":
            pixmap = QPixmap("Post OCV 95% confidence interval.png")
        self.label_12.setPixmap(pixmap)

    def WithoutCriteriaOCV_CCV(self):
        if self.PrePost == "Pre":
            pixmap = QPixmap("OCV-CCV.png")
        if self.PrePost == "Post":
            pixmap = QPixmap("OCV-CCV 2.png")
        self.label_12.setPixmap(pixmap)

    def WithCriteriaOCV_CCV(self):
        if self.PrePost == "Pre":
            pixmap = QPixmap("Pre OCV-CCV WITHIN SD RANGE.png")
        if self.PrePost == "Post":
            pixmap = QPixmap("Post OCV-CCV WITHIN SD RANGE.png")
        self.label_12.setPixmap(pixmap)

    def WithCriteriaOCV(self):
        if self.PrePost == "Pre":
            pixmap = QPixmap("Pre OCV WITHIN SD RANGE.png")
        if self.PrePost == "Post":
            pixmap = QPixmap("Post OCV WITHIN SD RANGE.png")
        self.label_12.setPixmap(pixmap)

    def WithoutCriteriaOCV(self):
        if self.PrePost == "Pre":
            pixmap = QPixmap("OCV.png")
        if self.PrePost == "Post":
            pixmap = QPixmap("OCV 2.png")
        self.label_12.setPixmap(pixmap)

    def WithCriteriaCCV(self):
        if self.PrePost == "Pre":
            pixmap = QPixmap("Pre CCV WITHIN SD RANGE.png")
        if self.PrePost == "Post":
            pixmap = QPixmap("Post CCV WITHIN SD RANGE.png")
        self.label_12.setPixmap(pixmap)

    def WithoutCriteriaCCV(self):
        if self.PrePost == "Pre":
            pixmap = QPixmap("CCV.png")
        if self.PrePost == "Post":
            pixmap = QPixmap("CCV 2.png")
        self.label_12.setPixmap(pixmap)

    def findingCriteraforGraph(self, data, columns_1, type):
        # this function is to find the current data for building graph and for different test case which kind
        # crtiera for use to filter the data
        if "WITHIN" in type:
            if columns_1["Tabbed?"] == "Tabbed" and "Post" in type:
                if "OCV-CCV" in type:
                    status = "CCV > or = " + str(columns_1["Post-Tab CCV"]) + " V"
                    OCV_CCV = data[data["Post-OCV"] >= columns_1["Post-Tab OCV"]]["Post-OCV"] - data[data["Post-CCV"] >= columns_1["Post-Tab CCV"]]["Post-CCV"]
                    graphData = round(OCV_CCV.dropna(), 3)
                    print("Tabbed and post OCV-CCV")
                elif "OCV" in type:
                    status = "OCV > or = " + str(columns_1["Post-Tab OCV"]) + " V"
                    graphData = round(data[data >= columns_1["Post-Tab OCV"]], 3)
                    self.x = columns_1["Post-Tab OCV"]
                    print("Tabbed and post OCV")
                elif "CCV" in type:
                    status = "CCV > or = " + str(columns_1["Post-Tab CCV"]) + " V"
                    graphData = round(data[data >= columns_1["Post-Tab CCV"]], 3)
                    self.x = columns_1["Post-Tab CCV"]
                    print("Tabbed and post CCV")

            elif columns_1["Tabbed?"] == "Tabbed" and "Pre" in type:
                if "OCV-CCV" in type:
                    print("Tabbed and pre OCV-CCV: None")
                    return
                elif "OCV" in type:
                    status = "OCV > or = " + str(columns_1["Pre-Tab OCV"]) + " V"
                    graphData = round(data[data >= columns_1["Pre-Tab OCV"]], 3)
                    self.x = columns_1["Pre-Tab OCV"]
                    print("Tabbed and pre OCV")
                elif "CCV" in type:
                    print("Tabbed and pre CCV: None")
                    return

            elif columns_1["Section No."] == 2 and "Post" in type:
                if "OCV-CCV" in type:
                    status = "CCV > or = " + str(columns_1["Profile One CCV Min"]) + " V"
                    OCV_CCV = data[data["Pre-OCV"] >= columns_1["Profile One OCV Min"]]["Pre-OCV"] - data[data["Pre-CCV"] >= columns_1["Profile One CCV Min"]]["Pre-CCV"]
                    graphData = round(OCV_CCV.dropna(), 3)
                    print("sec2 post OCV-CCV")
                elif "OCV" in type:
                    status = "OCV > or = " + str(columns_1["Profile One OCV Min"]) + " V"
                    graphData = round(data[data >= columns_1["Profile One OCV Min"]], 3)
                    self.x = columns_1["Profile One OCV Min"]
                    print("sec2 post OCV")
                elif "CCV" in type:
                    status = "CCV > or = " + str(columns_1["Profile One CCV Min"]) + " V"
                    graphData = round(data[data >= columns_1["Profile One CCV Min"]], 3)
                    self.x = columns_1["Profile One CCV Min"]
                    print("sec2 post CCV")

            elif columns_1["Profile No."] == 2 and "Post" in type:
                if "OCV-CCV" in type:
                    status = "CCV > or = " + str(columns_1["Profile Two CCV Min"]) + " V"
                    OCV_CCV = data[data["Post-OCV"] >= columns_1["Profile Two OCV Min"]]["Post-OCV"] - data[data["Post-CCV"] >= columns_1["Profile Two CCV Min"]]["Post-CCV"]
                    graphData = round(OCV_CCV.dropna(), 3)
                    print("pro2 post OCV-CCV")
                elif "OCV" in type:
                    status = "OCV > or = " + str(columns_1["Profile Two OCV Min"]) + " V"
                    graphData = round(data[data >= columns_1["Profile Two OCV Min"]], 3)
                    self.x = columns_1["Profile Two OCV Min"]
                    print("pro2 post OCV")
                elif "CCV" in type:
                    status = "CCV > or = " + str(columns_1["Profile Two CCV Min"]) + " V"
                    graphData = round(data[data >= columns_1["Profile Two CCV Min"]], 3)
                    self.x = columns_1["Profile Two CCV Min"]
                    print("pro2 post CCV")

            else:
                if "OCV-CCV" in type:
                    status = "CCV > or = " + str(columns_1["Profile One CCV Min"]) + " V"
                    OCV_CCV = data[data["Pre-OCV"] >= columns_1["Profile One OCV Min"]]["Pre-OCV"] - data[data["Pre-CCV"] >= columns_1["Profile One CCV Min"]]["Pre-CCV"]
                    graphData = round(OCV_CCV.dropna(), 3)
                    print("Profile one for all OCV-CCV")
                elif "OCV" in type:
                    status = "OCV > or = " + str(columns_1["Profile One OCV Min"]) + " V"
                    graphData = round(data[data >= columns_1["Profile One OCV Min"]], 3)
                    self.x = columns_1["Profile One OCV Min"]
                    print("Profile one for all OCV")
                elif "CCV" in type:
                    status = "CCV > or = " + str(columns_1["Profile One CCV Min"]) + " V"
                    graphData = round(data[data >= columns_1["Profile One CCV Min"]], 3)
                    self.x = columns_1["Profile One CCV Min"]
                    print("Profile one for all CCV")

        else:
            if "OCV-CCV 2" in type:
                status = "Total Lot"
                graphData = round(data["Post-OCV"] - data["Post-CCV"], 3)
                print("Everything post for OCV-CCV")
            elif "OCV-CCV" in type:
                status = "Total Lot"
                graphData = round(data["Pre-OCV"] - data["Pre-CCV"], 3)
                print("Everything for OCV-CCV")
            else:
                status = "Total Lot"
                print("Everything !")
                graphData = round(data, 3)
        return graphData, status

    def screeningGraphsmallOCVorCCV(self, data, columns_1, type):
        sub = str.maketrans("0123456789", "₀₁₂₃₄₅₆₇₈₉")
        # for building graphs in PyQt5
        graphData, status = self.findingCriteraforGraph(data, columns_1, type)
        if "confidence interval" in type:
            cofidence_interval = self.mean_confidence_interval(graphData)
            maxToMin = round(cofidence_interval[2]-cofidence_interval[1], 3)
        else:
            maxToMin = round(graphData.max() - graphData.min(), 3)
        # if the maxToMin is less then 0.02 mark everthing out
        if maxToMin <= 0.02:
            bar_width = 0.0005
            xticks_step = 0.001
            # looking for 95% confidence interval, to get the min and max of the 95% confidence interval for the graph
            if "confidence interval" in type:
                cofidence_interval = self.mean_confidence_interval(graphData)
                graphstart = round(cofidence_interval[1], 3)
                graphend = round(cofidence_interval[2], 3)
                graphData = graphData[graphData.between(graphstart, graphend)]
            else:
                graphstart = graphData.min() - 0.001
                graphend = graphData.max() + 0.001

            # x is for how many time each data show up
            x = graphData.value_counts()
            # Y_Value and X_value are actually reversed Y_Value should be X_Value.
            Y_Value = []
            X_Value = []
            for val, cnt in x.iteritems():
                Y_Value.append(val)
                X_Value.append(cnt)
        else:
            # if the maxToMin is bigger than 0.02 the graph should looking benning the data more the 0.001,
            # this function will group all the data between the ben and remake the data.
            xticks_step = (maxToMin // 0.02) * 0.001
            bar_width = xticks_step / 2
            if "confidence interval" in type:
                cofidence_interval = self.mean_confidence_interval(graphData)
                graphstart = round(cofidence_interval[1], 3)
                graphend = round(cofidence_interval[2], 3)
                graphData = graphData[graphData.between(graphstart, graphend-0.001)]
            else:
                graphstart = round(graphData.min() - xticks_step, 3)
                graphend = round(graphData.max() + xticks_step, 3)

            endOfTheInterval = round(graphstart, 3)
            X_Value = []
            Y_Value = []
            if xticks_step == 0.001:
                x = graphData.value_counts()
                for val, cnt in x.iteritems():
                    Y_Value.append(val)
                    X_Value.append(cnt)
            else:
                while endOfTheInterval <= graphend:
                    middleOfTheInterval = round((2 * endOfTheInterval + xticks_step) / 2, 3)
                    Y_Value.append(middleOfTheInterval)
                    if endOfTheInterval == graphstart:
                        X_Value.append(len(graphData[graphData.between(endOfTheInterval, round(endOfTheInterval + xticks_step, 3))]))

                    else:
                        X_Value.append(
                            len(graphData[graphData.between(endOfTheInterval + 0.0001, round(endOfTheInterval + xticks_step, 3))]))
                    endOfTheInterval = round(endOfTheInterval + xticks_step, 3)

        fig, ax = plt.subplots()

        if "WITHIN" in type:
            plt.title(self.PrePost.upper() + " SCREENING : " + type[4:] + " HISTOGRAM", fontsize=22)
        elif "OCV 2" in type or "OCV-CCV 2" in type or "CCV 2" in type:
            plt.title(self.PrePost.upper() + " SCREENING : " + type[:-1] + " HISTOGRAM", fontsize=22)
        else:
            plt.title(self.PrePost.upper() + " SCREENING: " + type + " HISTOGRAM", fontsize=22)
        fig.set_size_inches(23, 18)
        plt.rcParams.update({"font.size": 22})
        plt.tick_params(labelsize=22)
        plt.bar(Y_Value, X_Value, width=bar_width)
        # plt.xticks(np.arange(min(Y_Value), max(Y_Value)+1, 0.05))
        # plt.xticks(Z_Value)
        plt.xticks(np.arange(graphstart, graphend+0.001, step=xticks_step), fontsize=14)

        if maxToMin <= 0.02:
            for i in range(len(Y_Value)):
                if X_Value[i] != 0:
                    plt.text(Y_Value[i] - (xticks_step * .1), X_Value[i] + 0.01*X_Value[i], str(X_Value[i]), fontsize=20)
        else:
            for i in range(len(Y_Value)):
                if X_Value[i] != 0:
                    plt.text(Y_Value[i] - (xticks_step * .2), X_Value[i] + 0.01*X_Value[i], str(X_Value[i]), fontsize=15)

        if "OCV" in type:
            plt.xlabel("Open Circuit Voltage (Volts)", fontsize=22)
        else:
            plt.xlabel("Closed Circuit Voltage (Volts)", fontsize=22)
        plt.ylabel("Frequency", fontsize=22)
        plt.text(0.005, 0.995, 'Test Number: ' + str(columns_1["Test No"]), horizontalalignment='left', verticalalignment='top',
                 transform=ax.transAxes)
        plt.text(0.005, 0.975, 'Sample Name:' + str(columns_1["Cell Name"]), horizontalalignment='left',
                 verticalalignment='top',
                 transform=ax.transAxes)
        plt.text(0.005, 0.955, 'Status: ' + status, horizontalalignment='left',
                 verticalalignment='top',
                 transform=ax.transAxes)
        plt.suptitle("Chemistry: " + str(columns_1["Chemistry"]).translate(sub), fontsize=22, x=0.83, y=0.95)
        plt.grid(which="major", axis='y')
        # make the Yaxis increment in integer
        for axis in [ax.yaxis]:
            axis.set_major_locator(ticker.MaxNLocator(integer=True))
        plt.margins(y=0.1, tight=True)
        #plt.show()
        fig.tight_layout()
        fig.savefig(type + ".png", dpi=100)

    def getTestNumber5(self, x, PreOrPost):
        path_datafile = os.getcwd() + r"\\Screening_Data\\" + x
        data_file = pd.read_csv(path_datafile, sep="\t")
        # Load the template info for report
        path_template = os.getcwd() + r"\\Screening_Template\\" + x
        data_file1 = pd.read_csv(path_template, sep="\t")
        columns_1 = data_file1.loc[0].fillna("")
        testNumber = x[:-4]
        self.label_3.setText(testNumber)
        # PreOrPost to choice which graph to build
        # build the pre Graph
        if PreOrPost == "Pre":
            self.PrePost = "Pre"
            if not data_file["Pre-OCV"].isnull().all():
                self.screeningGraphsmallOCVorCCV(data_file["Pre-OCV"].dropna(), columns_1, "Pre OCV WITHIN SD RANGE")
                self.screeningGraphsmallOCVorCCV(data_file["Pre-OCV"].dropna(), columns_1, "OCV")
                self.screeningGraphsmallOCVorCCV(data_file["Pre-OCV"].dropna(), columns_1, "Pre OCV 95% confidence interval")
                self.pushButton_17.setIcon(QtGui.QIcon(QtGui.QPixmap("Pre OCV WITHIN SD RANGE.png")))
                self.pushButton_16.setIcon(QtGui.QIcon(QtGui.QPixmap("OCV.png")))
                self.pushButton_12.setIcon(QtGui.QIcon(QtGui.QPixmap("Pre OCV 95% confidence interval.png")))
                pixmap = QPixmap("Pre OCV WITHIN SD RANGE.png")
                self.label_12.setPixmap(pixmap)

            if not data_file["Pre-CCV"].isnull().all():
                if columns_1["Tabbed?"] == "Not Tabbed":
                    self.screeningGraphsmallOCVorCCV(data_file["Pre-CCV"].dropna(), columns_1,
                                                     "Pre CCV 95% confidence interval")
                    self.pushButton_15.setIcon(QtGui.QIcon(QtGui.QPixmap("Pre CCV 95% confidence interval.png")))
                    self.screeningGraphsmallOCVorCCV(data_file["Pre-CCV"].dropna(), columns_1, "Pre CCV WITHIN SD RANGE")
                    self.screeningGraphsmallOCVorCCV(data_file["Pre-CCV"].dropna(), columns_1, "CCV")
                    self.pushButton_19.setIcon(QtGui.QIcon(QtGui.QPixmap("Pre CCV WITHIN SD RANGE.png")))
                    self.pushButton_18.setIcon(QtGui.QIcon(QtGui.QPixmap("CCV.png")))
                    # making OCV-CCV graph
                    self.screeningGraphsmallOCVorCCV(data_file, columns_1, "Pre OCV-CCV WITHIN SD RANGE")
                    self.screeningGraphsmallOCVorCCV(data_file, columns_1, "OCV-CCV")
                    self.pushButton_21.setIcon(QtGui.QIcon(QtGui.QPixmap("Pre OCV-CCV WITHIN SD RANGE.png")))
                    self.pushButton_20.setIcon(QtGui.QIcon(QtGui.QPixmap("OCV-CCV")))
            else:
                # there is not Pre-Tabb CCV so i am going to make all the icon blank page
                self.pushButton_15.setIcon(QtGui.QIcon(QtGui.QPixmap("BlankPage.png")))
                self.pushButton_15.setEnabled(False)
                self.pushButton_19.setIcon(QtGui.QIcon(QtGui.QPixmap("BlankPage.png")))
                self.pushButton_19.setEnabled(False)
                self.pushButton_18.setIcon(QtGui.QIcon(QtGui.QPixmap("BlankPage.png")))
                self.pushButton_18.setEnabled(False)
                self.pushButton_21.setIcon(QtGui.QIcon(QtGui.QPixmap("BlankPage.png")))
                self.pushButton_21.setEnabled(False)
                self.pushButton_20.setIcon(QtGui.QIcon(QtGui.QPixmap("BlankPage.png")))
                self.pushButton_20.setEnabled(False)

        # builind the post graph
        # make globe verable for checking pre or post for the buttom click
        if PreOrPost == "Post":
            self.PrePost = "Post"
            self.setWindowTitle("Post-Graph")
            if not data_file["Post-OCV"].isnull().all():
                self.screeningGraphsmallOCVorCCV(data_file["Post-OCV"].dropna(), columns_1, "Post OCV WITHIN SD RANGE")
                self.screeningGraphsmallOCVorCCV(data_file["Post-OCV"].dropna(), columns_1, "OCV 2")
                self.screeningGraphsmallOCVorCCV(data_file["Post-OCV"].dropna(), columns_1,
                                                 "Post OCV 95% confidence interval")
                self.pushButton_17.setIcon(QtGui.QIcon(QtGui.QPixmap("Post OCV WITHIN SD RANGE.png")))
                self.pushButton_16.setIcon(QtGui.QIcon(QtGui.QPixmap("OCV 2.png")))
                self.pushButton_12.setIcon(QtGui.QIcon(QtGui.QPixmap("Post OCV 95% confidence interval.png")))
                pixmap = QPixmap("Post OCV WITHIN SD RANGE.png")
                self.label_12.setPixmap(pixmap)

            if not data_file["Post-CCV"].isnull().all():
                self.screeningGraphsmallOCVorCCV(data_file["Post-CCV"].dropna(), columns_1,
                                                 "Post CCV 95% confidence interval")
                self.pushButton_15.setIcon(QtGui.QIcon(QtGui.QPixmap("Post CCV 95% confidence interval.png")))
                self.screeningGraphsmallOCVorCCV(data_file["Post-CCV"].dropna(), columns_1, "Post CCV WITHIN SD RANGE")
                self.screeningGraphsmallOCVorCCV(data_file["Post-CCV"].dropna(), columns_1, "CCV 2")
                self.pushButton_19.setIcon(QtGui.QIcon(QtGui.QPixmap("Post CCV WITHIN SD RANGE.png")))
                self.pushButton_18.setIcon(QtGui.QIcon(QtGui.QPixmap("CCV 2.png")))
                # making OCV-CCV graph
                self.screeningGraphsmallOCVorCCV(data_file, columns_1, "Post OCV-CCV WITHIN SD RANGE")
                self.screeningGraphsmallOCVorCCV(data_file, columns_1, "OCV-CCV 2")
                self.pushButton_21.setIcon(QtGui.QIcon(QtGui.QPixmap("Post OCV-CCV WITHIN SD RANGE.png")))
                self.pushButton_20.setIcon(QtGui.QIcon(QtGui.QPixmap("OCV-CCV 2")))


# for debugging propose
if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    qt_app = Graph_Pre_WithFunction()
    qt_app.getTestNumber5("14856-00.txt", "Pre")
    qt_app.show()
    app.exec_()


