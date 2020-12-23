import os
import math
import pandas as pd
import time
from PyQt5.QtWidgets import QDialog
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.uic import loadUi
from Screening_System_PyQt5 import client
import datetime
from multiprocessing import Process
import sys
from PyQt5.QtCore import QThread
import time
from PyQt5.QtCore import Qt, QThread, pyqtSignal

from win32com.client import Dispatch
from PyQt5.QtCore import QTimer, QTime

pd.options.display.max_columns = 999
pd.options.display.max_rows = 999
pd.set_option("display.precision", 6)

'''
class External(QThread):
    """
    Runs a counter thread.
    """
    countChanged = pyqtSignal(int)
    finalOutput = pyqtSignal(str)

    def __init__(self, timer, testing_type, testing_value2, parent=None):
        QThread.__init__(self, parent)
        self.timer = timer
        self.testing_type = testing_type
        self.testing_value2 = testing_value2
        self.load = Dispatch('BKServers.DCLoad85xx')

    def run(self):

        port = "COM3"
        baudrate = "19200"

        def test(cmd, results):
            if results:
                print(cmd + "failed:")
                print(results)
                exit(1)
            else:
                print(cmd)

        self.load.Initialize(port, baudrate)
        test("Set to remote control", self.load.SetRemoteControl())
        test("Set Remote Sense to enable", self.load.SetRemoteSense(1))

        if self.testing_type == "Constant Current":
            print("Peterlooking")
            print(int(self.timer) * 10)
            testing_value = self.testing_value2 / 1000
            print(testing_value)
            test("Set to constant current", self.load.Setmode("cc"))
            test("Set Transient to CC ",
                 self.load.SetTransient("cc", 0, 0.01, testing_value, int(self.timer) * 10, "pulse"))
        elif self.testing_type == "Constant Resistor":
            test("Set to constant current", self.load.Setmode("cr"))
            test("Set Transient to CC ",
                 self.load.SetTransient("cr", 4000, 1, self.testing_value2, int(self.timer) * 10, "pulse"))

        test("Set function to Transient", self.load.SetFunction("transient"))
        self.load.TurnLoadOn()
        self.load.TriggerLoad()

        start_time = time.time()
        t_end = time.time() + int(self.timer) + 1
        values = []
        ProBarcount = 0
        # this is increament for progressbar in while loop
        ProBarIncreamnt = 7/(int(self.timer) + 1)

        while time.time() < t_end:
            values.append(self.load.GetInputValues()[0])
            time.sleep(0.01)
            ProBarcount = ProBarcount + ProBarIncreamnt
            self.countChanged.emit(ProBarcount)

        self.countChanged.emit(100)
        print("Final values is ")
        print(values)

        self.load.TurnLoadOff()
        test("Set Function to fix", self.load.SetFunction("fixed"))
        test("Set to local control", self.load.SetLocalControl())

        self.finalOutput.emit(str(min(values)))
'''


class Screening_DataCollection_WithFunction(QDialog):

    def __init__(self):
        super(Screening_DataCollection_WithFunction, self).__init__()
        loadUi("RGui_Screening_DataCollection.ui", self)
        self.lineEdit_1.returnPressed.connect(self.scanBarcode)
        self.lineEdit_2.returnPressed.connect(self.Sample_Sn)
        self.lineEdit_3.returnPressed.connect(self.date_ReturnPressed)
        self.lineEdit_4.returnPressed.connect(self.startProgessBar)
        self.lineEdit_6.returnPressed.connect(self.CommentsPressed)
        self.pushButton.clicked.connect(self.CancelButton)
        self.pushButton_2.clicked.connect(self.SaveButton)
        self.pushButton_3.clicked.connect(self.date_ReturnPressed)
        self.pushButton_4.clicked.connect(self.startProgessBar)
        self.timer = QTimer()
        self.counter = 0

    storing_data = []
    lot_number = 0

    def pwCCV(self, timer, testing_type, testing_value2,):
        load = Dispatch('BKServers.DCLoad85xx')
        port = "COM6"
        baudrate = "19200"
        def test(cmd, results):
            if results:
                print(cmd + "failed:")
                print(results)
                exit(1)
            else:
                print(cmd)
        load.Initialize(port, baudrate)
        test("Set to remote control", load.SetRemoteControl())
        test("Set Remote Sense to enable", load.SetRemoteSense(1))

        if testing_type == "Constant Current":
            testing_value = testing_value2 / 1000
            test("Set to constant current", load.Setmode("cc"))
            test("Set Transient to CC ",
                 load.SetTransient("cc", 0, 0.01, testing_value, int(timer) * 10, "pulse"))
        elif testing_type == "Constant Resistor":
            test("Set to constant Resistor", load.Setmode("cr"))
            test("Set Transient to CC ",
                 load.SetTransient("cr", 4000, 1, testing_value2, int(timer) * 10, "pulse"))

        test("Set function to Transient", load.SetFunction("transient"))
        load.TurnLoadOn()
        load.TriggerLoad()

        start_time = time.time()
        t_end = time.time() + int(timer) + 1
        values = []
        ProBarcount = 0
        # this is increament for progressbar in while loop
        ProBarIncreamnt = 7/(int(timer) + 1)

        while time.time() < t_end:
            values.append(load.GetInputValues()[0])
            time.sleep(0.01)
            ProBarcount = ProBarcount + ProBarIncreamnt

        print("Final values is ")
        print(values)

        load.TurnLoadOff()
        test("Set Function to fix", load.SetFunction("fixed"))
        test("Set to local control", load.SetLocalControl())
        value = min(values)
        self.lineEdit_5.setText(str(value))
        if self.label.text() != "Pre-Tabbed":
            print("Not Pre Tabbed")
            if float(value) < float(self.label_35.text()):
                self.label_15.setText(self.label_15.text() + "The CCV criteria was not met!")
            else:
                self.label_15.setText(self.label_15.text() + " ")
        # adding OCV tolerance:
        if self.label.text() == "Post-Tabbed":
            if abs(float(self.label_39.text()) - float(self.lineEdit_4.text())) > float(self.label_40.text()):
                self.label_15.setText(self.label_15.text() + " The ocv tolerance is exceed!")
        self.lineEdit_6.setFocus()

    def CancelButton(self):
        self.close()

    def CommentsPressed(self):
        self.pushButton_2.setDefault(True)
        self.pushButton_2.setFocus()

    def SaveButton(self):
        print("SaveButton pressed")
        #dir_path = os.path.dirname(os.path.realpath(__file__))
        dir_path = os.path.dirname(sys.argv[0])
        path_data = dir_path + r"\Screening_Data\\" + self.label_1.text()+".txt"
        #if there is no file, created the file
        if os.path.exists(path_data) is False:
            df = pd.DataFrame(columns=["Barcode", "Serial#", "Pre-OCV", "Pre-CCV", "Post-OCV", "Post-CCV", "Date",
                                       "Lot Number", "Comments", "Pre-screen pass", "Post-screen pass"])
            df.to_csv(path_data, sep="\t", index=False)
        columns_1 = []
        #if this is pre-screening, store in here
        if self.label.text() == "Profile One" or self.label.text() == "Pre-Tabbed" or self.label.text() == "Section One":
            print("PeterTesting333:"+self.label.text())
            columns_1.append(self.lineEdit_1.text())
            columns_1.append(self.lineEdit_2.text())
            columns_1.append(self.lineEdit_4.text())
            if self.label.text() == "Pre-Tabbed":
                columns_1.append("")
            else:
                columns_1.append(self.lineEdit_5.text())
            columns_1.append("")
            columns_1.append("")
            columns_1.append(self.lineEdit_3.text())
            columns_1.append(self.lot_number)
            columns_1.append(self.lineEdit_6.text())
            #check if the criteria was met
            if float(self.lineEdit_4.text()) < float(self.label_33.text()):
                columns_1.append("Fail")
            elif (self.label.text() == "Profile One" or self.label.text() == "Section One") and (float(self.lineEdit_5.text()) < float(self.label_35.text())):
                columns_1.append("Fail")
            else:
                columns_1.append("Pass")
            columns_1.append("")

            data_file = pd.read_csv(path_data, sep="\t")
            BarcodeCunt = data_file["Barcode"].last_valid_index()
            if BarcodeCunt is None:
                BarcodeCunt = 0
            else:
                BarcodeCunt = BarcodeCunt + 1

            print("PeterTesting333:"+str(columns_1))
            data_file.loc[BarcodeCunt] = columns_1
            data_file.to_csv(path_data, sep="\t", index=False)

            ui = Screening_DataCollection_WithFunction()
            ui.getTestNumber(self.label_1.text()+".txt")
            self.close()
            ui.show()
            ui.exec_()
        else:
            #this is post-screening, store within the pre-screening
            data_file = pd.read_csv(path_data, sep="\t")
            if int(self.lineEdit_1.text()) in data_file["Barcode"].values:
                data_file.at[data_file[data_file["Barcode"] == int(self.lineEdit_1.text())].index[0], "Post-OCV"] = self.lineEdit_4.text()
                data_file.at[data_file[data_file["Barcode"] == int(self.lineEdit_1.text())].index[0], "Post-CCV"] = self.lineEdit_5.text()
            if float(self.lineEdit_4.text()) < float(self.label_33.text()):
                data_file.loc[data_file[data_file["Barcode"] == int(self.lineEdit_1.text())].index[0], "Post-screen pass"] = "Fail"
            elif float(self.lineEdit_5.text()) < float(self.label_35.text()) and (self.label.text() == "Profile One" or
                                                                                  self.label.text() == "Section One"):
                data_file.loc[data_file[data_file["Barcode"] == int(self.lineEdit_1.text())].index[0], "Post-screen pass"] = "Fail"
            else:
                data_file.loc[data_file[data_file["Barcode"] == int(self.lineEdit_1.text())].index[0], "Post-screen pass"] = "Pass"
            data_file.to_csv(path_data, sep="\t", index=False)
            ui = Screening_DataCollection_WithFunction()
            ui.getTestNumber("Post" + self.label_1.text() + ".txt")
            self.close()
            ui.show()
            ui.exec_()

    def startProgessBar(self):
        self.label_33.setFocus()
        if self.lcdNumber.value() == 0.0 or self.label.text() == "Pre-Tabbed":
            self.lineEdit_6.setFocus()
            return
        else:
            testing_type = str(self.label_28.text())
            testing_value2 = float(self.label_30.text())
            timer = int(self.lcdNumber.value())

            # passin all the varabile to QThread
            # peterfixing1
            self.pwCCV(timer=timer, testing_type=testing_type, testing_value2=testing_value2)

            #self.calc = External(timer=timer, testing_type=testing_type, testing_value2=testing_value2)
            #self.calc.countChanged.connect(self.onCountChanged)
            #self.calc.finalOutput.connect(self.onCountChanged2)
            #self.calc.start()

    # def onCountChanged(self, value):
    #     # receive the emit singal to change ProgressBar
    #     self.progressBar.setValue(value)

    def onCountChanged2(self, value):
        self.lineEdit_5.setText(str(value))
        # checking with criteria
        if self.label.text() is not "Pre-Tabbed":
            print("Not Pre Tabbed")
            if float(value) < float(self.label_35.text()):
                self.label_15.setText(self.label_15.text() + "The CCV criteria was not met!")
            else:
                self.label_15.setText(self.label_15.text() + " ")
        self.lineEdit_6.setFocus()

    def getOCVpw(self):
        load = Dispatch('BKServers.DCLoad85xx')
        port = "COM6"
        baudrate = "19200"

        def test(cmd, results):
            if results:
                print(cmd + "failed:")
                print(results)
                exit(1)
            else:
                print(cmd)

        load.Initialize(port, baudrate)
        test("Set to remote control", load.SetRemoteControl())
        test("Set Remote Sense to enable", load.SetRemoteSense(1))
        load.TurnLoadOn()
        values = load.GetInputValues()
        load.TurnLoadOff()
        test("Set to local control", load.SetLocalControl())
        return values

    def date_ReturnPressed(self):
        OCV = self.getOCVpw()
        if float(OCV[0]) < float(self.label_33.text()):
            self.label_15.setText("The OCV criteria was not met!       ")
        else:
            self.label_15.setText(" ")
        self.lineEdit_4.setText(str(OCV[0]))
        self.lineEdit_4.setFocus()

    def Sample_Sn(self):
        self.lineEdit_3.setText(str(datetime.date.today()))
        self.lineEdit_3.setFocus()

    def scanBarcode(self):
        self.lineEdit_2.setFocus()
        #dir_path = os.path.dirname(os.path.realpath(__file__))
        dir_path = os.path.dirname(sys.argv[0])
        path_data = dir_path + r"\Screening_Data\\" + self.label_1.text() + ".txt"
        #check if barcode exit for pre-screening or barcode doesn't exit in post-screening or scanne twice with post-screening
        if os.path.exists(path_data) is True:
            data_file = pd.read_csv(path_data, sep="\t")
            if self.label.text() == "Profile One" or self.label.text() == "Section One" or self.label.text() == "Pre-Tabbed":
                if int(self.lineEdit_1.text()) in data_file["Barcode"].values:
                    #msgbox = QtWidgets.QMessageBox(self)
                    #msgbox.setText("This Barcode is already used, please check.")
                    #msgbox.exec()
                    buttonReply = QtWidgets.QMessageBox.question(self, 'Warning',
                                                                 "Barcode already exists, "
                                                                 "do you want to replace value for the same barcode?",
                                                                 QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                                                 QtWidgets.QMessageBox.No)
                    if buttonReply == QtWidgets.QMessageBox.Yes:
                        data_file = data_file[data_file["Barcode"] != int(self.lineEdit_1.text())]
                        data_file.to_csv(path_data, sep="\t", index=False)
                        self.lineEdit_2.setFocus()
                    else:
                        self.lineEdit_1.setFocus()
                    return
                else:
                    self.lineEdit_2.setFocus()
            else:
                if int(self.lineEdit_1.text()) in data_file["Barcode"].values and math.isnan(data_file[data_file["Barcode"] == int(self.lineEdit_1.text())]["Post-OCV"]):
                    # pre-value show up
                    self.label_36.setText('Pre-Value: ')
                    print(str(data_file[data_file["Barcode"] == int(self.lineEdit_1.text())]["Pre-OCV"]))
                    self.label_39.setText(str(data_file[data_file["Barcode"] == int(self.lineEdit_1.text())]["Pre-OCV"].item()))
                    if self.label.text() != "Post-Tabbed":
                        self.label_37.setText('Pre-Value: ' + str(
                            data_file[data_file["Barcode"] == int(self.lineEdit_1.text())]["Pre-CCV"].values))
                    self.lineEdit_2.setFocus()
                else:
                    msgbox = QtWidgets.QMessageBox(self)
                    msgbox.setText("This Barcode you scanned has No pre-screening value. Or This Post-screening value already recorded.")
                    msgbox.exec()
                    self.lineEdit_1.setFocus()

    def getTestNumber(self, x):
        if "Post" in x:
            y = x[4:]
            # dir_path = os.path.dirname(os.path.realpath(__file__))
            dir_path = os.path.dirname(sys.argv[0])
            path_Template = dir_path + r"\\Screening_Template\\" + y
        else:
            # dir_path = os.path.dirname(os.path.realpath(__file__))
            dir_path = os.path.dirname(sys.argv[0])
            path_Template = dir_path + r"\\Screening_Template\\" + x
        data_file = pd.read_csv(path_Template, sep="\t")
        columns_1 = data_file.loc[0].fillna("")
        self.lot_number = columns_1[0]
        self.label_1.setText(str(columns_1[1]))
        self.label_5.setText(str(columns_1[1]))
        if "Post" in x:
            if columns_1[33] == 2:
                print("Profile two")
                self.label.setText("Profile Two")
            elif columns_1[32] == "Tabbed":
                print("Tabbed")
                self.label.setText("Post-Tabbed")
            elif columns_1[34] == 2:
                print("Section 2")
                self.label.setText("Section Two")

            if columns_1[37] == "Constant Current":
                self.label_28.setText("Constant Current")
            else:
                self.label_28.setText("Constant Resistor")
                self.label_31.setText(" Ω")
        else:
            if columns_1[33] == 2:
                print("Profile two")
                self.label.setText("Profile One")
            elif columns_1[32] == "Tabbed":
                print("Tabbed")
                self.label.setText("Pre-Tabbed")
                self.lineEdit_5.setEnabled(False)
            elif columns_1[34] == 2:
                print("Section 2")
                self.label.setText("Section One")
            else:
                self.label.setText("Profile One")

            if columns_1[36] == "Constant Current":
                self.label_28.setText("Constant Current")
            else:
                self.label_28.setText("Constant Resistor")
                self.label_31.setText(" Ω")

        if self.label.text() == "Profile One" or self.label.text() == "Section One" or self.label.text() == "Section Two":
            self.label_30.setText(str(columns_1[17]))
            self.lcdNumber.display(columns_1[18])
            self.label_33.setText(str(columns_1[19]))
            self.label_35.setText(str(columns_1[20]))
        elif self.label.text() == "Pre-Tabbed":
            self.label_30.setText("")
            self.lcdNumber.display(0.0)
            self.label_33.setText(str(columns_1[23]))
            self.label_35.setText("")
        elif self.label.text() == "Profile Two":
            self.label_30.setText(str(columns_1[27]))
            self.lcdNumber.display(columns_1[28])
            self.label_33.setText(str(columns_1[29]))
            self.label_35.setText(str(columns_1[30]))
        elif self.label.text() == "Post-Tabbed":
            self.label_30.setText(str(columns_1[17]))
            self.lcdNumber.display(columns_1[18])
            self.label_33.setText(str(columns_1[24]))
            self.label_35.setText(str(columns_1[25]))
            self.label_38.setText("Tolerance: ")
            self.label_40.setText(str(columns_1['OCV Tab Tolerance']))
            if columns_1[36] == "Constant Current":
                self.label_28.setText("Constant Current")
            else:
                self.label_28.setText("Constant Resistor")
                self.label_31.setText(" Ω")
        path_data = dir_path + r"\Screening_Data\\" + self.label_1.text() + ".txt"
        if os.path.exists(path_data) is False:
            self.label_17.setText("0")
            self.label_19.setText("0")
            self.label_21.setText("0")
            self.label_23.setText("0")
            self.label_25.setText("0")
            self.label_27.setText("0")
        else:
            print(path_data)
            data_file = pd.read_csv(path_data, sep="\t")
            total_number = data_file["Barcode"].last_valid_index()+1
            self.label_17.setText(str(total_number))
            pre_sc_number = data_file["Pre-OCV"].last_valid_index()+1
            self.label_21.setText(str(pre_sc_number))
            if self.label.text() == "Profile One" or self.label.text() == "Section One" or self.label.text() == "Pre-Tabbed":
                if total_number >= columns_1[16]:
                    msgbox = QtWidgets.QMessageBox(self)
                    msgbox.setText("The Total sample Count is reached")
                    msgbox.exec()
            if data_file[data_file["Pre-screen pass"] == "fail"].last_valid_index() is not None:
                print(data_file[data_file["Pre-screen pass"] == "fail"])
                pre_sc_fail = data_file[data_file["Pre-screen pass"] == "fail"]["Pre-screen pass"].size
                self.label_23.setText(str(pre_sc_fail))
            if data_file["Post-OCV"].dropna().size is not 0:
                self.label_25.setText(str(data_file["Post-OCV"].dropna().size))
                if data_file[data_file["Post-screen pass"] == "fail"].last_valid_index() is not None:
                    self.label_27.setText(str(data_file[data_file["Post-screen pass"] == "fail"]["Post-screen pass"].size))

            if float(self.label_23.text()) >= float(self.label_27.text()):
                self.label_19.setText(self.label_23.text())
            else:
                self.label_19.setText(self.label_27.text())

'''
    def GettingCCV(self):
        if self.lcdNumber.value() == 0.0 or self.label.text() == "Pre-Tabbed":
            return
        else:
            self.label_36.setText("Taking CCV Right Now")
            self.lineEdit_5.setText("Taking CCV Right Now")
            testing_type = str(self.label_28.text())
            testing_value2 = float(self.label_30.text())
            timer = int(self.lcdNumber.value())

            load = Dispatch('BKServers.DCLoad85xx')
            port = "COM3"
            baudrate = "19200"

            def test(cmd, results):
                if results:
                    print(cmd + "failed:")
                    print(results)
                    exit(1)
                else:
                    print(cmd)

            load.Initialize(port, baudrate)
            test("Set to remote control", load.SetRemoteControl())
            test("Set Remote Sense to enable", load.SetRemoteSense(1))

            if testing_type == "Constant Current":
                print("Peterlooking")
                print(int(timer) * 10)
                testing_value = testing_value2 / 1000
                print(testing_value)
                test("Set to constant current", load.Setmode("cc"))
                test("Set Transient to CC ",
                     load.SetTransient("cc", 0, 0.01, testing_value, int(timer) * 10, "pulse"))
            elif testing_type == "Constant Resistor":
                test("Set to constant current", load.Setmode("cr"))
                test("Set Transient to CC ",
                     load.SetTransient("cr", 0, 0.01, testing_value2, int(timer) * 10, "pulse"))

            test("Set function to Transient", load.SetFunction("transient"))
            load.TurnLoadOn()
            load.TriggerLoad()

            msgbox = QtWidgets.QMessageBox(self)
            msgbox.setWindowTitle("Taking CCV")
            msgbox.show()

            #self.timer.setInterval(100)
            #self.timer.timeout.connect(self.storing_data.append(load.GetInputValues()[0]))
            #self.timer.start()

            start_time = time.time()
            t_end = time.time() + int(timer) + 1
            values = []

            while time.time() < t_end:
                values.append(load.GetInputValues()[0])
                time.sleep(0.05)

            msgbox.close()
            self.label_36.setText("")
            self.lineEdit_5.setText(min(values))
            self.lineEdit_6.setFocus()
            print("Final values is ")
            print(values)
            print(self.storing_data)
            load.TurnLoadOff()
            test("Set Function to fix", load.SetFunction("fixed"))
            test("Set to local control", load.SetLocalControl())

    def recurring_timer(self):
        self.storing_data.append(client.gettingValue()[0])

        self.counter += 1
        #self.lineEdit_5.setText(str(values))
        if self.counter >= 5000:
            self.timer.stop()
        def timer_start(self):
        if self.lcdNumber.value() == 0.0 or self.label.text() == "Pre-Tabbed":
            self.lineEdit_6.setFocus()
        else:
            self.storing_data = []
            self.time_left_int = self.lcdNumber.value()
            self.my_qtimer = QtCore.QTimer(self)
            self.my_qtimer.timeout.connect(self.timer_timeout)
            self.my_qtimer.start(350)
            self.update_gui()

    def timer_timeout(self):
        self.time_left_int -= .5

        self.storing_data.append(client.gettingValue()[0])
        #storing_data_2.append(client.gettingValue()[0])

        if self.time_left_int <= 0:
            self.lineEdit_6.setFocus()
            self.my_qtimer.stop()

            values = min(self.storing_data)
            self.lineEdit_5.setText(values)
            self.lineEdit_6.setFocus()
            self.my_qtimer.stop()
            #here is the fast data collection
            print(self.storing_data)
            #print(storing_data_2)
            #checking with criteria
            if self.label.text() is not "Pre-Tabbed":
                print("Not Pre Tabbed")
                if float(values) < float(self.label_35.text()):
                    self.label_15.setText(self.label_15.text()+"The CCV criteria was not met!")
                else:
                    self.label_15.setText(self.label_15.text() + " ")

        self.update_gui()

    def update_gui(self):
        self.lcdNumber.display(str(self.time_left_int))
'''

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    qt_app = Screening_DataCollection_WithFunction()
    qt_app.getTestNumber("14665A01.txt")
    #qt_app.getTestNumber2("14575A00.txt")
    qt_app.show()
    app.exec_()










