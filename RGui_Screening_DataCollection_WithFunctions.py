import os
import math
import pandas as pd
from PyQt5.QtWidgets import QDialog
from PyQt5 import QtWidgets
from PyQt5.uic import loadUi
import datetime
import sys
import time

from Screening_System_PyQt5 import dcload
from PyQt5.QtCore import QTimer, QTime

pd.options.display.max_columns = 999
pd.options.display.max_rows = 999
pd.set_option("display.precision", 6)


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
        self.replace_data = False
        self.fast_data = False
        self.storing_data = []
        self.lot_number = 0
        # on peter's desk port COM2,  At the Lab port COM3
        self.setting_file_path = os.path.dirname(sys.argv[0]) + r'\\BK_Prescions_setting.csv'
        self.setting_file = pd.read_csv(self.setting_file_path, sep=',')
        self.com_port = self.setting_file['com_port'].values[0]
        self.baudrate = self.setting_file['baurate'].values[0]
        self.BK_model = self.setting_file['model'].values[0]

    def pwCCV(self, timer, testing_type, testing_value2, ):
        load = dcload.DCLoad()

        def test(cmd, results):
            if results:
                print(cmd + "failed:")
                print(results)
                exit()
            else:
                print(cmd)
        load.Initialize(self.com_port, self.baudrate)
        test("Set to remote control", load.SetRemoteControl())
        test("Set Remote Sense to enable", load.SetRemoteSense(1))

        if testing_type == "Constant Current":
            testing_value = testing_value2 / 1000
            test("Set to constant current", load.SetMode("cc"))

            if str(self.BK_model) == "8500":
                test("Set Transient to CC ",
                     load.SetTransient("cc", 0, 1, testing_value, int(timer) * 10, "toggled"))
            else:
                test("Set Transient to CC ",
                     load.SetTransient("cc", testing_value, int(timer) * 10, 0, 1, "toggled"))
        elif testing_type == "Constant Resistor":
            test("Set to constant Resistor", load.SetMode("cr"))
            if str(self.BK_model) == "8500":
                test("Set Transient to CR ",
                     load.SetTransient("cr", 4000, 1, testing_value2, int(timer) * 10, "toggled"))
            else:
                test("Set Transient to CR ",
                     load.SetTransient("cr", testing_value2, int(timer) * 10, 4000, 1, "toggled"))

        test("Set function to Transient", load.SetFunction("transient"))
        load.TurnLoadOn()
        load.TriggerLoad()

        values = []
        data_log = []
        trigger_blo = True
        completed = 0
        progressbar_step = int(timer)
        start_time = time.time()
        t_end = time.time() + int(timer)
        if self.fast_data:
            t_end_space = time.time() + int(timer) + 4
            progressbar_step += 5
        else:
            t_end_space = time.time() + int(timer)
            progressbar_step += 1

        progressbar_step = 100/(progressbar_step*10)

        # the adding one sec is for latency
        while time.time() < (t_end_space + 1):
            values.append(load.GetInputValues()[0])
            data_log.append((load.GetInputValues()[0], load.GetInputValues()[1], round(time.time() - start_time, 5)))
            time.sleep(0.01)

            if completed < 100:
                completed += progressbar_step
                self.progressBar.setValue(completed)

            if time.time() > t_end and trigger_blo:
                load.TriggerLoad()
                trigger_blo = False

        load.TurnLoadOff()
        self.progressBar.setValue(100)

        print("Final values is ")
        print(data_log)
        # print(values)
        # print(len(values))
        # print(time_log)
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
        self.storing_data = data_log

    def CancelButton(self):
        self.close()

    def CommentsPressed(self):
        self.pushButton_2.setDefault(True)
        self.pushButton_2.setFocus()

    def SaveButton(self):
        print("SaveButton pressed")
        # dir_path = os.path.dirname(os.path.realpath(__file__))
        dir_path = os.path.dirname(sys.argv[0])
        path_data = dir_path + r"\Screening_Data\\" + self.label_1.text() + ".txt"
        fast_data_dir = dir_path + r"\Fast_data_collection\\" + self.label_1.text() + ".txt"
        # if there is no file, created the file
        if os.path.exists(path_data) is False:
            df = pd.DataFrame(columns=["Barcode", "Serial#", "Pre-OCV", "Pre-CCV", "Post-OCV", "Post-CCV", "Date",
                                       "Lot Number", "Comments", "Pre-screen pass", "Post-screen pass"])
            df.to_csv(path_data, sep="\t", index=False)
        if self.fast_data:
            if os.path.exists(fast_data_dir) is False:
                df = pd.DataFrame(columns=["Barcode", "Pre", "Post"])
                df.to_csv(fast_data_dir, sep="\t", index=False)
        columns_1 = []
        fast_data_column = []
        # if this is pre-screening, store in here
        if self.label.text() == "Profile One" or self.label.text() == "Pre-Tabbed" or self.label.text() == "Section One":
            if self.replace_data:
                print("replacing Data")
                data_file = pd.read_csv(path_data, sep="\t")

                if self.fast_data:
                    fast_data_file = pd.read_csv(fast_data_dir, sep="\t")
                    if self.label.text() != "Pre-Tabbed":
                        fast_data_file.at[
                            fast_data_file[fast_data_file["Barcode"] == int(self.lineEdit_1.text())].index[
                                0], "Pre"] = str(self.storing_data)
                    else:
                        fast_data_file.at[
                            fast_data_file[fast_data_file["Barcode"] == int(self.lineEdit_1.text())].index[
                                0], "Pre"] = ''
                    fast_data_file.to_csv(fast_data_dir, sep="\t", index=False)

                print(f"Testing3 {data_file[data_file['Barcode'] == int(self.lineEdit_1.text())].index[0]}")
                if int(self.lineEdit_1.text()) in data_file["Barcode"].values:
                    data_file.at[data_file[data_file["Barcode"] == int(self.lineEdit_1.text())].index[
                                     0], "Pre-OCV"] = self.lineEdit_4.text()
                    if self.label.text() != "Pre-Tabbed":
                        data_file.at[data_file[data_file["Barcode"] == int(self.lineEdit_1.text())].index[
                                         0], "Pre-CCV"] = self.lineEdit_5.text()
                if float(self.lineEdit_4.text()) < float(self.label_33.text()):
                    data_file.loc[data_file[data_file["Barcode"] == int(self.lineEdit_1.text())].index[
                                      0], "Pre-screen pass"] = "Fail"
                elif self.label.text() != "Pre-Tabbed" and float(self.lineEdit_5.text()) < float(self.label_35.text()):
                    data_file.loc[data_file[data_file["Barcode"] == int(self.lineEdit_1.text())].index[
                                      0], "Pre-screen pass"] = "Fail"
                else:
                    data_file.loc[data_file[data_file["Barcode"] == int(self.lineEdit_1.text())].index[
                                      0], "Pre-screen pass"] = "Pass"
                data_file.to_csv(path_data, sep="\t", index=False)
            else:
                columns_1.append(self.lineEdit_1.text())
                fast_data_column.append(self.lineEdit_1.text())
                columns_1.append(self.lineEdit_2.text())
                columns_1.append(self.lineEdit_4.text())
                if self.label.text() == "Pre-Tabbed":
                    columns_1.append("")
                    fast_data_column.append("")
                else:
                    columns_1.append(self.lineEdit_5.text())
                    fast_data_column.append(self.storing_data)

                # placeholder for the post
                fast_data_column.append('')
                columns_1.append("")
                columns_1.append("")
                columns_1.append(self.lineEdit_3.text())
                columns_1.append(self.lot_number)
                columns_1.append(self.lineEdit_6.text())
                # check if the criteria was met
                if float(self.lineEdit_4.text()) < float(self.label_33.text()):
                    columns_1.append("Fail")
                elif (self.label.text() == "Profile One" or self.label.text() == "Section One") and (
                        float(self.lineEdit_5.text()) < float(self.label_35.text())):
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

                data_file.loc[BarcodeCunt] = columns_1
                data_file.to_csv(path_data, sep="\t", index=False)

                if self.fast_data:
                    fast_data_file = pd.read_csv(fast_data_dir, sep="\t")
                    fast_data_file.loc[BarcodeCunt] = fast_data_column
                    fast_data_file.to_csv(fast_data_dir, sep="\t", index=False)

            ui = Screening_DataCollection_WithFunction()
            ui.getTestNumber(self.label_1.text() + ".txt")
            self.close()
            ui.show()
            ui.exec_()
        else:
            # this is post-screening, store within the pre-screening
            data_file = pd.read_csv(path_data, sep="\t")
            if int(self.lineEdit_1.text()) in data_file["Barcode"].values:
                data_file.at[data_file[data_file["Barcode"] == int(self.lineEdit_1.text())].index[
                                 0], "Post-OCV"] = self.lineEdit_4.text()
                data_file.at[data_file[data_file["Barcode"] == int(self.lineEdit_1.text())].index[
                                 0], "Post-CCV"] = self.lineEdit_5.text()
            if float(self.lineEdit_4.text()) < float(self.label_33.text()):
                data_file.loc[data_file[data_file["Barcode"] == int(self.lineEdit_1.text())].index[
                                  0], "Post-screen pass"] = "Fail"
            elif float(self.lineEdit_5.text()) < float(self.label_35.text()) and (self.label.text() == "Profile Two" or
                                                                                  self.label.text() == "Section Two"):
                data_file.loc[data_file[data_file["Barcode"] == int(self.lineEdit_1.text())].index[
                                  0], "Post-screen pass"] = "Fail"
            else:
                data_file.loc[data_file[data_file["Barcode"] == int(self.lineEdit_1.text())].index[
                                  0], "Post-screen pass"] = "Pass"
            data_file.to_csv(path_data, sep="\t", index=False)
            if self.fast_data:
                fast_data_file = pd.read_csv(fast_data_dir, sep="\t")
                fast_data_file.loc[fast_data_file[fast_data_file["Barcode"] == int(self.lineEdit_1.text())].index[0],
                                   "Post"] = str(self.storing_data)
                fast_data_file.to_csv(fast_data_dir, sep="\t", index=False)

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

            # passing all the variable to QThread
            self.pwCCV(timer=timer, testing_type=testing_type, testing_value2=testing_value2)

            # self.calc = External(timer=timer, testing_type=testing_type, testing_value2=testing_value2)
            # self.calc.countChanged.connect(self.onCountChanged)
            # self.calc.finalOutput.connect(self.onCountChanged2)
            # self.calc.start()

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
        load = dcload.DCLoad()
        def test(cmd, results):
            if results:
                print(cmd + "failed:")
                print(results)
                exit()
            else:
                print(cmd)

        load.Initialize(self.com_port, self.baudrate)
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

    def check_duplicated_barcode(self, data_file, state):
        data_file_2 = data_file.fillna("")
        if int(self.lineEdit_1.text()) in data_file["Barcode"].values:
            buttonReply = QtWidgets.QMessageBox.question(self, 'Warning',
                                                         "Barcode already exists, "
                                                         "do you want to replace value for the same barcode?",
                                                         QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                                         QtWidgets.QMessageBox.No)
            if buttonReply == QtWidgets.QMessageBox.Yes:
                self.replace_data = True
                self.lineEdit_2.setText(
                    data_file_2[data_file_2["Barcode"] == int(self.lineEdit_1.text())]['Serial#'].item())
                self.lineEdit_3.setText(
                    data_file_2[data_file_2["Barcode"] == int(self.lineEdit_1.text())]['Date'].item())
                self.lineEdit_4.setText(str(
                    data_file_2[data_file_2["Barcode"] == int(self.lineEdit_1.text())][state + '-OCV'].item()))
                if self.label.text() != "Pre-Tabbed":
                    self.lineEdit_5.setText(str(
                        data_file_2[data_file_2["Barcode"] == int(self.lineEdit_1.text())][state + '-CCV'].item()))
                self.lineEdit_6.setFocus()
            else:
                self.lineEdit_1.setFocus()

    def scanBarcode(self):
        self.lineEdit_2.setFocus()
        # dir_path = os.path.dirname(os.path.realpath(__file__))
        dir_path = os.path.dirname(sys.argv[0])
        path_data = dir_path + r"\Screening_Data\\" + self.label_1.text() + ".txt"
        # if the barcode is empty pop up error window
        if self.lineEdit_1.text() == "":
            msgbox = QtWidgets.QMessageBox(self)
            msgbox.setText("Barcode can't be empty, please scan barcode.")
            msgbox.exec()
            self.lineEdit_1.setFocus()
            return
        # check if barcode exit for pre-screening or barcode doesn't exit in post-screening or scan twice with
        if os.path.exists(path_data) is True:
            data_file = pd.read_csv(path_data, sep="\t")
            # pre-screening
            if self.label.text() in ["Profile One", "Section One", "Pre-Tabbed"]:
                # duplicated barcode
                self.check_duplicated_barcode(data_file, 'Pre')
            else:
                # has ocv and CCV is not nan
                # check for no OCV
                if int(self.lineEdit_1.text()) not in data_file["Barcode"].values:
                    msgbox = QtWidgets.QMessageBox(self)
                    msgbox.setText("This Barcode has NO OCV, please check.")
                    msgbox.exec()
                    self.lineEdit_1.setFocus()
                    return

                if int(self.lineEdit_1.text()) in data_file["Barcode"].values and \
                        not math.isnan(data_file[data_file["Barcode"] == int(self.lineEdit_1.text())]["Post-OCV"]):
                    self.check_duplicated_barcode(data_file, 'Post')

                # pre-value show up
                self.label_36.setText('Pre-OCV: ')
                self.label_39.setText(
                    str(round(data_file[data_file["Barcode"] == int(self.lineEdit_1.text())]["Pre-OCV"].item(),
                              4)))
                if self.label.text() != "Post-Tabbed":
                    self.label_37.setText("Pre-CCV:  " + str(
                        round(data_file[data_file["Barcode"] == int(self.lineEdit_1.text())]["Pre-CCV"].item(),
                              4)))

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
        # fast Data
        if data_file["High Rate?"].item() == "Yes":
            self.fast_data = True
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
            total_number = data_file["Barcode"].last_valid_index() + 1
            self.label_17.setText(str(total_number))
            pre_sc_number = data_file["Pre-OCV"].last_valid_index() + 1
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
                    self.label_27.setText(
                        str(data_file[data_file["Post-screen pass"] == "fail"]["Post-screen pass"].size))

            if float(self.label_23.text()) >= float(self.label_27.text()):
                self.label_19.setText(self.label_23.text())
            else:
                self.label_19.setText(self.label_27.text())


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    qt_app = Screening_DataCollection_WithFunction()
    qt_app.getTestNumber("14665A01.txt")
    # qt_app.getTestNumber2("14575A00.txt")
    qt_app.show()

    app.exec_()
