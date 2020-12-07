from PyQt5 import QtWidgets
from PyQt5 import QtGui
import sys
from PyQt5.QtWidgets import QMainWindow
from Screening_System_PyQt5 import QtTesting2
import pandas as pd
import glob, os
pd.options.display.max_columns = 999
pd.options.display.max_rows = 999
pd.set_option("display.precision", 6)


#update the file_link here
file_link = r"C:\Users\wangp\PycharmProjects\untitled\Screening_System_PyQt5\Screening_Data\435823-04CCK.086.txt"
data_file = pd.read_csv(file_link, skiprows=3, sep="\t")
#print(data_file)
#print(data_file.shape)
#print(data_file.shape[1])


class MyQtApp(QtTesting2.Ui_MainWindow, QtWidgets.QMainWindow):
    def __init__(self):
        super(MyQtApp, self).__init__()
        self.setupUi(self)
        #open with Maxwindow size
        #self.showMaximized()
        self.setWindowTitle("Screening System - Build Alpha 1.0")
        self.populate_tree_widget()
        self.populate_list_widget()
        self.submit_PB.clicked.connect(self.fill_form)
        self.school_TB.clicked.connect(self.select_photo)

    def populate_tree_widget(self):
        self.tableWidget_1.clear()
        Test_Time = data_file["Test Time (Hr)"]
        #item = QtWidgets.QTableWidgetItem(self.tableWidget_1)
        self.tableWidget_1.setRowCount(data_file.shape[0])
        self.tableWidget_1.setColumnCount(data_file.shape[1])
        headerList = []
        #print(headerList)
        for data_header in data_file:
            headerList.append(data_header)
        #print(headerList)
        self.tableWidget_1.setHorizontalHeaderLabels(headerList)
        #for listing all the panda datas
        for index, data in data_file.iterrows():
            #print("index=" + str(index))
            for column_number, eachrow in enumerate(data):
                #print("Eachrow=" + str(eachrow))
                self.tableWidget_1.setItem(index, column_number, QtWidgets.QTableWidgetItem(str(eachrow)))

    def populate_list_widget(self):
        self.listWidget.clear()
        os.chdir("Screening_Data")
        Screening_Data_file = []
        for file in glob.glob("*.txt"):
            self.listWidget.addItem(file)


    def fill_form(self):
        name = self.name_LE.text()
        if not name:
            QtWidgets.QMessageBox.about(self, "Name Required", "Hey! Fill the Name box")
            self.name_LE.setStyleSheet("QLineEdit { background-color: red }")
            return
        school_LE = self.school_LE.text()
        if not school_LE:
            QtWidgets.QMessageBox.about(self, "School Required", "Please, insert the school link")
            self.school_LE.setStyleSheet("QLineEdit { background-color: red }")
            return
        photo = self.school_LE.text()
        print(photo)


    def select_photo(self):
        photo_path, ext = QtWidgets.QFileDialog.getOpenFileName(self, "Select Photo")
        if photo_path:
            self.school_LE.setText(photo_path)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    qt_app = MyQtApp()
    qt_app.show()
    app.exec_()