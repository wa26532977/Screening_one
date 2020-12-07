import os
import pandas as pd
from PyQt5.QtWidgets import QDialog
from PyQt5 import QtWidgets
from Screening_System_PyQt5 import Screening_Report_WordGenerator
from Screening_System_PyQt5 import RawData_Report_WordDoc
from Screening_System_PyQt5 import GraphPicture_Report_Combin
from PyQt5.uic import loadUi
from docx import Document
import sys
import _thread

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

    def combine_word_documents(self, files, testname):
        merged_document = Document()

        for index, file in enumerate(files):
            sub_doc = Document(os.getcwd() + r"\\Report_Word\\" + file)

            for element in sub_doc.element.body:
                merged_document.element.body.append(element)

        merged_document.save(os.getcwd() + r"\\Final_Report\\" + testname + '.docx')

        _thread.start_new_thread(os.system, (os.getcwd() + r"\\Final_Report\\" + testname + '_Graph.docx',))
        _thread.start_new_thread(os.system, (os.getcwd() + r"\\Final_Report\\" + testname + '.docx',))


    def OK_Pressed(self):
        # the first gui is for statistic table
        while self.listWidget.currentItem() is None:
            msgbox = QtWidgets.QMessageBox(self)
            msgbox.setText("No item is selected, please select an item.")
            msgbox.exec()
            return
        Screening_Report_WordGenerator.screening_Report_wordGenerator(self.listWidget.currentItem().text())
        # the statistic table report is generated when gui is generate the code are together
        RawData_Report_WordDoc.RawData_report_wordDoc(self.listWidget.currentItem().text())
        GraphPicture_Report_Combin.graphPicture_combin(self.listWidget.currentItem().text())
        testName = self.listWidget.currentItem().text()[0:-4]
        files = []
        files.append("merged" + testName + ".docx")
        files.append(testName + "Pre-StaticReport.docx")
        if os.path.exists(os.getcwd() + r"\\Report_Word\\" + testName + "Post-StaticReport.docx"):
            files.append(testName + "Post-StaticReport.docx")
        files.append(testName + "RawData.docx")
        print(files)
        self.combine_word_documents(files, testName)
        self.close()

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
