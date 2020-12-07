import os
from PyQt5.QtWidgets import QDialog
from PyQt5.uic import loadUi
from Screening_System_PyQt5 import RGui_Screening_DataCollection_WithFunctions
import pandas as pd
import sys
pd.options.display.max_columns = 999
pd.options.display.max_rows = 999
pd.set_option("display.precision", 6)

class Data_add_WithFunctions(QDialog):
    def __init__(self):
        super(Data_add_WithFunctions, self).__init__()
        loadUi("RGui_Data_add.ui", self)
        self.populate_list_widget()
        self.lineEdit.textChanged.connect(self.search_In_List)
        self.lineEdit.returnPressed.connect(self.return_pressed)
        self.listWidget.itemClicked.connect(self.item_Clicked)
        self.pushButton_2.clicked.connect(self.OK_Pressed)
        self.pushButton.clicked.connect(self.Cancel_Pressed)

    def return_pressed(self):
        self.OK_Pressed()

    def Cancel_Pressed(self):
        self.close()

    def OK_Pressed(self):
        if self.listWidget.currentItem() is None:
            print("NO item is selected, please select an item!")
            return
        ui = RGui_Screening_DataCollection_WithFunctions.Screening_DataCollection_WithFunction()
        #created the excel data frame in screening Data(new)
        #sent the selected name to RGui_Screening_DataCollection_WithFunctions so I can load the template
        ui.getTestNumber(self.listWidget.currentItem().text())
        self.close()
        ui.show()
        ui.exec_()

    def item_Clicked(self):
        search = self.listWidget.currentItem().text()
        self.lineEdit.setText(search)
        self.lineEdit_2.setText(search)

    def search_In_List(self):
        search = self.lineEdit.text()
        self.lineEdit.setText(search)
        self.listWidget.clear()
        #dir_path = os.path.dirname(os.path.realpath(__file__))
        dir_path = os.path.dirname(sys.argv[0])
        path_data = dir_path + r"\\Screening_Template"
        if search == "":
            self.populate_list_widget()
        else:
            for r, d, f in os.walk(path_data):
                for file in f:
                    if search in file:
                        self.listWidget.addItem(file)
        #hightlight the item when list size is 1
        if self.listWidget.count() == 1:
            self.listWidget.setCurrentRow(0)
            print(self.listWidget.currentItem().text())

    def populate_list_widget(self):
        self.listWidget.clear()
        #dir_path = os.path.dirname(os.path.realpath(__file__))
        dir_path = os.path.dirname(sys.argv[0])
        path_data = dir_path + r"\\Screening_Template"

        for r, d, f in os.walk(path_data):
            for file in f:
                if ".txt" in file:
                    self.listWidget.addItem(file)