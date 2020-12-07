import os
import pandas as pd
import sys
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog
from PyQt5.uic import loadUi
from Screening_System_PyQt5 import RGui_TemplateFile_add_withFunction

pd.options.display.max_columns = 999
pd.options.display.max_rows = 999
pd.set_option("display.precision", 6)


class templateFile_Open_WithFunction(QDialog):
    def __init__(self):
        super(templateFile_Open_WithFunction, self).__init__()
        loadUi("RGui_TemplateFile_Open.ui", self)
        self.populate_list_widget()
        self.lineEdit.textChanged.connect(self.search_In_List)
        self.listWidget.itemClicked.connect(self.item_Clicked)
        self.pushButton_2.clicked.connect(self.OK_Pressed)

    def item_Clicked(self):
        search = self.listWidget.currentItem().text()
        self.lineEdit.setText(search)

    def OK_Pressed(self):
        ui = RGui_TemplateFile_add_withFunction.Lift2Coding()
        while self.listWidget.currentItem() is None:
            msgbox = QtWidgets.QMessageBox(self)
            msgbox.setText("No item is selected, please select an item.")
            msgbox.exec()
            return
        ui.getTestNumber(self.listWidget.currentItem().text())
        self.close()
        ui.show()
        ui.exec_()

    def search_In_List(self):
        search = self.lineEdit.text()
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
        # hightlight the item when list size is 1
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