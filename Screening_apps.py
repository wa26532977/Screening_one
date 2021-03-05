from PyQt5 import QtWidgets
import sys
from Screening_System_PyQt5 import RGui_Screening_mainWindow
import pandas as pd
import dcload
import glob
import os

from Screening_System_PyQt5 import RGui_Data_Converting
from Screening_System_PyQt5 import RGui_Print
from Screening_System_PyQt5 import RGui_SetPath
from Screening_System_PyQt5 import RGui_Anaylsis_searchWIthFunctions
from Screening_System_PyQt5 import RGui_Screening_DataCollection

from Screening_System_PyQt5 import RGui_Data_OpenWithFunctions
from Screening_System_PyQt5 import RGui_TemplateFile_add_withFunction
from Screening_System_PyQt5 import RGui_templateFile_Open_withFunction
from Screening_System_PyQt5 import RGui_Data_add_withFunctions
from Screening_System_PyQt5 import RGui_TemplateFile_Dupplicate_withFunction
from Screening_System_PyQt5 import RGui_Report_Open_WithFunction
from Screening_System_PyQt5 import RGui_Report_Open_WithFunction_2
from Screening_System_PyQt5 import RGui_Report_RawData_WithFunction
from Screening_System_PyQt5 import RGui_Report_FrontPage_WithFunction
from Screening_System_PyQt5 import RGui_Graph_Pre_WithFunction
from Screening_System_PyQt5 import RGui_Report_StatisticsTable_WithFunctions

pd.options.display.max_columns = 999
pd.options.display.max_rows = 999
pd.set_option("display.precision", 6)


class Screening_app(RGui_Screening_mainWindow.Ui_ScreeningSys_mainWindow, QtWidgets.QMainWindow):

    def __init__(self):
        super(Screening_app, self).__init__()
        self.setupUi(self)
        self.showMaximized()
        self.actionAdd_New_Tempate_File.triggered.connect(self.New_TemplateFile_Add_Clicked)
        self.actionOpen_Template_File.triggered.connect(self.New_TemplateFile_Open_Clicked)
        self.actionDupplicate_a_Template_File.triggered.connect(self.New_TemplateFile_Dupplicate_Clicked)
        self.actionAdd_New_Date_File.triggered.connect(self.New_Data_Add_Clicked)
        self.actionOpen_Date_File.triggered.connect(self.New_Data_Open_Clicked)
        self.actionConverting_Final_Screening_Data_to_excel_Format.triggered.connect(self.New_Data_Coverting_Clicked)
        self.actionPrint.triggered.connect(self.New_Print_Clicked)
        self.actionSet_Path.triggered.connect(self.New_SetPath_Clicked)
        self.actionSample_Search.triggered.connect(self.New_Ananylsis_Seach_Clicked)
        #self.actionPost_Graph.triggered.connect(self.New_Graph_Post_Clicked)
        #self.actionPre_Graph_2.triggered.connect(self.New_Graph_Pre_Clicked)
        #self.actionStatistic_Table.triggered.connect(self.New_StatisticTable_Clicked)
        #self.actionRaw_Data_Table.triggered.connect(self.New_RawData_Clicked)
        self.actionView_All_Report_And_Graph.triggered.connect(self.Selected_Report_Open_Clicked)
        self.actionPrint_All_Peport_And_Graph.triggered.connect(self.Selected_Report_Open_Clicked_2)
        # need work, this will print all the report and graph
        #self.actionPrint_All_Peport_And_Graph.triggered.connect(self.Selected_Report_Open_Clicked)
        #Just for now, later need to change under data.
        self.actionPrinter_Setup.triggered.connect(self.New_Screening_DataCollection_Clicked)
        #Just For Testing

    def New_Screening_DataCollection_Clicked(self):
        DataCollection = QtWidgets.QDialog()
        ui = RGui_Screening_DataCollection.Ui_Dialog()
        ui.setupUi(DataCollection)
        DataCollection.show()
        DataCollection.exec_()

    def New_Ananylsis_Seach_Clicked(self):
        ui = RGui_Anaylsis_searchWIthFunctions.AnalysisSearchWithFunction()
        ui.show()
        ui.exec_()

    def New_SetPath_Clicked(self):
        SetPath = QtWidgets.QDialog()
        ui = RGui_SetPath.Ui_Dialog()
        ui.setupUi(SetPath)
        SetPath.show()
        SetPath.exec_()

    def New_Print_Clicked(self):
        Gui_print = QtWidgets.QDialog()
        ui = RGui_Print.Ui_Dialog()
        ui.setupUi(Gui_print)
        Gui_print.show()
        Gui_print.exec_()

    def New_Data_Coverting_Clicked(self):
        Data_converting = QtWidgets.QDialog()
        ui = RGui_Data_Converting.Ui_Dialog()
        ui.setupUi(Data_converting)
        Data_converting.show()
        Data_converting.exec_()

    def New_FrontPage_Clicked(self):
        ui = RGui_Report_FrontPage_WithFunction.Report_FrontPage_WithFunction()
        ui.show()
        ui.exec_()

    def Selected_Report_Open_Clicked(self):
        ui = RGui_Report_Open_WithFunction.Report_Open_WithFunctions()
        ui.show()
        ui.exec_()

    def Selected_Report_Open_Clicked_2(self):
        ui = RGui_Report_Open_WithFunction_2.Report_Open_WithFunctions()
        ui.show()
        ui.exec_()

    def New_Data_Open_Clicked(self):
        ui = RGui_Data_OpenWithFunctions.Data_Open_WithFunctions()
        ui.show()
        ui.exec_()

    def New_TemplateFile_Add_Clicked(self):
        ui = RGui_TemplateFile_add_withFunction.Lift2Coding()
        ui.show()
        ui.exec_()

    def New_TemplateFile_Open_Clicked(self):
        ui = RGui_templateFile_Open_withFunction.templateFile_Open_WithFunction()
        ui.show()
        ui.exec_()

    def New_TemplateFile_Dupplicate_Clicked(self):
        ui = RGui_TemplateFile_Dupplicate_withFunction.templateFile_Dupplicate_WithFunction()
        ui.show()
        ui.exec_()

    def New_Data_Add_Clicked(self):
        ui = RGui_Data_add_withFunctions.Data_add_WithFunctions()
        ui.show()
        ui.exec_()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    qt_app = Screening_app()
    qt_app.show()
    sys._excepthook = sys.excepthook

    def exception_hook(exctype, value, traceback):
        print(exctype, value, traceback)
        sys.excepthook(exctype, value, traceback)
        sys.exit(1)

    sys.excepthook = exception_hook
    # this is fun

    app.exec_()