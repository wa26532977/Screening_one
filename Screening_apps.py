from PyQt5 import QtWidgets, uic
import sys
import pandas as pd
# import dcload
# import glob
# import os
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
from Screening_System_PyQt5 import RGui_Report_FrontPage_WithFunction
# from Screening_System_PyQt5 import RGui_Report_Open_WithFunction_2
# from Screening_System_PyQt5 import RGui_Report_RawData_WithFunction


pd.options.display.max_columns = 999
pd.options.display.max_rows = 999
pd.set_option("display.precision", 6)


class Screening_app(QtWidgets.QMainWindow):

    def __init__(self):
        super(Screening_app, self).__init__()
        uic.loadUi('RGui_Screening_mainWindow.ui', self)
        self.showMaximized()
        self.actionAdd_New_Tempate_File.triggered.connect(self.New_TemplateFile_Add_Clicked)
        self.actionOpen_Template_File.triggered.connect(self.New_TemplateFile_Open_Clicked)
        self.actionDupplicate_a_Template_File.triggered.connect(self.New_TemplateFile_Dupplicate_Clicked)
        self.actionAdd_New_Date_File.triggered.connect(self.New_Data_Add_Clicked)
        self.actionOpen_Date_File.triggered.connect(self.New_Data_Open_Clicked)
        self.actionConverting_Final_Screening_Data_to_excel_Format.triggered.connect(self.New_Data_Coverting_Clicked)
        self.actionPrint.triggered.connect(self.New_Print_Clicked)
        self.actionSet_Path.triggered.connect(self.New_SetPath_Clicked)
        self.actionSample_Data_Viewing.triggered.connect(self.New_Ananylsis_Seach_Clicked)
        self.actionView_All_Report_And_Graph.triggered.connect(self.Selected_Report_Open_Clicked)
        self.actionPrint_All_Peport_And_Graph.triggered.connect(self.Selected_Report_Open_Clicked_2)
        # need work, this will print all the report and graph
        # self.actionPrint_All_Peport_And_Graph.triggered.connect(self.Selected_Report_Open_Clicked)
        # Just for now, later need to change under data.
        self.actionPrinter_Setup.triggered.connect(self.New_Screening_DataCollection_Clicked)

    def New_Screening_DataCollection_Clicked(self):
        data_collection = QtWidgets.QDialog()
        ui = RGui_Screening_DataCollection.Ui_Dialog()
        ui.setupUi(data_collection)
        data_collection.show()
        data_collection.exec_()

    def New_Ananylsis_Seach_Clicked(self):
        ui = RGui_Anaylsis_searchWIthFunctions.AnalysisSearchWithFunction()
        ui.show()
        ui.exec_()

    def New_SetPath_Clicked(self):
        set_path = QtWidgets.QDialog()
        ui = RGui_SetPath.Ui_Dialog()
        ui.setupUi(set_path)
        set_path.show()
        set_path.exec_()

    def New_Print_Clicked(self):
        gui_print = QtWidgets.QDialog()
        ui = RGui_Print.Ui_Dialog()
        ui.setupUi(gui_print)
        gui_print.show()
        gui_print.exec_()

    def New_Data_Coverting_Clicked(self):
        data_converting = QtWidgets.QDialog()
        ui = RGui_Data_Converting.Ui_Dialog()
        ui.setupUi(data_converting)
        data_converting.show()
        data_converting.exec_()

    def New_FrontPage_Clicked(self):
        ui = RGui_Report_FrontPage_WithFunction.Report_FrontPage_WithFunction()
        ui.show()
        ui.exec_()

    def Selected_Report_Open_Clicked(self):
        ui = RGui_Report_Open_WithFunction.Report_Open_WithFunctions()
        ui.show()
        ui.exec_()

    def Selected_Report_Open_Clicked_2(self):
        ui = RGui_Report_Open_WithFunction.Report_Open_WithFunctions()
        ui.show_report_toggle()
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