import os
from PyQt5 import QtWidgets, uic
import sys
import pandas as pd
from Screening_System_PyQt5 import RGui_Data_Converting, RGui_Print, RGui_Anaylsis_searchWIthFunctions, \
    RGui_Screening_DataCollection, RGui_Data_OpenWithFunctions, RGui_TemplateFile_add_withFunction, \
    RGui_templateFile_Open_withFunction, RGui_Data_add_withFunctions, RGui_TemplateFile_Dupplicate_withFunction, \
    RGui_Report_Open_WithFunction, RGui_Report_FrontPage_WithFunction, RGui_BK_PRECISION_setting_withFuction, \
    RGui_SetPath_withFunctions, RGui_FastData_select_withFuntion

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
        self.actionBK_Power_Bank_Config.triggered.connect(self.new_bk_setting_clicked)
        self.actionFast_Data_Graph.triggered.connect(self.fast_data_select)

    def fast_data_select(self):
        ui = RGui_FastData_select_withFuntion.FastDataWithFunction()
        ui.show()
        ui.exec_()

    def new_bk_setting_clicked(self):
        ui = RGui_BK_PRECISION_setting_withFuction.BKPrecisionSettingWithFunction()
        ui.show()
        ui.exec_()

    def New_Ananylsis_Seach_Clicked(self):
        ui = RGui_Anaylsis_searchWIthFunctions.AnalysisSearchWithFunction()
        ui.show()
        ui.exec_()

    def New_SetPath_Clicked(self):
        ui = RGui_SetPath_withFunctions.SetPathWithFunction()
        ui.show()
        ui.exec_()

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
        # print(exctype, value, traceback)
        # error_log = pd.read_csv(os.path.dirname(sys.argv[0]) + r"\\Error_log.txt", sep='\t')
        sys.excepthook(exctype, value, traceback)
        sys.exit(1)


    sys.excepthook = exception_hook
    # this is fun

    app.exec_()
