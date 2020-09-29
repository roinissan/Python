import win32com.client


import Manager
import os
import sys

class MacroInjection:

    def __init__(self,file_path):
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        elif __file__:
            application_path = os.path.dirname(__file__)

        config_path = os.path.join(application_path, "macro.txt")
        self.macro_path = config_path
        self.file_path = file_path


    def injection(self,file_path):
        macro_path = self.macro_path
        excel_path = file_path
        try:
            with open(macro_path, "r") as myfile:
                macro = myfile.read()
        except Exception:
            print("could not find the file")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.Interactive = False
        workbook = excel.Workbooks.Open(Filename=excel_path)
        excelModule = workbook.VBProject.VBComponents.Add(1)
        excelModule.CodeModule.AddFromString(macro)
        excel.Application.Run('roi')
        x =os.path.splitext(file_path)[0]
        for sheet in workbook.Worksheets:
            for chartObject in sheet.ChartObjects():
                # print(sheet.Name + ':' + chartObject.Name)
                chartObject.Chart.Export( os.path.dirname(file_path)+ "/chart.png")
        excel.Workbooks(1).Close(SaveChanges=1)
        excel.Application.Quit()
        del excel


    def convert_to_xlsm_xlsx(self,file_path,xlsx_or_xlsm):
        if (xlsx_or_xlsm == 51):
            new_file_path = os.path.splitext(file_path)[0] + ".xlsx"
        else:
            new_file_path = os.path.splitext(file_path)[0] + ".xlsm"
        excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(file_path)
        excel.DisplayAlerts = False
        wb.DoNotPromptForConvert = True
        wb.CheckCompatibility = False
        wb.SaveAs(new_file_path, FileFormat=xlsx_or_xlsm, ConflictResolution=2)
        excel.Application.Quit()

    def killProcess(self):
        r = os.popen('tasklist /v').read().strip().split('\n')
        for i in range(len(r)):
            if("EXCEL" in r[i]):
                os.system("taskkill /f /im  EXCEL.EXE")
    def run(self):
        new_file_path = os.path.splitext(self.file_path)[0] + ".xlsm"
        self.convert_to_xlsm_xlsx(self.file_path,52)
        self.injection(new_file_path)
        self.convert_to_xlsm_xlsx(new_file_path,51)