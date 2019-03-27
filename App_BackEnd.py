import xlrd
import openpyxl
import os, sys
from xlrd import open_workbook, XLRDError
from openpyxl import Workbook
from openpyxl.styles import Alignment 
from Python_CANoe import CANoe


class ExcelPy:
    def __init__(self):
        #### First test case ####
        #print("Hello world !")
        pass

        #### Test File entered #####
    def TestFile(self,PathExc):
        try:
            self.book = xlrd.open_workbook(PathExc)
        except XLRDError:
            raise RuntimeError("You have entered a wrong File ! Please Check")
        else:
            raise RuntimeError("Thank you!")
        
    def ParseFSheet(self,PathExc):
        #self.temp = os.path.splitext(PathExc)
        self.list_Values = []
        self.list_EnVar = []

        self.book = xlrd.open_workbook(PathExc)
        self.sheet = self.book.sheet_by_index(0)
        
        self.List_EnVar = self.sheet.col_values(0)
        self.List_Values = self.sheet.col_values(1)
        
        return self.List_EnVar, self.List_Values 

    def ParseSSheet(self,PathExc,WSheetVal):
        self.LTest_EnVar = []
        self.LTest_Value = []
        
        self.book = xlrd.open_workbook(PathExc)
        self.sheet = self.book.sheet_by_index(WSheetVal)
        
        self.LTest_EnVar = self.sheet.col_values(0)
        self.LTest_Value = self.sheet.col_values(1)
        
        return self.LTest_Value, self.LTest_EnVar
        '''
        for i in range(1,self.sheet.nrows):
            self.cell_value = self.sheet.cell(i,1).value
            self.cell_type = self.sheet.cell_type(i,1)
            if(self.cell_type != 0):
                self.LTest_EnVar.append(self.sheet.cell(i,1).value)
                self.LTest_Value.append(int(self.sheet.cell(i,2).value))
            else:
                break '''         

    def LogEditorPass(self,PathExc,WSheetVal):
        self.wb = openpyxl.load_workbook(PathExc)
        
        self.sheet = self.wb.worksheets[WSheetVal]
        self.sheet.merge_cells('C1:D5')
        
        self.cell = self.sheet.cell(row=1, column=3)
        self.cell.value = 'Test OK'
        self.cell.alignment = Alignment(horizontal='center', vertical='center')

        self.wb.save(PathExc)

    def LogEditorNotPass(self,PathExc,WSheetVal):
        self.wb = openpyxl.load_workbook(PathExc)
        
        self.sheet = self.wb.worksheets[WSheetVal]
        self.sheet.merge_cells('C1:D5')
        
        self.cell = self.sheet.cell(row=1, column=3)
        self.cell.value = 'Test Not OK'
        self.cell.alignment = Alignment(horizontal='center', vertical='center')

        self.wb.save(PathExc)


class Py_CANoe:
    def __init__(self):
        pass
    # Checking path will be done in Python_CANoe.py
    def Script_Editor(self,ConfigFile,EnvVar,EnvVal):
        self.file = open("Script.py","w")
        
        self.file.write("from Python_CANoe import CANoe \n")
        self.file.write("var = CANoe() \n")
        self.file.write("var.open_simulation('%s') \n"%ConfigFile)
        self.file.write("var.start_Measurement() \n")
        
        for i in range(len(EnvVal)):
            self.file.write("var.set_EnvVar('%s',%s) \n"%(EnvVar[i],EnvVal[i]))
        
        self.file.close()

    def PyCan_Exec(self,path):
        self.root = CANoe()
        self.root.open_simulation(path) #CANoe function restriction are in package
        self.root.start_Measurement()
        self.root.set_EnvVar()
        

'''
PathExc = "Python_ExcelParse.xlsx"
inst = ExcelPy()
print inst.ParseFSheet(PathExc)
print inst.ParseSSheet(PathExc)
print inst.LogEditor(PathExc)

PathCAN = ""
PyCan = Py_CANoe()
print PyCan.PyCan_Exec(PathCAN)'''
