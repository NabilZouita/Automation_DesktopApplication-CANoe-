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
        self.book = xlrd.open_workbook(PathExc)
        self.sheet = self.book.sheet_by_index(0)
        self.list_Values = []
        self.list_EnVar = []
        self.List_EnVar = self.sheet.col_values(0)
        self.List_Values = self.sheet.col_values(1)
        return self.List_EnVar, self.List_Values 

    def ParseSSheet(self,PathExc):
        self.book = xlrd.open_workbook(PathExc)
        self.sheet = self.book.sheet_by_index(1)
        self.LTest_EnVar = []
        self.LTest_Value = []
        for i in range(1,self.sheet.nrows):
            self.cell_value = self.sheet.cell(i,1).value
            self.cell_type = self.sheet.cell_type(i,1)
            if(self.cell_type != 0):
                self.LTest_EnVar.append(self.sheet.cell(i,1).value)
                self.LTest_Value.append(int(self.sheet.cell(i,2).value))
            else:
                break

        return self.LTest_Value, self.LTest_EnVar 

    def LogEditor(self,PathExc):
        self.wb = openpyxl.load_workbook(PathExc)
        self.sheet = self.wb.worksheets[1]
        self.sheet.merge_cells('E2:F6')
        self.cell = self.sheet.cell(row=2, column=5)
        self.cell.value = 'Test OK'
        self.cell.alignment = Alignment(horizontal='center', vertical='center')

        self.wb.save(PathExc)

class Py_CANoe:
    def __init__(self):
        pass
    # Checking path will be done in Python_CANoe.py
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
