# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'QtCreatorMainWindow.ui'
#
# Created by: PyQt4 UI code generator 4.11.4
#
# WARNING! All changes made in this file will be lost!
import os, sys
import subprocess
import App_BackEnd
from PyQt4 import QtCore, QtGui

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

try:
    _encoding = QtGui.QApplication.UnicodeUTF8
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig)

class Ui_MainWindow(QtGui.QWidget):
    def setupUi(self, MainWindow):
    
        MainWindow.setObjectName(_fromUtf8("MainWindow"))
        MainWindow.resize(532, 377)
        
        self.centralWidget = QtGui.QWidget(MainWindow)
        self.centralWidget.setObjectName(_fromUtf8("centralWidget"))
        
        ########## MainFrame Layout Setting: Begin #############
        self.frame = QtGui.QFrame(self.centralWidget)
        
        self.frame.setGeometry(QtCore.QRect(20, 10, 491, 311))
        self.frame.setFrameShape(QtGui.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtGui.QFrame.Raised)
        self.frame.setObjectName(_fromUtf8("frame"))
        ########## MainFrame Layout setting: End ############

        ########## PushButtons Layout Design : Begin ##########
        self.verticalLayoutWidget = QtGui.QWidget(self.frame)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(0, 0, 270, 181))
        self.verticalLayoutWidget.setObjectName(_fromUtf8("verticalLayoutWidget"))
        
        self.verticalLayout = QtGui.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setMargin(11)
        self.verticalLayout.setSpacing(6)
        self.verticalLayout.setObjectName(_fromUtf8("verticalLayout"))
        
        self.widget = QtGui.QWidget(self.verticalLayoutWidget)

        ######### Size Policy Layout : Begin #########
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.widget.sizePolicy().hasHeightForWidth())
        ######### Size Policy Layout : End ##########
        
        self.widget.setSizePolicy(sizePolicy)
        self.widget.setObjectName(_fromUtf8("widget"))

        ######## PushButton1 Layout : Begin ###########    
        self.pushButton = QtGui.QPushButton(self.widget)
        
        self.pushButton.setGeometry(QtCore.QRect(2, 10, 119, 61))
    
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Minimum, QtGui.QSizePolicy.Fixed)
        
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton.sizePolicy().hasHeightForWidth())
        
        self.pushButton.setSizePolicy(sizePolicy)
        
        font = QtGui.QFont()
        
        font.setFamily(_fromUtf8("Arial"))
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        
        self.pushButton.setFont(font)
        self.pushButton.setDefault(True)   
        self.pushButton.setObjectName(_fromUtf8("pushButton"))
        self.pushButton.setStatusTip("Open & Load Excel Template")
        self.pushButton.clicked.connect(self.load_sheet)
        ######## PushButton1 Layout Design : End #########

        ######## PushButton2 Layout Desing : Begin #######
        self.pushButton_2 = QtGui.QPushButton(self.widget) 
        self.pushButton_2.setGeometry(QtCore.QRect(2, 90, 119, 61))
        
        font = QtGui.QFont()
        
        font.setFamily(_fromUtf8("Arial"))
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        
        self.pushButton_2.setFont(font)
        self.pushButton_2.setDefault(True)
        self.pushButton_2.setEnabled(False)
        
        self.pushButton_2.setObjectName(_fromUtf8("pushButton_2"))
        self.pushButton_2.setStatusTip("Choose & Load CANoe Configuration")
        self.pushButton_2.clicked.connect(self.load_config)
        ######## PushButton2 Layout Design : End ##########

        ######## PushButton3 Layout Design : Begin #########
        self.pushButton_3 = QtGui.QPushButton(self.widget)
        self.pushButton_3.setGeometry(QtCore.QRect(138, 10, 111, 61))

        font = QtGui.QFont()
        
        font.setFamily(_fromUtf8("Arial"))
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        
        self.pushButton_3.setFont(font)
        self.pushButton_3.setDefault(True)
        self.pushButton_3.setObjectName(_fromUtf8("pushButton_3"))
        self.pushButton_3.setEnabled(False)
        self.pushButton_3.setStatusTip("Generate & Save Script")
        self.pushButton_3.clicked.connect(self.generate_script)
        """ PushButton 3 Layout Design : End"""

        """ PushButton 4 Layout Design : Begin """
        self.pushButton_4 = QtGui.QPushButton(self.widget)
        
        self.pushButton_4.setGeometry(QtCore.QRect(138, 90, 111, 61))

        font = QtGui.QFont()
        
        font.setFamily(_fromUtf8("Arial"))
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        
        self.pushButton_4.setFont(font)
        self.pushButton_4.setDefault(True)
        self.pushButton_4.setObjectName(_fromUtf8("pushButton_4"))
        self.pushButton_4.setStatusTip("Execute choosen script")
        self.pushButton_4.setEnabled(False)
        self.pushButton_4.clicked.connect(self.call_scriptPy)
        """ PushButton 4 Layout Design : End """

        ########### Layout Two lines : Begin #############
        self.line = QtGui.QFrame(self.widget)
        self.line.setGeometry(QtCore.QRect(0, 80, 261, 20))
        self.line.setLineWidth(2)
        self.line.setFrameShadow(QtGui.QFrame.Sunken)
        self.line.setObjectName(_fromUtf8("line"))
        
        self.line_2 = QtGui.QFrame(self.widget)
        self.line_2.setGeometry(QtCore.QRect(120, 0, 20, 181))
        self.line_2.setLineWidth(2)
        self.line_2.setFrameShape(QtGui.QFrame.VLine)
        self.line_2.setFrameShadow(QtGui.QFrame.Sunken)
        self.line_2.setObjectName(_fromUtf8("line_2"))
        ########### Two lines Edit code : End #############        
        self.verticalLayout.addWidget(self.widget)
        ########### Design RadioButton : Begin ###########
        self.verticalLayoutWidget_2 = QtGui.QWidget(self.frame)
        
        self.verticalLayoutWidget_2.setGeometry(QtCore.QRect(300, 10, 181, 100))
        self.verticalLayoutWidget_2.setObjectName(_fromUtf8("verticalLayoutWidget_2"))
        
        self.verticalLayout_2 = QtGui.QVBoxLayout(self.verticalLayoutWidget_2)
        
        self.verticalLayout_2.setMargin(11)
        self.verticalLayout_2.setSpacing(6)
        self.verticalLayout_2.setObjectName(_fromUtf8("verticalLayout_2"))
        ############ RadioButtons Layout Code ###########
        self.widget_2 = QtGui.QWidget(self.verticalLayoutWidget_2)

        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Expanding)
        
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.widget_2.sizePolicy().hasHeightForWidth())

        self.widget_2.setSizePolicy(sizePolicy)
        self.widget_2.setObjectName(_fromUtf8("widget_2"))
        ######## Layout : RadioButton1 ##########
        self.radioButton = QtGui.QRadioButton(self.widget_2)
        
        self.radioButton.setGeometry(QtCore.QRect(2, 0, 82, 21))
        
        font = QtGui.QFont()
        
        font.setFamily(_fromUtf8("Arial"))
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        
        self.radioButton.setFont(font)
        self.radioButton.setObjectName(_fromUtf8("radioButton"))
        self.radioButton.setStatusTip("Conducted Test State : PASS !")
        self.radioButton.clicked.connect(self.test_passed)
        ######## The End : RadioButton1 Design Layout ##########
        
        ######## RadioButton2 Design Layout : Begin ########
        self.radioButton_2 = QtGui.QRadioButton(self.widget_2)
        
        self.radioButton_2.setGeometry(QtCore.QRect(2, 40, 111, 20))

        font = QtGui.QFont()
        
        font.setFamily(_fromUtf8("Arial"))
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        
        self.radioButton_2.setFont(font)

        self.radioButton_2.setObjectName(_fromUtf8("radioButton_2"))
        self.radioButton_2.setStatusTip("Conducted test state : FAIL !")
        self.radioButton_2.clicked.connect(self.test_notPassed)

        self.verticalLayout_2.addWidget(self.widget_2)
        ######## The End : RadioButton2 #########

        ############ The End : RadioButton Layout code ##############

        ####### Design Layout Text Edit : Begin #########
        self.horizontalLayoutWidget = QtGui.QWidget(self.frame)
        
        self.horizontalLayoutWidget.setGeometry(QtCore.QRect(10, 170, 470, 151))
        self.horizontalLayoutWidget.setObjectName(_fromUtf8("horizontalLayoutWidget"))
        
        self.horizontalLayout = QtGui.QHBoxLayout(self.horizontalLayoutWidget)
        
        self.horizontalLayout.setMargin(11)
        self.horizontalLayout.setSpacing(6)
        self.horizontalLayout.setObjectName(_fromUtf8("horizontalLayout"))
        
        self.textEdit = QtGui.QTextEdit(self.horizontalLayoutWidget)
        
        font = QtGui.QFont()
        
        font.setFamily(_fromUtf8("Arial"))
        font.setPointSize(16)
        font.setBold(True)
        font.setWeight(75)
        
        self.textEdit.setFont(font)
        self.textEdit.setObjectName(_fromUtf8("textEdit"))
        
        self.horizontalLayout.addWidget(self.textEdit)
        ###### Text Edit Design Layout : End ######"" 

        """ Deleted Text Edit Layout
        self.horizontalLayoutWidget = QtGui.QWidget(self.frame)
        
        self.horizontalLayoutWidget.setGeometry(QtCore.QRect(0, 140, 371, 91))
        self.horizontalLayoutWidget.setObjectName(_fromUtf8("horizontalLayoutWidget"))
        
        self.horizontalLayout = QtGui.QHBoxLayout(self.horizontalLayoutWidget)
        
        self.horizontalLayout.setMargin(11)
        self.horizontalLayout.setSpacing(6)
        self.horizontalLayout.setObjectName(_fromUtf8("horizontalLayout"))
        
        self.widget_3 = QtGui.QWidget(self.horizontalLayoutWidget)
        
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Minimum, QtGui.QSizePolicy.Minimum)
        
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.widget_3.sizePolicy().hasHeightForWidth())
        
        self.widget_3.setSizePolicy(sizePolicy)
        self.widget_3.setObjectName(_fromUtf8("widget_3"))
        
        self.textEdit = QtGui.QTextEdit(self.widget_3)
        
        self.textEdit.setGeometry(QtCore.QRect(0, 0, 337, 59))
        self.textEdit.setObjectName(_fromUtf8("textEdit"))
        #self.textEdit.clicked.connect(self.TextEditor)
        
        self.horizontalLayout.addWidget(self.widget_3)
        Here Ends the deleted part """
        
        ########## ComboBox + LCD + ProgressBar Layout : Begin ###################
        self.verticalLayoutWidget_3 = QtGui.QWidget(self.frame)
        
        self.verticalLayoutWidget_3.setGeometry(QtCore.QRect(300, 80, 181, 101))
        self.verticalLayoutWidget_3.setObjectName(_fromUtf8("verticalLayoutWidget_3"))
        
        self.verticalLayout_3 = QtGui.QVBoxLayout(self.verticalLayoutWidget_3)
        
        self.verticalLayout_3.setMargin(11)
        self.verticalLayout_3.setSpacing(6)
        self.verticalLayout_3.setObjectName(_fromUtf8("verticalLayout_3"))
        
        self.widget_4 = QtGui.QWidget(self.verticalLayoutWidget_3)
        
        self.widget_4.setObjectName(_fromUtf8("widget_4"))

        ############## ComboBox Layout Code ##################
        self.comboBox = QtGui.QComboBox(self.widget_4)
        
        self.comboBox.setGeometry(QtCore.QRect(2, 10, 101, 31))

        font = QtGui.QFont()
        
        font.setFamily(_fromUtf8("Arial"))
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        
        self.comboBox.setFont(font)

        self.comboBox.setObjectName(_fromUtf8("comboBox"))
        self.comboBox.setStatusTip("Complete list of Tests in the Template")
        self.comboBox.currentIndexChanged.connect(self.SecondSheet)  
        """
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        self.comboBox.addItem(_fromUtf8(""))
        """
        ######## The End ##############

        ######## Begin QLCDNumber Widget #######
        """ SpinBox Widget Deleted 
        self.spinBox = QtGui.QSpinBox(self.widget_4)
        
        self.spinBox.setGeometry(QtCore.QRect(110, 0.75, 42, 22))
        self.spinBox.setObjectName(_fromUtf8("spinBox"))
        self.spinBox.setMinimum(0)
        #self.spinBox.valueChanged.connect(self.ValueChange)"""

        self.progressBar = QtGui.QProgressBar(self.widget_4)
        
        self.progressBar.setGeometry(QtCore.QRect(2, 50, 161, 21))
        
        font = QtGui.QFont()
        
        font.setFamily(_fromUtf8("Arial"))
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        
        self.progressBar.setFont(font)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName(_fromUtf8("progressBar"))
        self.progressBar.setStatusTip("Process state progressing")
        
        self.lcdNumber = QtGui.QLCDNumber(self.widget_4)
        
        self.lcdNumber.setGeometry(QtCore.QRect(115, 10, 41, 31))
        self.lcdNumber.setFrameShape(QtGui.QFrame.WinPanel)
        self.lcdNumber.setDigitCount(3)
        self.lcdNumber.setObjectName(_fromUtf8("lcdNumber"))
        self.lcdNumber.setStatusTip("The number of Tests available to execute")

        ######## End : QLCDNumber Widget ############
        self.verticalLayout_3.addWidget(self.widget_4)


        MainWindow.setCentralWidget(self.centralWidget)
        
        self.menuBar = QtGui.QMenuBar(MainWindow)
        
        self.menuBar.setGeometry(QtCore.QRect(0, 0, 407, 21))
        self.menuBar.setObjectName(_fromUtf8("menuBar"))
        
        self.menuFile = QtGui.QMenu(self.menuBar)
        
        self.menuFile.setObjectName(_fromUtf8("menuFile"))
        
        self.menuHelp = QtGui.QMenu(self.menuBar)
        
        self.menuHelp.setObjectName(_fromUtf8("menuHelp"))
        
        MainWindow.setMenuBar(self.menuBar)
        
        self.mainToolBar = QtGui.QToolBar(MainWindow)
        
        self.mainToolBar.setObjectName(_fromUtf8("mainToolBar"))
        
        MainWindow.addToolBar(QtCore.Qt.TopToolBarArea, self.mainToolBar)
        
        self.statusBar = QtGui.QStatusBar(MainWindow)
        
        self.statusBar.setObjectName(_fromUtf8("statusBar"))
        
        MainWindow.setStatusBar(self.statusBar)
        
        self.actionOpen_Sheet = QtGui.QAction(MainWindow)
        
        self.actionOpen_Sheet.setObjectName(_fromUtf8("actionOpen_Sheet"))
        self.actionQuit.setShortcut("Ctrl+O")
        self.actionQuit.setStatusTip("Open & Choose Template !")
        
        self.actionOpen_Config = QtGui.QAction(MainWindow)
        
        self.actionOpen_Config.setObjectName(_fromUtf8("actionOpen_Config"))
        
        self.actionOpen_Script = QtGui.QAction(MainWindow)
        
        self.actionOpen_Script.setObjectName(_fromUtf8("actionOpen_Script"))
        
        self.actionSave_As = QtGui.QAction(MainWindow)
        
        self.actionSave_As.setObjectName(_fromUtf8("actionSave_As"))
        
        self.actionQuit = QtGui.QAction(MainWindow)
        
        self.actionQuit.setObjectName(_fromUtf8("actionQuit"))
        self.actionQuit.setShortcut("Ctrl+Q")
        self.actionQuit.setStatusTip("Close the Application !")
        self.actionQuit.triggered.connect(self.close_application)
        
        self.actionAbout_this_App = QtGui.QAction(MainWindow)
        
        self.actionAbout_this_App.setObjectName(_fromUtf8("actionAbout_this_App"))
        self.actionAbout_this_App.setStatusTip("App Summary !")
        self.actionAbout_this_App.triggered.connect(self.app_summary)
        
        self.actionAbout = QtGui.QAction(MainWindow)
        
        self.actionAbout.setObjectName(_fromUtf8("actionAbout"))
        self.actionAbout.setStatusTip("App documentation !")
        self.actionAbout.triggered.connect(self.app_documentation)
        
        self.menuFile.addAction(self.actionOpen_Sheet)
        self.menuFile.addAction(self.actionOpen_Config)
        self.menuFile.addAction(self.actionOpen_Script)
        
        self.menuFile.addSeparator()
        
        self.menuFile.addAction(self.actionSave_As)
        self.menuFile.addAction(self.actionQuit)
        
        self.menuHelp.addAction(self.actionAbout_this_App)
        self.menuHelp.addAction(self.actionAbout)
        
        self.menuBar.addAction(self.menuFile.menuAction())
        self.menuBar.addAction(self.menuHelp.menuAction())

        self.retranslateUi(MainWindow)
        
        QtCore.QMetaObject.connectSlotsByName(MainWindow)


    def load_sheet(self):
        try:
            self.ExcelSheet = QtGui.QFileDialog.getOpenFileName(self,'Single File','*.xlsx')
            
            self.InstVar = App_BackEnd.ExcelPy()
            self.EnvVar = self.InstVar.ParseFSheet(str(self.ExcelSheet))[0]
            self.ValVar = self.InstVar.ParseFSheet(str(self.ExcelSheet))[1]
            
            #print self.EnvVar
            #print self.ValVar

            progess_Var = self.ProgressBar()
            
            self.message = ("Loaded successfully: %s \n"%os.path.split(str(self.ExcelSheet))[1])
            
            self.textEdit.clear()
            self.textEdit.setTextColor(QtGui.QColor("green"))
            self.textEdit.insertPlainText(self.message)

            comboBox_Var = self.ValueChange()

            self.instVar = App_BackEnd.ExcelPy()
            self.lcdNumber.display(self.instVar.GetNumbSheets(str(self.ExcelSheet))-1)

            self.pushButton_2.setEnabled(True)

        except IOError:
            self.message = "You did not choose a file \n Please check again!"
            
            self.textEdit.clear()
            self.textEdit.setTextColor(QtGui.QColor("red"))
            self.textEdit.insertPlainText(self.message)
        except:
            raise

        #This is for seeing the var type : non str
        #print self.ExcelSheet 
        
        """ Trying to change the type of var """
        #test = os.path.splitext(str(ExcelSheet))
        #print test   
        #print len(self.ValVar)
        #return self.VarTemp


    def load_config(self):
        self.ConfigFile = QtGui.QFileDialog.getOpenFileName(self,'Single File','*.cfg')
        
        self.FileName = os.path.split(str(self.ConfigFile))[1]
        
        if (self.FileName != ""):
            #print self.ConfigFile
            
            self.message = ("Loaded successfully: %s"%os.path.split(str(self.ConfigFile))[1])
            
            self.textEdit.clear()

            progess_Var = self.ProgressBar()

            self.textEdit.setTextColor(QtGui.QColor("green"))
            self.textEdit.insertPlainText(self.message)
            
            self.pushButton_3.setEnabled(True)
        else:
            self.message = ("You did not choose a correct Config File \n Please Check !")
            
            self.textEdit.clear()
            self.textEdit.setTextColor(QtGui.QColor("red"))
            self.textEdit.insertPlainText(self.message)

        """ Caling PyCan_Exec from here """
        #self.VarTemp = App_BackEnd.Py_CANoe()
        #print self.VarTemp.PyCan_Exec(self.ConfigFile)


    def generate_script(self):
        """
        self.file = open("Script.py","w")
        self.file.write("from Python_CANoe import CANoe \n")
        self.file.write("var = CANoe() \n")
        self.file.write("var.open_simulation('%s') \n"%self.ConfigFile)
        self.file.write("var.start_Measurement() \n")
        
        for i in range(len(self.ValVar)):
            self.file.write("var.set_EnvVar('%s',%s) \n"%(self.EnvVar[i],self.ValVar[i]))
        
        self.file.close()"""
        self.inst = App_BackEnd.Py_CANoe()
        self.inst.Script_Editor(self.ConfigFile,self.EnvVar,self.ValVar)
        try:
            self.message = ("Script generated successfully.")
            self.textEdit.clear()
            self.textEdit.setTextColor(QtGui.QColor("green"))

            progess_Var = self.ProgressBar()
            self.textEdit.insertPlainText(self.message)

            name = QtGui.QFileDialog.getSaveFileName(self,'Generate & Save !')
            file = open(name,'w')
            file.close()

            self.counter = 0

            self.pushButton_4.setEnabled(True)

        except IOError:
            self.message = "You did not save your file \n File will execute automatically!"
            
            self.textEdit.clear()
            self.textEdit.setTextColor(QtGui.QColor("red"))
            self.textEdit.insertPlainText(self.message)


    def call_scriptPy(self,**keywords):
        try:
            self.message = ("This will take few seconds ! Please wait ^_^ ...")
            self.textEdit.clear()
            self.textEdit.setTextColor(QtGui.QColor("Blue"))
            self.textEdit.insertPlainText(self.message)
            if (self.counter == 0):
                self.PyScript = QtGui.QFileDialog.getOpenFileName(self,'Single File', '*.py')
                subprocess.call(["python",str(self.PyScript)])
            else:
                subprocess.call(["python","Script.py"])
                progess_Var = self.ProgressBar()
        except:
            self.message = "No script has been choosen \n Please proceed again!"
            
            self.textEdit.clear()
            self.textEdit.setTextColor(QtGui.QColor("red"))
            self.textEdit.insertPlainText(self.message)

        self.message = ("Execution started ! Please check CANoe")
        self.textEdit.clear()
        self.textEdit.setTextColor(QtGui.QColor("Blue"))
        self.textEdit.insertPlainText(self.message)


        '''    
        self.PyScript = QtGui.QFileDialog.getOpenFileName(self,'Single File','*.py')
        subprocess.call(["python",str(self.PyScript)])'''

        '''
        if (callBool == True):
            self.PyScript = QtGui.QFileDialog.getOpenFileName(self,'Single File','*.py')
            subprocess.call(["python",str(self.PyScript)])
        else:
            self.PyScript = QtGui.QFileDialog.getOpenFileName(self,'Single File','*.py')
            subprocess.call(["python",str(self.PyScript)])
        '''


    def test_passed(self):
        self.state = self.radioButton.isChecked()
        if self.state == True:
            self.inst = App_BackEnd.ExcelPy()
            self.inst.LogEditorPass(str(self.ExcelSheet),self.varTemp)
        else:
            pass


    def test_notPassed(self):
        self.state = self.radioButton_2.isChecked()
        if self.state == True:
            self.inst = App_BackEnd.ExcelPy()
            self.inst.LogEditorNotPass(str(self.ExcelSheet),self.varTemp)
        else:
            pass


    def ValueChange(self):
        #self.message = ("current value:"+str(self.spinBox.value()))

        """ Trying to get number of sheets automatically """
        self.inst = App_BackEnd.ExcelPy()
        TempRes = self.inst.GetNumbSheets(str(self.ExcelSheet))
        print TempRes

        #Previous Value before change is "self.spinBox.value()"
        for i in range(0,TempRes):
            self.comboBox.addItem(_fromUtf8(""))
            self.comboBox.setItemText(i, _translate("MainWindow", "Test %d"%i, None))


    def SecondSheet(self,index):
        self.varTemp = index
        print ("Current index:",index)
        
        self.InstVar = App_BackEnd.ExcelPy()
        
        self.envVar = self.InstVar.ParseSSheet((str(self.ExcelSheet)),(index))[0]
        self.valVar = self.InstVar.ParseSSheet((str(self.ExcelSheet)),(index))[1]
        
        print self.envVar
        print self.valVar

        self.instVar = App_BackEnd.Py_CANoe()
        self.instVar.SScript_Editor(self.envVar,self.valVar)

        self.counter = 1


    def ProgressBar(self):
        self.completed = 0

        while self.completed < 100:
            self.completed += 0.0001
            self.progressBar.setValue(self.completed)


    def close_application(self):
        choice = QtGui.QMessageBox.question(self, 'NOTICE',"Are you sure ?", 
                                            QtGui.QMessageBox.Yes | QtGui.QMessageBox.No)
        if choice == QtGui.QMessageBox.Yes:
            sys.exit()
        else:
            pass


    def app_summary(self):
        message_box = QtGui.QMessageBox()
        
        message_box.setIcon(message_box.Information)
        message_box.setWindowTitle("About this Application")
        message_box.setWindowIcon(QtGui.QIcon('python_logo.png'))
        message_box.setText("PFE2019_AltranTunisie_IHM")
        message_box.setInformativeText("Automating CANoe process \nThis was done as End of Studies Project")
        #message_box.setDetailedText("PFE2019_AltranTunisie_IHM")

        #message_box.setStandardButtons(QtGui.QMessageBox.Ok)
        #message_box.setDefaultButton(QtGui.QMessageBox.Ok)
        #message_box.setEscapeButton(QtGui.QMessageBox.Ok)

        message_box.exec_()

        '''
        MainWindow = QtGui.QDialog()
        ui = Ui_MainWindow()
        MainWindow.exec_()
        '''
        '''
        self.textedit = QtGui.QTextEdit()
        MainWindow.setCentralWidget(self.textedit)'''

        '''
        MainWindow = QtGui.QDialog()
        ui = Ui_MainWindow()
        with open("Test.txt","r") as f:
            for line in f:
                print line.strip()

        MainWindow.exec_()'''

        '''
        MainWindow = QtGui.QDialog()
        ui = Ui_MainWindow()
        #ui = QtGui.QPlainTextEdit()
        text = open("Test.txt").read()
        ui.setPlainText(text)
        with open("Test.txt","r") as f:
            for line in f:
                print line.strip()


        MainWindow.show()'''

    def app_documentation(self):
        TextFile = QtGui.QDialog()

        text = open("test.txt","r").read()

        TextFile.exec_()


    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow", None))
        
        self.pushButton.setText(_translate("MainWindow", "Template Excel", None))
        
        self.pushButton_2.setText(_translate("MainWindow", "Configuration", None))
        
        self.pushButton_3.setText(_translate("MainWindow", "Generate", None))
        
        self.pushButton_4.setText(_translate("MainWindow", "Execute", None))
        
        self.radioButton.setText(_translate("MainWindow","Test OK",None))
        
        self.radioButton_2.setText(_translate("MainWindow","Test Not OK",None))
        """
        self.comboBox.setItemText(0, _translate("MainWindow", "Test 1", None))
        self.comboBox.setItemText(1, _translate("MainWindow", "Test 2", None))
        self.comboBox.setItemText(2, _translate("MainWindow", "Test 3", None))
        self.comboBox.setItemText(3, _translate("MainWindow", "Test 4", None))
        self.comboBox.setItemText(4, _translate("MainWindow", "Test 5", None))
        self.comboBox.setItemText(5, _translate("MainWindow", "Test 6", None))
        self.comboBox.setItemText(6, _translate("MainWindow", "Test 7", None))
        self.comboBox.setItemText(7, _translate("MainWindow", "Test 8", None))
        self.comboBox.setItemText(8, _translate("MainWindow", "Test 9", None))
        self.comboBox.setItemText(9, _translate("MainWindow", "Test 10", None))
        self.comboBox.setItemText(10, _translate("MainWindow", "Test 11", None))
        self.comboBox.setItemText(11, _translate("MainWindow", "Test 12", None))
        self.comboBox.setItemText(12, _translate("MainWindow", "Test 13", None))
        self.comboBox.setItemText(13, _translate("MainWindow", "Test 14", None))
        self.comboBox.setItemText(14, _translate("MainWindow", "Test 15", None))
        self.comboBox.setItemText(15, _translate("MainWindow", "Test 16", None))
        self.comboBox.setItemText(16, _translate("MainWindow", "Test 17", None))
        self.comboBox.setItemText(17, _translate("MainWindow", "Test 18", None))
        self.comboBox.setItemText(18, _translate("MainWindow", "Test 19", None))
        self.comboBox.setItemText(19, _translate("MainWindow", "Test 20", None))
        self.comboBox.setItemText(20, _translate("MainWindow", "Test 21", None))
        self.comboBox.setItemText(21, _translate("MainWindow", "Test 22", None))
        self.comboBox.setItemText(22, _translate("MainWindow", "Test 23", None))
        self.comboBox.setItemText(23, _translate("MainWindow", "Test 24", None))
        self.comboBox.setItemText(24, _translate("MainWindow", "Test 25", None))
        self.comboBox.setItemText(25, _translate("MainWindow", "Test 26", None))
        self.comboBox.setItemText(26, _translate("MainWindow", "Test 27", None))
        self.comboBox.setItemText(27, _translate("MainWindow", "Test 28", None))
        self.comboBox.setItemText(28, _translate("MainWindow", "Test 29", None))
        self.comboBox.setItemText(29, _translate("MainWindow", "Test 30", None))
        self.comboBox.setItemText(30, _translate("MainWindow", "Test 31", None))
        self.comboBox.setItemText(31, _translate("MainWindow", "Test 32", None))
        self.comboBox.setItemText(32, _translate("MainWindow", "Test 33", None))
        self.comboBox.setItemText(33, _translate("MainWindow", "Test 34", None))
        self.comboBox.setItemText(34, _translate("MainWindow", "Test 35", None))
        """
        self.menuFile.setTitle(_translate("MainWindow", "File", None))
        
        self.menuHelp.setTitle(_translate("MainWindow", "Help", None))
        
        self.actionOpen_Sheet.setText(_translate("MainWindow", "Open Sheet", None))
        
        self.actionOpen_Config.setText(_translate("MainWindow", "Open Config", None))
        
        self.actionOpen_Script.setText(_translate("MainWindow", "Open Script", None))
        
        self.actionSave_As.setText(_translate("MainWindow", "Save As", None))
        
        self.actionQuit.setText(_translate("MainWindow", "Quit", None))
        
        self.actionAbout_this_App.setText(_translate("MainWindow", "About this App", None))
        
        self.actionAbout.setText(_translate("MainWindow", "Documentation", None))


if __name__ == "__main__":
    import sys
    app = QtGui.QApplication(sys.argv)
    MainWindow = QtGui.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.setWindowIcon(QtGui.QIcon('python_logo.png'))
    MainWindow.setWindowTitle("PFE2019_AltranTunisie_IHM")
    MainWindow.show()
    sys.exit(app.exec_())
