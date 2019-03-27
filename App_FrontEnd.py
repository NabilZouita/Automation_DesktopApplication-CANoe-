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
        MainWindow.resize(407, 300)
        MainWindow.setWindowIcon(QtGui.QIcon('python_logo.png'))
        
        self.centralWidget = QtGui.QWidget(MainWindow)
        self.centralWidget.setObjectName(_fromUtf8("centralWidget"))
        
        self.frame = QtGui.QFrame(self.centralWidget)
        self.frame.setGeometry(QtCore.QRect(20, 10, 381, 241))
        self.frame.setFrameShape(QtGui.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtGui.QFrame.Raised)
        self.frame.setObjectName(_fromUtf8("frame"))
        
        self.verticalLayoutWidget = QtGui.QWidget(self.frame)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(0, 0, 181, 131))
        self.verticalLayoutWidget.setObjectName(_fromUtf8("verticalLayoutWidget"))
        
        self.verticalLayout = QtGui.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setMargin(11)
        self.verticalLayout.setSpacing(6)
        self.verticalLayout.setObjectName(_fromUtf8("verticalLayout"))
        
        self.widget = QtGui.QWidget(self.verticalLayoutWidget)
        self.widget.setObjectName(_fromUtf8("widget"))
        
        self.pushButton = QtGui.QPushButton(self.widget)
        self.pushButton.setGeometry(QtCore.QRect(10, 5, 81, 23))
        self.pushButton.setObjectName(_fromUtf8("pushButton"))
        self.pushButton.clicked.connect(self.load_sheet)

        self.pushButton_2 = QtGui.QPushButton(self.widget)
        self.pushButton_2.setGeometry(QtCore.QRect(10, 30, 81, 23))
        self.pushButton_2.setObjectName(_fromUtf8("pushButton_2"))
        self.pushButton_2.clicked.connect(self.load_config)
             
        self.pushButton_3 = QtGui.QPushButton(self.widget)
        self.pushButton_3.setGeometry(QtCore.QRect(10, 55, 81, 23))
        self.pushButton_3.setObjectName(_fromUtf8("pushButton_3"))
        self.pushButton_3.setEnabled(False)
        self.pushButton_3.clicked.connect(self.generate_script)
        
        self.pushButton_4 = QtGui.QPushButton(self.widget)
        self.pushButton_4.setGeometry(QtCore.QRect(10, 80, 81, 23))
        self.pushButton_4.setObjectName(_fromUtf8("pushButton_4"))
        self.pushButton_4.setEnabled(False)
        self.pushButton_4.clicked.connect(self.call_scriptPy)
        
        self.verticalLayout.addWidget(self.widget)

        ########### Added Features in GUI ###########
        self.verticalLayoutWidget_2 = QtGui.QWidget(self.frame)
        self.verticalLayoutWidget_2.setGeometry(QtCore.QRect(190, 0, 181, 131))
        self.verticalLayoutWidget_2.setObjectName(_fromUtf8("verticalLayoutWidget_2"))
        
        self.verticalLayout_2 = QtGui.QVBoxLayout(self.verticalLayoutWidget_2)
        self.verticalLayout_2.setMargin(11)
        self.verticalLayout_2.setSpacing(6)
        self.verticalLayout_2.setObjectName(_fromUtf8("verticalLayout_2"))
        
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
        
        self.widget_2 = QtGui.QWidget(self.frame)
        self.widget_2.setGeometry(QtCore.QRect(190, 0, 179, 149))
        self.widget_2.setObjectName(_fromUtf8("widget_2"))
        
        self.radioButton = QtGui.QRadioButton(self.widget_2)
        self.radioButton.setGeometry(QtCore.QRect(20, 40, 82, 17))
        self.radioButton.setObjectName(_fromUtf8("radioButton"))
        self.radioButton.clicked.connect(self.test_passed)
        
        self.radioButton_2 = QtGui.QRadioButton(self.widget_2)
        self.radioButton_2.setGeometry(QtCore.QRect(20, 80, 82, 17))
        self.radioButton_2.setObjectName(_fromUtf8("radioButton_2"))
        self.radioButton_2.clicked.connect(self.test_notPassed)
        ############ The End ##############

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
        
        self.actionAbout = QtGui.QAction(MainWindow)
        self.actionAbout.setObjectName(_fromUtf8("actionAbout"))
        
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
        self.ExcelSheet = QtGui.QFileDialog.getOpenFileName(self,'Single File','*.xlsx')
        
        #This is for seeing the var type : non str
        #print self.ExcelSheet 
        
        """ Trying to change the type of var """
        #test = os.path.splitext(str(ExcelSheet))
        #print test
        
        self.InstVar = App_BackEnd.ExcelPy()
        self.EnvVar = self.InstVar.ParseFSheet(str(self.ExcelSheet))[0]
        self.ValVar = self.InstVar.ParseFSheet(str(self.ExcelSheet))[1]
        
        #print self.EnvVar
        #print self.ValVar
        
        self.message = ("Loaded successfully: %s \n"%os.path.split(str(self.ExcelSheet))[1])
        
        self.textEdit.clear()
        self.textEdit.setTextColor(QtGui.QColor("green"))
        self.textEdit.insertPlainText(self.message)

        #print len(self.ValVar)
        #return self.VarTemp

    def load_config(self):
        self.ConfigFile = QtGui.QFileDialog.getOpenFileName(self,'Single File','*.cfg')
        self.pushButton_3.setEnabled(True)
        self.pushButton_4.setEnabled(True)
        #print self.ConfigFile
        self.message = ("Loaded successfull: %s"%os.path.split(str(self.ConfigFile))[1])
        
        self.textEdit.clear()
        self.textEdit.setTextColor(QtGui.QColor("green"))
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

        self.inst = App_BackEnd.ExcelPy()
        self.inst.Script_Editor(self.ConfigFile,self.EnvVar,self.ValVar)

        self.message = ("Your script has been generated successfully:")
        
        self.textEdit.clear()
        self.textEdit.setTextColor(QtGui.QColor("green"))
        self.textEdit.insertPlainText(self.message)

    def call_scriptPy(self):     
        self.message = ("Your Test will start in minute! Please wait a little ^_^ ...")
        
        self.textEdit.clear()
        self.textEdit.setTextColor(QtGui.QColor("Blue"))
        self.textEdit.insertPlainText(self.message)
  
        self.PyScript = QtGui.QFileDialog.getOpenFileName(self,'Single File','*.py')
        subprocess.call(["python",str(self.PyScript)])

        self.message = ("Execution started ! Please check CANoe")

        self.textEdit.clear()
        self.textEdit.setTextColor(QtGui.QColor("Blue"))
        self.textEdit.insertPlainText(self.message)

    def close_application(self):
        choice = QtGui.QMessageBox.question(self, 'NOTICE',"Are you sure you want Exit?", 
                                            QtGui.QMessageBox.Yes | QtGui.QMessageBox.No)
        if choice == QtGui.QMessageBox.Yes:
            sys.exit()
        else:
            pass

    def test_passed(self):
        self.state = self.radioButton.isChecked()
        if self.state == True:
            self.inst = App_BackEnd.ExcelPy()
            self.inst.LogEditorPass(str(self.ExcelSheet))
        else:
            pass

    def test_notPassed(self):
        self.state = self.radioButton_2.isChecked()
        if self.state == True:
            self.inst = App_BackEnd.ExcelPy()
            self.inst.LogEditorNotPass(str(self.ExcelSheet))
        else:
            pass

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow", None))
        self.pushButton.setText(_translate("MainWindow", "Load", None))
        self.pushButton_2.setText(_translate("MainWindow", "Configuration", None))
        self.pushButton_3.setText(_translate("MainWindow", "Generate", None))
        self.pushButton_4.setText(_translate("MainWindow", "Execute", None))
        self.radioButton.setText(_translate("MainWindow","Test OK",None))
        self.radioButton_2.setText(_translate("MainWindow","Test Not OK",None))
        self.menuFile.setTitle(_translate("MainWindow", "File", None))
        self.menuHelp.setTitle(_translate("MainWindow", "Help", None))
        self.actionOpen_Sheet.setText(_translate("MainWindow", "Open Sheet", None))
        self.actionOpen_Config.setText(_translate("MainWindow", "Open Config", None))
        self.actionOpen_Script.setText(_translate("MainWindow", "Open Script", None))
        self.actionSave_As.setText(_translate("MainWindow", "Save As", None))
        self.actionQuit.setText(_translate("MainWindow", "Quit", None))
        self.actionAbout_this_App.setText(_translate("MainWindow", "About this App", None))
        self.actionAbout.setText(_translate("MainWindow", "About...", None))


if __name__ == "__main__":
    import sys
    app = QtGui.QApplication(sys.argv)
    MainWindow = QtGui.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

