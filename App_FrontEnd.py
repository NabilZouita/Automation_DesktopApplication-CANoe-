# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'QtCreatorMainWindow.ui'
#
# Created by: PyQt4 UI code generator 4.11.4
#
# WARNING! All changes made in this file will be lost!

import sys
import subprocess
from PyQt4 import QtCore, QtGui
import HelloWorld
import App_BackEnd


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

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        
        MainWindow.setObjectName(_fromUtf8("MainWindow"))
        MainWindow.resize(400, 300)
        
        self.centralWidget = QtGui.QWidget(MainWindow)
        self.centralWidget.setObjectName(_fromUtf8("centralWidget"))
        
        self.verticalLayoutWidget = QtGui.QWidget(self.centralWidget)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(10, 10, 371, 231))
        self.verticalLayoutWidget.setObjectName(_fromUtf8("verticalLayoutWidget"))
        
        self.verticalLayout = QtGui.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setMargin(11)
        self.verticalLayout.setSpacing(6)
        self.verticalLayout.setObjectName(_fromUtf8("verticalLayout"))
        
        self.frame = QtGui.QFrame(self.verticalLayoutWidget)
        self.frame.setFrameShape(QtGui.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtGui.QFrame.Raised)
        self.frame.setObjectName(_fromUtf8("frame"))
        
        self.pushButton = QtGui.QPushButton(self.frame)
        self.pushButton.clicked.connect(self.call_scriptpy)
        self.pushButton.setGeometry(QtCore.QRect(30, 70, 75, 23))
        self.pushButton.setObjectName(_fromUtf8("pushButton"))
        self.verticalLayout.addWidget(self.frame)
        MainWindow.setCentralWidget(self.centralWidget)
        
        self.menuBar = QtGui.QMenuBar(MainWindow)
        self.menuBar.setGeometry(QtCore.QRect(0, 0, 400, 21))
        self.menuBar.setObjectName(_fromUtf8("menuBar"))
        MainWindow.setMenuBar(self.menuBar)
        
        self.mainToolBar = QtGui.QToolBar(MainWindow)
        self.mainToolBar.setObjectName(_fromUtf8("mainToolBar"))
        MainWindow.addToolBar(QtCore.Qt.TopToolBarArea, self.mainToolBar)
        
        self.statusBar = QtGui.QStatusBar(MainWindow)
        self.statusBar.setObjectName(_fromUtf8("statusBar"))
        MainWindow.setStatusBar(self.statusBar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def call_scriptpy(self):
        self.path = "Template_Test.xlsx"
        ##### Call Backend GUI Application ####
        self.var2 = App_BackEnd.ExcelPy()
        print self.var2.ParseFSheet(self.path)

        ##### Open HelloWorld.py #####
        #self.var = HelloWorld.callFunc()
        #print self.var.PrintFun()

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow", None))
        self.pushButton.setText(_translate("MainWindow", "Execute", None))


if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    MainWindow = QtGui.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

