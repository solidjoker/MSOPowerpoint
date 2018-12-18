
# coding: utf-8

# In[1]:


import sys
import os
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QWidget, QInputDialog, QApplication, QFileDialog, QMainWindow, QMessageBox
from PyQt5.QtCore import pyqtSlot

from MSOPowerpoint import MSOPowerpoint


# In[3]:


class Example(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.setObjectName("MainWindow")
        self.resize(1000, 500)
        self.setWindowTitle('MSOPowerpointWindow')
        
    def initUI(self):
        self.centralwidget = QtWidgets.QWidget(self)
        
        self.btnSelect = QtWidgets.QPushButton('Select a pptx or input the pptx path', self)
        self.btnSelect.setGeometry(QtCore.QRect(20, 20, 400, 40))
        self.btnSelect.setObjectName('btnSelect')
        self.btnSelect.clicked.connect(self.pptxSelect)

        self.labelSelect = QtWidgets.QLabel(self)
        self.labelSelect.setGeometry(QtCore.QRect(20, 80, 180, 40))
        self.labelSelect.setObjectName('labelSelect')
        self.labelSelect.setText('Selected Powerpoint @')
        
        self.textSelect = QtWidgets.QLineEdit('',self)
        self.textSelect.setGeometry(QtCore.QRect(200, 80, 700, 40))
        self.textSelect.setObjectName('textSelect')
        
        self.btnExport = QtWidgets.QPushButton('Export pptx shapes info', self)
        self.btnExport.setGeometry(QtCore.QRect(20, 140, 400, 40))
        self.btnExport.setObjectName('btnExport')
        self.btnExport.clicked.connect(self.pptxExport)

        self.btnSelectExcel = QtWidgets.QPushButton('or Select a Excel', self)
        self.btnSelectExcel.setGeometry(QtCore.QRect(440, 140, 400, 40))
        self.btnSelectExcel.setObjectName('btnSelectExcel')
        self.btnSelectExcel.clicked.connect(self.excelSelect)
        
        self.labelExport = QtWidgets.QLabel(self)
        self.labelExport.setGeometry(QtCore.QRect(20, 200, 180, 40))
        self.labelExport.setObjectName('labelExport')
        self.labelExport.setText('Information excel @')
        
        self.textExport = QtWidgets.QLineEdit('',self)
        self.textExport.setGeometry(QtCore.QRect(200, 200, 700, 40))
        self.textExport.setObjectName('textExport')

        self.btnCreate = QtWidgets.QPushButton('Create pptx from Excel', self)
        self.btnCreate.setGeometry(QtCore.QRect(20, 260, 400, 40))
        self.btnCreate.setObjectName('btnCreate')
        self.btnCreate.clicked.connect(self.pptxCreate)
        
        self.show()
        

        
    @pyqtSlot()
    def pptxSelect(self):
        fileName, filetype = QFileDialog.getOpenFileName(self, 
                                                         caption='select a pptx',directory=os.path.curdir,filter='Powerpoint Files (*.pptx)')
        print(filetype)
        if os.path.exists(fileName):
            self.textSelect.setText(str(fileName))
        else:
            pass # 弹出警告

    @pyqtSlot()
    def pptxExport(self):
        if self.message(Text='analyse pptx?\nYes or No') == 16384:
            filename = self.textSelect.text()
            if os.path.exists(filename):
                MP = MSOPowerpoint(FileName=filename, WithWindow=1)
                MP.exportShapesInfo()
                self.textExport.setText(str(MP.exportPPTtoExcel()))
                self.message(Text='Done',YesNo=None)
            else:
                pass # 弹出警告

    @pyqtSlot()
    def excelSelect(self):
        fileName, filetype = QFileDialog.getOpenFileName(self, 
                                                         caption='select a pptx',directory=os.path.curdir,filter='Excel Files (*.xlsx)')
        print(filetype)
        if os.path.exists(fileName):
            self.textExport.setText(str(fileName))
        else:
            pass # 弹出警告
        
    @pyqtSlot()
    def pptxCreate(self):
        if self.message(Text='export pptx?\nYes or No') == 16384:
            filename = self.textExport.text()
            if os.path.exists(filename):
                MP = MSOPowerpoint(Blank=None)
                newfilename = MP.createPPT(filename)
                self.message(Text='please find the newfilename\n%s\n%s'%(os.path.dirname(newfilename),os.path.basename(newfilename)),YesNo=None)
            else:
                pass # 弹出警告
        
    @pyqtSlot()
    def message(self,Text=None,YesNo=True):
        if YesNo:
            return QMessageBox.question(self,'Message',Text,QMessageBox.StandardButtons(QMessageBox.Yes | QMessageBox.No))
        else:
            return QMessageBox.question(self,'Message',Text,QMessageBox.StandardButtons(QMessageBox.Ok))

        


# In[4]:


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())

