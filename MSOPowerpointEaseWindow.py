
# coding: utf-8

# In[3]:


import sys
import os
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QWidget, QInputDialog, QApplication, QFileDialog, QMainWindow, QMessageBox
from PyQt5.QtCore import pyqtSlot

from MSOPowerpointEase import MSOPowerpointEase


# In[5]:


class Example(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.setObjectName("MainWindow")
        self.resize(1080, 540)
        self.setWindowTitle('MSOPowerpointWindowEase')
        
    def initUI(self):
        self.centralwidget = QtWidgets.QWidget(self)
        
        self.btnSelect = QtWidgets.QPushButton('Select a pptx or input the pptx path', self)
        self.btnSelect.setGeometry(QtCore.QRect(20, 20, 400, 40))
        self.btnSelect.setObjectName('btnSelect')
        self.btnSelect.clicked.connect(self.pptxSelect)
        
        self.checkBoxTable = QtWidgets.QCheckBox('save table as picture',self)
        self.checkBoxTable.setGeometry(QtCore.QRect(500, 20, 200 ,20))
        self.checkBoxTable.setObjectName("checkBoxTable")
        self.checkBoxTable.setChecked(False)
        
        self.checkBoxChart = QtWidgets.QCheckBox('keep chart autosize',self)
        self.checkBoxChart.setGeometry(QtCore.QRect(500, 40, 200 ,20))
        self.checkBoxChart.setObjectName("checkBoxChart")
        self.checkBoxChart.setChecked(False)
        
        self.labelSelect = QtWidgets.QLabel(self)
        self.labelSelect.setGeometry(QtCore.QRect(20, 80, 180, 40))
        self.labelSelect.setObjectName('labelSelect')
        self.labelSelect.setText('Selected Powerpoint @')
        
        self.textSelect = QtWidgets.QLineEdit('',self)
        self.textSelect.setGeometry(QtCore.QRect(200, 80, 800, 40))
        self.textSelect.setObjectName('textSelect')
        
        self.btnExport = QtWidgets.QPushButton('Export pptx shapes info', self)
        self.btnExport.setGeometry(QtCore.QRect(20, 140, 400, 40))
        self.btnExport.setObjectName('btnExport')
        self.btnExport.clicked.connect(self.pptxExport)

        self.btnSelectDir = QtWidgets.QPushButton('or Select a Folder', self)
        self.btnSelectDir.setGeometry(QtCore.QRect(440, 140, 400, 40))
        self.btnSelectDir.setObjectName('btnSelectDir')
        self.btnSelectDir.clicked.connect(self.dirSelect)
        
        self.labelDirname = QtWidgets.QLabel(self)
        self.labelDirname.setGeometry(QtCore.QRect(20, 200, 180, 40))
        self.labelDirname.setObjectName('labelDirname')
        self.labelDirname.setText('Information folder @')
        
        self.textDirname = QtWidgets.QLineEdit('',self)
        self.textDirname.setGeometry(QtCore.QRect(200, 200, 800, 40))
        self.textDirname.setObjectName('textDirname')
        
        self.labelExportExcel = QtWidgets.QLabel(self)
        self.labelExportExcel.setGeometry(QtCore.QRect(20, 260, 180, 40))
        self.labelExportExcel.setObjectName('labelExportExcel')
        self.labelExportExcel.setText('Information excel @')
        
        self.textExportExcel = QtWidgets.QLineEdit('',self)
        self.textExportExcel.setGeometry(QtCore.QRect(200, 260, 800, 40))
        self.textExportExcel.setObjectName('textExportExcel')
        
        self.labelTemplate = QtWidgets.QLabel(self)
        self.labelTemplate.setGeometry(QtCore.QRect(20, 320, 180, 40))
        self.labelTemplate.setObjectName('labelTemplate')
        self.labelTemplate.setText('PPTX Template @')
        
        self.textTemplate = QtWidgets.QLineEdit('',self)
        self.textTemplate.setGeometry(QtCore.QRect(200, 320, 800, 40))
        self.textTemplate.setObjectName('textExportExcel')

        self.btnCreate = QtWidgets.QPushButton('Create pptx from Excel', self)
        self.btnCreate.setGeometry(QtCore.QRect(20, 380, 400, 40))
        self.btnCreate.setObjectName('btnCreate')
        self.btnCreate.clicked.connect(self.pptxCreate)
        
        self.show()


    @pyqtSlot()
    def pptxSelect(self):
        fileName, filetype = QFileDialog.getOpenFileName(self, 
                                                         caption='select a pptx',directory=os.path.curdir,filter='Powerpoint Files (*.pptx)')
        print(filetype)
        print('checkBoxChart:%s'%self.checkBoxChart.checkState())
        if os.path.exists(fileName):
            self.textSelect.setText(str(fileName))

    @pyqtSlot()
    def pptxExport(self):
        if self.message(Text='analyse pptx?\nYes or No') == 16384:
            filename = self.textSelect.text()
            if os.path.exists(filename):
                MPE = MSOPowerpointEase(Blank=False)
                dirname = MPE.linkPPTtoExcel(FileName=filename,keepChartListObjects=self.checkBoxChart.checkState())
                if dirname:
                    self.textDirname.setText(str(dirname))
                    self.textExportExcel.setText('powerpoint_data.xlsx')
                    self.textTemplate.setText('powerpoint.pptx')
                    self.message(Text='Done!',YesNo=None)
                else:
                    self.message(Text='There\'s no shapes in %s'%filename,YesNo=None)
                    
    @pyqtSlot()
    def dirSelect(self):                         
        dirname = QFileDialog.getExistingDirectory(self, 'select a folder', './')
        self.textDirname.setText(str(dirname))
        excelFile = 'powerpoint_data.xlsx'
        if not os.path.exists(os.path.join(dirname,excelFile)):
            self.message(Text="There's no '%s' in folder:%s"%(excelFile,dirname),YesNo=None)
        else:
            self.textExportExcel.setText(excelFile)
        pptFile = 'powerpoint.pptx'
        if not os.path.exists(os.path.join(dirname,pptFile)):
            self.message(Text="There's no '%s' in folder:%s"%(pptFile,dirname),YesNo=None)
        else:
            self.textTemplate.setText(str(pptFile))        
        
                      
    @pyqtSlot()
    def pptxCreate(self):
        if self.message(Text='export pptx?\nYes or No') == 16384:
            dirname = self.textDirname.text()
            MPE = MSOPowerpointEase(Blank=None)
            pptFile = MPE.linkExceltoPPT(dirname,saveTableAsPicture=self.checkBoxTable.checkState(),
                                         keepChartListObjects=self.checkBoxChart.checkState())                
            self.message(Text='please find the newfilename\n%s\n%s'%(os.path.dirname(pptFile),os.path.basename(pptFile)),YesNo=None)
        
    @pyqtSlot()
    def message(self,Text=None,YesNo=True):
        if YesNo:
            return QMessageBox.question(self,'Message',Text,QMessageBox.StandardButtons(QMessageBox.Yes | QMessageBox.No))
        else:
            return QMessageBox.question(self,'Message',Text,QMessageBox.StandardButtons(QMessageBox.Ok))

        


# In[6]:


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())

