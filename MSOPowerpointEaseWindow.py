
# coding: utf-8

# In[1]:


import sys
import os
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QWidget, QInputDialog, QApplication, QFileDialog, QMainWindow, QMessageBox
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtGui import QIcon

from MSOPowerpointEase import MSOPowerpointEase


# In[2]:


class Example(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.setObjectName("MainWindow")
        self.setWindowTitle('MSOPowerpointWindowEase')
        self.resize(1080, 540)
        self.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint |   # 使能最小化按钮
                            QtCore.Qt.WindowCloseButtonHint |      # 使能关闭按钮
                            QtCore.Qt.WindowStaysOnTopHint)        # 窗体总在最前端
        self.setFixedSize(self.width(), self.height())             # 固定窗体大小        
        self.directory = os.path.curdir
        self.initUI()
        try:
            self.setWindowIcon(QIcon('SJ.jpg'))
        except:
            print(self.directory)
            traceback.print_exc()
        
        
    def initUI(self):
        self.centralwidget = QtWidgets.QWidget(self)
        self.geometryTop = 20
        self.geometryIntervaly = 60
        
        self.btnSelect = QtWidgets.QPushButton('Select a pptx or input the pptx path', self)
        self.btnSelect.setGeometry(QtCore.QRect(20, self.geometryTop, 400, 40))
        self.btnSelect.setObjectName('btnSelect')
        self.btnSelect.clicked.connect(self.pptxSelect)
        
        self.checkBoxTable = QtWidgets.QCheckBox('save table as picture',self)
        self.checkBoxTable.setGeometry(QtCore.QRect(500, self.geometryTop, 300 ,20))
        self.checkBoxTable.setObjectName("checkBoxTable")
        self.checkBoxTable.setChecked(False)
        
        self.checkBoxChart = QtWidgets.QCheckBox('keep chart autosize',self)
        self.checkBoxChart.setGeometry(QtCore.QRect(500, self.geometryTop+20, 300 ,20))
        self.checkBoxChart.setObjectName("checkBoxChart")
        self.checkBoxChart.setChecked(False)
        
        self.geometryTop += self.geometryIntervaly
        self.labelSelect = QtWidgets.QLabel(self)
        self.labelSelect.setGeometry(QtCore.QRect(20, self.geometryTop, 200, 40))
        self.labelSelect.setObjectName('labelSelect')
        self.labelSelect.setText('Selected Powerpoint @')
        
        self.textSelect = QtWidgets.QLineEdit('',self)
        self.textSelect.setGeometry(QtCore.QRect(220, self.geometryTop, 800, 40))
        self.textSelect.setObjectName('textSelect')
        
        self.geometryTop += self.geometryIntervaly
        self.btnExport = QtWidgets.QPushButton('Export pptx shapes info', self)
        self.btnExport.setGeometry(QtCore.QRect(20, self.geometryTop, 400, 40))
        self.btnExport.setObjectName('btnExport')
        self.btnExport.clicked.connect(self.pptxExport)

        self.btnSelectDir = QtWidgets.QPushButton('or Select a Folder', self)
        self.btnSelectDir.setGeometry(QtCore.QRect(440, self.geometryTop, 400, 40))
        self.btnSelectDir.setObjectName('btnSelectDir')
        self.btnSelectDir.clicked.connect(self.dirSelect)
        
        self.geometryTop += self.geometryIntervaly
        self.labelDirname = QtWidgets.QLabel(self)
        self.labelDirname.setGeometry(QtCore.QRect(20, self.geometryTop, 200, 40))
        self.labelDirname.setObjectName('labelDirname')
        self.labelDirname.setText('Information folder @')
        
        self.textDirname = QtWidgets.QLineEdit('',self)
        self.textDirname.setGeometry(QtCore.QRect(220, self.geometryTop, 800, 40))
        self.textDirname.setObjectName('textDirname')
        
        self.geometryTop += self.geometryIntervaly
        self.labelExportExcel = QtWidgets.QLabel(self)
        self.labelExportExcel.setGeometry(QtCore.QRect(20, self.geometryTop, 200, 40))
        self.labelExportExcel.setObjectName('labelExportExcel')
        self.labelExportExcel.setText('Information excel @')
        
        self.textExportExcel = QtWidgets.QLineEdit('',self)
        self.textExportExcel.setGeometry(QtCore.QRect(220, self.geometryTop, 800, 40))
        self.textExportExcel.setObjectName('textExportExcel')
        
        self.geometryTop += self.geometryIntervaly
        self.labelTemplate = QtWidgets.QLabel(self)
        self.labelTemplate.setGeometry(QtCore.QRect(20, self.geometryTop, 200, 40))
        self.labelTemplate.setObjectName('labelTemplate')
        self.labelTemplate.setText('PPTX Template @')
        
        self.textTemplate = QtWidgets.QLineEdit('',self)
        self.textTemplate.setGeometry(QtCore.QRect(220, self.geometryTop, 800, 40))
        self.textTemplate.setObjectName('textExportExcel')

        self.geometryTop += self.geometryIntervaly
        self.btnCreate = QtWidgets.QPushButton('Create pptx from Excel', self)
        self.btnCreate.setGeometry(QtCore.QRect(20, self.geometryTop, 400, 40))
        self.btnCreate.setObjectName('btnCreate')
        self.btnCreate.clicked.connect(self.pptxCreate)
        
        self.show()

    @pyqtSlot()
    def pptxSelect(self):
        fileName, filetype = QFileDialog.getOpenFileName(self, 
                                                         caption='select a pptx',
                                                         directory=self.directory,
                                                         filter='Powerpoint Files (*.pptx)')
        print(filetype)
        print('checkBoxChart:%s'%self.checkBoxChart.checkState())
        if os.path.exists(fileName):
            self.textSelect.setText(str(fileName))
            self.directory = os.path.dirname(fileName)

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
                    self.directory = dirname
                else:
                    self.message(Text='There\'s no shapes in %s'%filename,YesNo=None)
                    
    @pyqtSlot()
    def dirSelect(self):                         
        dirname = QFileDialog.getExistingDirectory(self, 'select a folder', './')
        self.textDirname.setText(str(dirname))
        self.directory = dirname
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
            self.directory = dirname
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

        


# In[3]:


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())

