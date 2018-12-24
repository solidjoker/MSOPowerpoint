
# coding: utf-8

# # MSOPowerpoint
# update: 2018-12-15
# 
# <li>利用ppt原先的板式，对于除Table、SmartShape之外的图形，替换Text、Data</li>
# <li>扫描获取ZOrder并排序,Table放最后，其他根据上、左顺序排</li>
# <li>需要处理Excel之间的变化</li>
# 
# <li>图表中的listobject需要去除</li>
# 
# <li>对于损坏的Chart的修复</li>
# 

# In[1]:


import os, time, datetime, pprint, traceback, tempfile, copy
from collections import defaultdict
from PIL import Image
import pandas as pd
import numpy as np
from itertools import chain

import win32com
from win32com.client import Dispatch,Constants

import MSOPowerpointConfig
import MSOPowerpointFunc
from MSOPowerpointBase import MSOPowerpointBase
from MSOPowerpointElement import MSOPowerpointElement


# In[2]:


class MSOPowerpointEase(MSOPowerpointBase,MSOPowerpointElement):
    '''
    MSOPowerPointTool for read and create pptx
    '''
    def __init__(self,FileName=None,ReadOnly=0,Untitled=0,WithWindow=-1,GetDpi=0,Visible=-1,Blank=1):
        '''
        initial pptApp
        reference to self.openPPT
        '''
        self.win32com = win32com
        self.appName = 'PowerPoint.Application'
        self.pptApp = self.win32com.client.Dispatch(self.appName)
        self.pptApp.Visible = 1
        self.pptApp.DisplayAlerts = 1 # set DisplayAlerts to ppAlertsNone
        self.pptPres = self.pptApp.Presentations
        self.msoShapeType = MSOPowerpointConfig.msoShapeType
        self.msoAutoShapeType = MSOPowerpointConfig.msoAutoShapeType
        self.msoTableStype = MSOPowerpointConfig.msoTableStyle
        self.xlChartType = MSOPowerpointConfig.xlChartType
        self.encoding = MSOPowerpointConfig.customInfo['encoding']
        self.tries = MSOPowerpointConfig.customInfo['tries'] # global classs tries
        self.seconds = MSOPowerpointConfig.customInfo['seconds'] # global time sleep seconds
        print('APP %s initialized!'%self.pptApp.Name)
        self.openPPT(FileName,ReadOnly,Untitled,WithWindow,GetDpi,Blank)
    def _funcGetDpi(self):
        '''
        getDpi
        '''
        dpi = 72
        ci = 0.05 # 误差
        w = h = 100 # 初始宽度
        im = Image.new('1',(w,h),'white')
        tempf = tempfile.mkstemp('.jpg')[1]
        im.save(tempf,dpi=(dpi,dpi))
        im.close()
        todel = False
        if not self.pptSlides.Count:
            todel = True
            self.addSlide()
        self.addPicture(tempf,AdjustDpi=False)
        ration = self.pptShape.Width / w
        if  not (1 - ci < ration < 1 + ci):
            dpi *= ration
        self.delShape()
        if todel:
            self.delSlide()
        return dpi
    def _help(self,obj=None):
        # 要改的！！！
        if not obj:
            '''
            默认Fill
            '''
            todel = False
            if not self.pptSlides.Count:
                todel = True
                self.addSlide()
            self.addShape()
            help(self.pptShape.Fill)
            self.delShape()
            if todel:
                self.delSlide()
        else:
            help(obj)
            
    def linkPPTtoExcel(self,FileName=None,keepChartListObjects=None):
        '''
        linkPPTtoExcel
        '''
        try:
            information = 'linkPPTtoExcel >>> Powerpoint:%s is exporting!'%(os.path.basename(FileName))
            self.slowprint(information)
            print('>'*40)
            self.openPPT(FileName=FileName)
            dirname,pptFile,excelFile,trackFile = self.linkPPTinitFiles(FileName=FileName)
            self.savePPT(newFileName=pptFile)
            self.ungroupShape() # 去除组合
            self.sortShapes(TableForward=True) # 排列顺序
            self.shapeList = []
            attrs = ['ZOrderPosition','Visible','Left','Top','Height','Width']
            keys = ['SlideNumber','ShapeType'] + attrs + ['ReplaceText','Text','ReplaceData','Data','ReplaceShapeSource','ShapeSource']
            # Shape.ZOrder(),msoBringToFront 0  # 到最后,msoSendToBack 1 # 到最前,msoBringForward 2,msoSendBackward 3
            print('linkPPTtoExcel, 1st Round for Summary start >>>')
            # 1rd Round Summary
            for SlideNumber in range(1,self.presInfo['SlidesCount']+1):
                pptSlide = self.pptSlides.Item(SlideNumber)
                self.linkPPTSlide(self.shapeList,attrs,keys,pptSlide)
                print('linkPPTtoExcel, slide %s scanned!'%SlideNumber)
                
            if not self.shapeList:
                print('There\'s no shapes in %s'%pptFile)
                return False
            df = pd.DataFrame(data=self.shapeList,index=range(len(self.shapeList)))
            df = df[keys]
            try:
                df.replace(np.nan,None)
            except:
                traceback.print_exc(limit=1)
            self.exportDfsToExcel([df],['Summary'],excelFile)
            # Init Excel
            ExcelApp, wkbs, wkb, shts, shtSum = self.initExcel(filename=excelFile,shtname='Summary')
            # 2nd Round Table
            print('linkPPTtoExcel, 2nd Round for Table start >>>')        
            self.linkPPTTable(df, wkb, shts, shtSum)        
            # 3rd Round Chart
            print('linkPPTtoExcel, 3rd Round for Chart start >>>')        
            self.linkPPTChart(df, wkb, shts, shtSum, keepChartListObjects)
            #wkbs.Close()
            #ExcelApp.Quit()  
            ExcelApp.Visible = 1
            os.startfile(dirname)
            information = 'linkPPTtoExcel finished >>> Please find the Excel:%s in the dir:%s!'%(os.path.basename(excelFile),os.path.dirname(excelFile))
            self.savePPT()
            self.slowprint(information)
            print('<'*40)
            return dirname
        except:
            traceback.print_exc()
            return False
    def linkPPTinitFiles(self,FileName):
        '''
        linkPPT, init files
        '''
        timestamp = time.strftime("%Y%m%d%H%M%S", time.localtime())
        dirname = '%s_%s'%(FileName[:FileName.rfind('.')],timestamp)
        os.mkdir(dirname)
        pptFile = os.path.join(dirname,'powerpoint.pptx')
        excelFile = os.path.join(dirname,'powerpoint_data.xlsx')
        trackFile =  os.path.join(dirname,'powerpoint_track.txt')
        return dirname,pptFile,excelFile,trackFile
    def linkPPTSlide(self,shapeList,attrs,keys,pptSlide):
        '''
        linkPPTSlide, in linkPPTtoExcel, 1st Round for Summary
        '''
        try:
            defaultdic = {}
            for key in keys:
                defaultdic.setdefault(key,None)
            for Shape in pptSlide.Shapes:
                shapedic = defaultdic.copy()
                shapedic['SlideNumber'] = pptSlide.SlideNumber
                for attr in attrs:
                    shapedic[attr] = int(getattr(Shape,attr))
                if Shape.HasTable:
                    shapedic['ShapeType'] = 'Table'
                    shapedic['ReplaceShapeSource'] = True
                    shapedic['ShapeSource'] = 'S_%s_%s_%s'%(shapedic['SlideNumber'],'Table',shapedic['ZOrderPosition'])
                elif Shape.HasChart:
                    if not Shape.Chart.ChartData.IsLinked:
                        shapedic['ShapeType'] = 'Chart'
                        shapedic['ReplaceData'] = True
                        shapedic['Data'] = 'S_%s_%s_%s'%(shapedic['SlideNumber'],'Chart',shapedic['ZOrderPosition'])     
                    else:
                        shapedic['ShapeType'] = 'Chart'
                elif Shape.HasTextFrame:
                    if Shape.TextFrame.TextRange.Text:
                        shapedic['ShapeType'] = 'Text'
                        shapedic['ReplaceText'] = True
                        shapedic['Text'] = Shape.TextFrame.TextRange.Text
                    else:
                        shapedic['ShapeType'] = 'Shape'                    
                else:
                    shapedic['ShapeType'] = 'Shape'
                shapeList.append(shapedic.copy())
        except:
            traceback.print_exc() 
    def linkPPTTable(self,df, wkb, shts, shtSum):
        '''
        linkPPTSlide, in linkPPTtoExcel, 2nd Round for Table
        '''
        
        j = list(df.columns).index('ShapeSource')
        for i in df.index.tolist():
            if df['ReplaceShapeSource'][i] == True:
                linkdic = df.transpose()[i].to_dict()    
                sht = shts.Add(After=wkb.Activesheet)
                sht.Name = linkdic['ShapeSource']
                self.selectSlide(linkdic['SlideNumber'])
                self.pptShape = self.pptSlide.Shapes(linkdic['ZOrderPosition'])
                self.setShape()
                self.pptShape.Copy()
                time.sleep(self.seconds/2)
                sht.Cells(1,1).Activate()
                sht.Paste()
                shtAddress = '%s!%s'%(sht.Name,sht.UsedRange.Address)
                shtSum.Activate()
                shtSum.Hyperlinks.Add(Anchor=shtSum.Cells(i+2,j+2),Address='',SubAddress=shtAddress,TextToDisplay=shtAddress)
                wkb.Save()
    def linkPPTChart(self,df, wkb, shts, shtSum, keepChartListObjects=None):
        '''
        linkPPTSlide, in linkPPTtoExcel, 3rd Round for Chart
        '''
        j = list(df.columns).index('Data')
        for i in df.index.tolist():
            if df['ReplaceData'][i] == True:
                linkdic = df.transpose()[i].to_dict()    
                sht = shts.Add(After=wkb.Activesheet)
                sht.Name = linkdic['Data']
                self.selectSlide(linkdic['SlideNumber'])
                self.pptShape = self.pptSlide.Shapes(linkdic['ZOrderPosition'])
                self.setShape()
                tries = self.tries
                while tries:
                    try:
                        pptWorkbook,pptWorksheet = self.initChartData() # 初始化
                        break
                    except:
                        tries -= 1
                        time.sleep(1)
                if not keepChartListObjects:
                    listObjectsCount = pptWorksheet.ListObjects.Count
                    if listObjectsCount:
                        for k in range(listObjectsCount):
                            pptWorksheet.ListObjects(1).Unlist()
                time.sleep(self.seconds/2)
                pptWorksheet.Cells.Copy(sht.Cells(1))
                pptWorkbook.Close()
                shtAddress = '%s!%s'%(sht.Name,sht.UsedRange.Address)
                shtSum.Activate()
                shtSum.Hyperlinks.Add(Anchor=shtSum.Cells(i+2,j+2),Address='',SubAddress=shtAddress,TextToDisplay=shtAddress)
                wkb.Save()   

    def linkExceltoPPT(self,dirname=None,saveTableAsPicture=None,keepChartListObjects=None):
        '''
        linkExceltoPPT
        '''
        pptFile,excelFile,trackFile = self.linkExcelinitFiles(dirname=dirname)
        if not os.path.exists(pptFile) or not os.path.exists(excelFile):
            print('%s/nor/n%s not found!'%(pptFile,excelFile))
            return False
        
        information = 'linkExceltoPPT >>> PPT:%s is creating!'%(os.path.basename(pptFile))
        self.slowprint(information)
        print('>'*40)
        
        self.openPPT(pptFile)
        df = pd.read_excel(excelFile,sheet_name='Summary').replace(np.nan,'')
        ExcelApp, wkbs, wkb, shts, shtSum = self.initExcel(filename=excelFile,shtname='Summary')
        attrs = ['ZOrderPosition','Visible','Left','Top','Height','Width']        
        linkelements = {'ReplaceText':'Text','ReplaceData':'Data','ReplaceShapeSource':'ShapeSource'}
        for i in df.index.tolist():
            for linkelement in linkelements:
                if df[linkelement][i] == True:
                    credic = df.transpose()[i].to_dict()
                    self.linkExcelElement(i,linkelements[linkelement],credic,attrs,wkb,shts,trackFile,
                                          saveTableAsPicture,keepChartListObjects)
        shtSum.Activate()
        wkbs.Parent.CutCopyMode = False
        wkb.Close(SaveChanges=False)
        wkbs.Close()
        ExcelApp.Quit() 
        if os.path.exists(trackFile):
            os.startfile(trackFile)
            return False
        information = 'linkExceltoPPT finished >>> please find the PPT:%s in the dir:%s!'%(os.path.basename(pptFile),dirname)
        self.savePPT()
        self.slowprint(information)
        print('<'*40)
        os.startfile(dirname)
        return pptFile
    def linkExcelElement(self,i,linktype,credic,attrs,wkb,shts,trackFile,saveTableAsPicture=None,keepChartListObjects=None):
        self.selectSlide(credic['SlideNumber'])
        self.pptShape = self.pptSlide.Shapes(credic['ZOrderPosition'])
        slideShapesCounts = self.pptSlide.Shapes.Count
        self.setShape()
        try:
            if linktype == 'Text':
                if self.pptShape.HasTextFrame:
                    self.pptShape.TextFrame.TextRange.Text = credic[linktype]
                    msg = 'index %s done!' % i
                    print(msg)
                else:
                    msg = 'index %s failed!' % i
                    print(msg)
                    self.linkExcelTrack(msg,trackFile)

            if linktype == 'Data':
                if self.pptShape.HasChart:
                    pptWorkbook,pptWorksheet = self.initChartData() # 初始化
                    shtdata = shts(credic[linktype])
                    shtdata.Activate()
                    shtdata.Cells.Copy()
                    time.sleep(self.seconds/2)
                    shtdata.Cells(1).PasteSpecial(Paste=-4163) # 复制数值
                    wkb.Parent.CutCopyMode = False
                    shtdata.Cells.Copy(pptWorksheet.Cells(1))
                    print(keepChartListObjects)
                    print(pptWorksheet.Name,pptWorksheet.Cells(1).CurrentRegion.Address)
                    if keepChartListObjects:
                        self.pptChart.SetSourceData(Source='=%s!%s'%(pptWorksheet.Name,pptWorksheet.Cells(1).CurrentRegion.Address))
                        self.pptChart.Refresh()
                    pptWorkbook.Close()
                    msg = 'index %s done!' % i
                    print(msg)
                else:
                    msg = 'index %s failed!' % i
                    print(msg)
                    self.linkExcelTrack(msg,trackFile)

            if linktype == 'ShapeSource':
                shtdata = shts(credic[linktype])
                shtdata.Activate()
                if not saveTableAsPicture:
                    if self.pptShape.HasTable:
                        shtdata.Cells.Copy()
                        self.pptShape.Table.Cell(1,1).Select()
                        self.pptWin = self.pptApp.ActiveWindow
                        self.pptWin.View.Paste()
                    else:
                        msg = 'index %s failed!' % i
                        print(msg)
                        self.linkExcelTrack(msg,trackFile)  
                        return False
                else:
                    self.pptShape.Delete()
                    shtdata.UsedRange.Copy()
                    self.Shape = self.pptSlide.Shapes.PasteSpecial(DataType=1)
                    self.setPositionSize(self.Shape,LockAspectRatio=0,
                                         Left=credic['Left'],Top=credic['Top'],Width=credic['Width'],Height=credic['Height'])
                    for i in range(credic['ZOrderPosition'],slideShapesCounts):
                        self.Shape.ZOrder(3)
                msg = 'index %s done!' % i
                print(msg)
                return True

        except:
            traceback.print_exc()
            msg = 'index %s failed!' % i
            print(msg,file=open(trackFile,'a'))
            return False
    def linkExcelinitFiles(self,dirname):
        '''
        linkPPT, init files
        '''
        pptFile = os.path.join(dirname,'powerpoint.pptx')
        excelFile = os.path.join(dirname,'powerpoint_data.xlsx')
        trackFile =  os.path.join(dirname,'powerpoint_record.txt')
        return pptFile,excelFile,trackFile
    def linkExcelTrack(self,msg,trackFile):
        print(msg,file=open(trackFile,'a'))
        return False
        


# #### Test_linkPPTtoExcel

# In[3]:


if __name__ == '__main__':
    ## export
    MPE = MSOPowerpointEase(Blank=False)
    FileName = r'C:/SmithYe/PythonProject3/OfficeApi/MSOPowerpoint/testfiles/teo.pptx'
    #FileName = r'C:/SmithYe/PythonProject3/OfficeApi/MSOPowerpoint/template.pptx'
    dirname = MPE.linkPPTtoExcel(FileName=FileName,keepChartListObjects=None)


# #### Test_linkExceltoPPT

# In[6]:


if __name__ == '__main__':
    # create
    MPE = MSOPowerpointEase(Blank=None)
    dirname = r'C:\SmithYe\PythonProject3\OfficeApi\MSOPowerpoint\testfiles\template_Group_20181219095409'
    pptFile = MPE.linkExceltoPPT(dirname,saveTableAsPicture=True,keepChartListObjects=True)

