
# coding: utf-8

# In[1]:


import os, time, datetime, pprint, traceback, tempfile
from collections import defaultdict
from PIL import Image
import pandas as pd
import numpy as np

import win32com
from win32com.client import Dispatch,Constants

import MSOPowerpointConfig
import MSOPowerpointFunc


# In[1]:


class MSOPowerpointBase():
    def quitAPP(self):
        '''
        close all ppt and quit app
        '''
        _ = input('Save and close all ppt and quit %s\nPlease input yes(y)\n') 
        if _.lower().startswith('y'):
            try:
                for Index in range(self.pptPres.Count,0,-1):
                    if self.pptPres(Index).Path:
                        self.pptPres(Index).Save()
                    else:
                        _filename = self.pptPres(Index).Name
                        _filename = '%s.pptx'%_filename if not _filename.endswith('.pptx') else _filename
                        self.pptPres(Index).SaveAs(os.path.join(os.path.curdir,_filename))
                    print('%s saved and closed'%self.pptPres(Index).FullName)
                    self.pptPres(Index).Close()
            except:
                traceback.print_exc()  
            finally:
                print('%s quit!'%self.pptApp.Name)
                self.pptApp.Quit()
                try:
                    self.pptApp = self.win32com.client.Dispatch(self.appName)
                    self.pptApp.Quit()
                    return True
                except:
                    traceback.print_exc()
                    return False
        else:
            return False      
    def openPPT(self,FileName=None,ReadOnly=0,Untitled=0,WithWindow=-1,GetDpi=0,Blank=True):
        '''
        to open a PPT 
        FileName: relpath or abspath
        ReadOnly: 0:read/write, 1:read only 
        Untitled: 0:no title for new pptx
        WithWindow: 0:no window, 1:show window
        GetDpi: 0:no adjust DPI, 1:adjust DPI
        Blank: True for blank ppt
        '''
        try:
            if FileName:
                assert os.path.exists(FileName),'Not such file %s!'%FileName
                FileName = os.path.realpath(FileName)
                print(FileName)
                self.closePPT(FileName=FileName,Save=True)
                self.pptSel = self.pptPres.Open(FileName=FileName,ReadOnly=ReadOnly,Untitled=Untitled,WithWindow=WithWindow)
                try:
                    self.win32com.client.gencache.EnsureDispatch(self.appName) # 尝试保持连接
                    time.sleep(1)
                except Exception as e:
                    print(e)
                print('Presentation:%s opened!'%self.pptSel.Name)
            else:
                if not Blank:
                    return True
                else:
                    self.pptSel = self.pptPres.Add(WithWindow=WithWindow)
                    print('Presentation:%s opened!'%self.pptSel.Name)
            self.pptSlides = self.pptSel.Slides
            if self.pptSlides.Count:
                self.selectSlide()
            self.getPresInfo(GetDpi=GetDpi)
            return self.pptSel
        except:
            traceback.print_exc()
            return False   
    def savePPT(self,newFileName=None,Close=False):
        '''
        Save PPT
        newFileName: rel or abs path
        close: True for close
        '''
        try:
            if newFileName:
                newFileName = newFileName.replace('/',os.path.sep)
                print('newFileName:%s'%newFileName)
                print('newFileNameRealpath:%s'%os.path.realpath(newFileName))

                if newFileName != os.path.realpath(newFileName):
                    self.pptSel.SaveAs(FileName=os.path.join(os.path.dirname(self.pptSel.FullName),os.path.basename(newFileName)))
                else:
                    self.pptSel.SaveAs(FileName=newFileName)
                print('Presentation:%s saved!'%newFileName)
            else:
                if self.pptSel.Path:
                    self.pptSel.Save()
                else:
                    self.pptSel.SaveAs(os.path.join(os.path.curdir,'%s.pptx'%self.pptSel.Name))
                print('Presentation:%s saved!'%self.pptSel.FullName)
            if Close:
                self.pptSel.Close()
                print('Presentation:%s closed!'%newFileName)
            return True
        except:
            traceback.print_exc()
            return False  
    def closePPT(self,FileName=None,Save=None):
        '''
        close PPT
        Save: True: to Save
        '''
        try:
            if not FileName:
                FileName = self.presInfo['FileName']
            names = {i.Name:i for i in self.pptPres} # pptPres中的ppt
            FileName = os.path.basename(FileName)
            FileName = '%s.pptx' % FileName if not FileName.endswith('.pptx') else FileName
            if FileName in names:
                if Save == True:
                    pptSel = names[FileName]
                    pptSel.Save()
                pptSel.Close()
                return True
            return False
        except:
            traceback.print_exc()
            return False  
    def setupPage(self,SlideWidth=7.5*16/9,SlideHeight=7.5,SlideOrientation=1):
        '''
        setupPage size, width and height
        '''       
        try:
            assert self.pptSel,'pptSel must be setted first!'
            pptPageSetup = self.pptSel.PageSetup        
            pptPageSetup.SlideWidth = SlideWidth * 72
            pptPageSetup.SlideHeight = SlideHeight * 72
            pptPageSetup.SlideOrientation = SlideOrientation # 0:纵向、1:横向
            print('SlideWidth:%.1f inch,SlideHeight:%.1f inch,SlideOrientation:%s'%(SlideWidth,SlideHeight,SlideOrientation))
            return True
        except:
            traceback.print_exc()
            return False   
    def setShape(self):
        if self.pptShape:
            if self.pptShape.HasTable:
                self.pptTable = self.pptShape.Table
            else:
                self.pptTable = None
            if self.pptShape.HasChart:
                self.pptChart = self.pptShape.Chart
                self.pptChartData = self.pptChart.ChartData
                self.pptSeriesCollection = self.pptChart.SeriesCollection
            else:
                self.pptChart,self.pptChartData,self.pptSeriesCollection = None,None,None
        else:
            self.pptShape,self.pptTable,self.pptChart,self.pptChartData,self.pptSeriesCollection = None,None,None,None,None
            
    def getPresInfo(self,GetDpi=0):
        '''
        getPresInfo
        '''
        try:
            self.SlideMaster = self.pptSel.SlideMaster
            self.presInfo = {'FileName': self.pptSel.FullName,
                             'SlidesCount': self.pptSlides.Count,
                             'Height': int(self.SlideMaster.Height), # pixel = inch * 72
                             'Width': int(self.SlideMaster.Width), # pixel = inch * 72
                             'ShapesCount': 'UnScanned',
                             # 'SlideMaster': self.pptSel.SlideMaster,
                             # 'CustomLayouts': self.pptSel.SlideMaster.CustomLayouts,
            }
            if GetDpi:
                self.presInfo['DPI'] = self._funcGetDpi()
            else:
                self.presInfo['DPI'] = 72
            # _ = 'Presentation: {FileName}\nSlidesCount: {SlidesCount}, Height: {Height} pixel, Width: {Width} pixel, DPI: {DPI}'
            # print(_.format(**self.presInfo))
            return True
        except:
            traceback.print_exc()
            return False          
    def getShapesText(self):
        '''
        print all shapes text in ppt
        '''
        self.shapesText = []
        try:
            pptSlides = self.pptSel.Slides
            for pptSlide in pptSlides:
                Shapes = pptSlide.Shapes
                for Shape in Shapes:
                    if Shape.HasTextFrame:
                        self.shapesText.append(Shape.TextFrame.TextRange.Text)
            return True
        except:
            traceback.print_exc()
            return False     
    def addSlide(self,Index=None,Layout=12):
        '''
        addSlide
        Layout:12 blank
        '''
        try:
            if not Index:
                Index = self.pptSlides.Count+1
            self.pptSlide = self.pptSlides.Add(Index=Index,Layout=Layout)
            self.getPresInfo()
            self.pptSlide.Select()
            print('Slide with Index:%s added!'% Index)
            return True
        except:
            traceback.print_exc()
            return False
    def selectSlide(self,Index=1):
        '''
        select a slide with Index
        '''
        try:
            if Index:
                self.pptSlide = self.pptSlides.Item(Index)
                if not self.pptSel.Windows.Count:
                    self.pptSel.NewWindow()
                self.pptSlide.Select()
                print('Slide Index:%s selected!'% Index)
                return True
            return False
        except:
            traceback.print_exc()
            return False    
    def delSlide(self,Index=None):
        '''
        delSlide:
        Index: None for current self.pptSlide
        '''
        try:
            if not Index:
                Index = self.pptSlide.SlideIndex
            print('SlideIndex:%s deleted!'%(Index))
            self.pptSlides(Index).Delete()
            self.getPresInfo()
            if Index > 1: 
                self.selectSlide(Index-1)
            return True
        except:
            traceback.print_exc()
            return False   
    
    def initExcel(self,filename=None,shtname='Summary'):
        '''
        init Excel
        ''' 
        try:
            appExcel = 'Excel.Application'
            ExcelApp = self.win32com.client.Dispatch(appExcel)
            ExcelApp.Visible = 1
            ExcelApp.DisplayAlerts = 1 # set DisplayAlerts to ppAlertsNone
            wkbs = ExcelApp.Workbooks
            if filename:
                wkb = wkbs.Open(filename)
                shts = wkb.WorkSheets
                if shtname:
                    try:
                        shtSum = shts(shtname)
                    except:
                        shtSum = shts(1)
                return ExcelApp, wkbs, wkb, shts, shtSum  
        except:
            traceback.print_exc(limit=1)
            return False,False,False,False,False   
    def slowprint(self,information):
        for i in information:
            print(i,end='')
            time.sleep(self.seconds/100)
        print('') 

    def setMSOobj(self,obj=None,r=None,g=None,b=None,ColorName=None,ShowMessage=True,**kwargs):
        '''
        MSO obj set
        '''
        if not obj:
            return False
        # set ForeColor
        for k,v in kwargs.items():
            if hasattr(obj,k):
                # method
                if callable(getattr(obj,k)):
                    if v == True:
                        try:
                            getattr(obj,k)()
                        except:
                            traceback.print_exc()
                    else:
                        try:
                            getattr(obj,k)(v)
                        except:
                            traceback.print_exc()               
                # attribute
                else:
                    if v != None:
                        try:
                            setattr(obj,k,v)
                        except:
                            traceback.print_exc()
        try:
            color = MSOPowerpointFunc.getRGB(r=r,g=g,b=b,ColorName=ColorName)
            if color != False:
                if hasattr(obj,'ForeColor') and hasattr(obj.ForeColor,'RGB'):
                    setattr(obj.ForeColor,'RGB',color)
                elif hasattr(obj,'Color') and hasattr(obj.Color,'RGB'):
                    setattr(obj.Color,'RGB',color)
        except:
            traceback.print_exc()
        if ShowMessage:
            print('MSOobj setted!')
        return True
    def setPositionSize(self,Shape=None,LockAspectRatio=1,Left=None,Top=None,Width=None,Height=None):
        '''
        set Shape Position and Size：
        LockAspectRatio: 1: lock, 0:unlock
        Left,Top: for position
        Width,Height: for size
        '''
        try:
            Shape = self.initShape(Shape=Shape)
            if not Shape: return False
            if Left:
                Shape.Left = Left
            if Top:
                Shape.Top = Top
            # 锁定纵横比
            Shape.LockAspectRatio = LockAspectRatio
            if Width:
                Shape.Width = Width
            if Height:
                Shape.Height = Height
            print('%s Shape Postion and Size setted!'%(Shape.Name))
            return True
        except:
            traceback.print_exc()
            return False

    def exportPPTtoExcel(self):
        return MSOPowerpointFunc.exportPPTtoExcel(self)
    def exportDfsToExcel(self,dfs,sheetnames,filename):
        MSOPowerpointFunc.exportDfsToExcel(dfs,sheetnames,filename)

