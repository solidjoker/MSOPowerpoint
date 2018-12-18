
# coding: utf-8

# In[2]:


import os, time, datetime, pprint, traceback, tempfile
from operator import itemgetter
from collections import defaultdict
from PIL import Image
import pandas as pd
import numpy as np
from itertools import chain

import win32com
from win32com.client import Dispatch,Constants

import MSOPowerpointConfig
import MSOPowerpointFunc
import MSOPowerpointBase


# In[3]:


class MSOPowerpointElement():
    def addShape(self,AutoShapeType=1,Left=0,Top=0,Width=100,Height=20):
        '''
        Shape: add AutoShape
        '''
        try:
            if type(AutoShapeType) is not int: AutoShapeType = self.msoAutoShapeType.get(AutoShapeType,None)
            if self.presInfo['SlidesCount'] == 0: self.addSlide()
            if AutoShapeType:
                self.pptShape = self.pptSlide.Shapes.AddShape(Type=AutoShapeType,Left=Left,Top=Top,Width=Width,Height=Height)
                self.setShape()
                print('Shape :%s ShapeType:%s added!'%(self.pptShape.Name,MSOPowerpointFunc.funcLookupKey(AutoShapeType,self.msoAutoShapeType)))
                return True
            return False
        except:
            traceback.print_exc()
            return False
    def selectShape(self,Index=None,SlideIndex=None,ShapeName=None,ZOrderPosition=None):
        '''
        Index
        selectShape according to self.shapesInfo
        SlideIndex: integer
        ShapeName or ZOrderPosition
        '''
        assert self.shapesInfo, 'Please run self.exportShapesInfo() first!'
        try:
            if Index != None:
                assert 1<=Index<=self.presInfo['ShapesCount'], 'Index must be in [1,%s]!'%self.presInfo['ShapesCount']
                for k,v in self.shapesInfo.items():
                    if Index in v.keys():
                        self.pptShape = self.pptSlides(v[Index]['SlideIndex']).Shapes(v[Index]['ShapeName'])
                        self.setShape()
                        print('Shape Index:%s,ShapeName:%s on Slide:%s selected!'%(Index,v[Index]['ShapeName'],v[Index]['SlideIndex']))
                        return True
            else:
                assert SlideIndex, 'Without Index, SlideIndex must be set!'
                assert any((ShapeName,ZOrderPosition)) and not all((ShapeName,ZOrderPosition)), 'only one of [ShapeName,ZOrderPosition]'
                if ShapeName:
                    for k,v in self.shapesInfo.items():
                        for vv in v.values():
                            if vv['SlideIndex'] == SlideIndex and vv['ShapeName'] == ShapeName:
                                self.pptShape = self.pptSlides(vv['SlideIndex']).Shapes(vv['ShapeName'])
                                self.setShape()
                                print('ShapeName:%s on Slide:%s selected!'%(ShapeName,vv['SlideIndex']))
                                return True                                
                else:
                    for k,v in self.shapesInfo.items():
                        for vv in v.values():
                            if vv['SlideIndex'] == SlideIndex and vv['ShapeZOrderPosition'] == ZOrderPosition:
                                self.pptShape = self.pptSlides(vv['SlideIndex']).Shapes(vv['ShapeName'])
                                self.setShape()
                                print('ShapeZOrderPosition:%s on Slide:%s selected!'%(ZOrderPosition,vv['SlideIndex']))
                                return True                             
                return False
        except:
            traceback.print_exc()
            return False   
    def delShape(self):
        '''
        del Shape
        '''
        try:
            print('Shape:%s deleted!'%(self.pptShape.Name))
            self.pptShape.Delete()
            self.pptShape = None
            self.setShape()
            return True
        except:
            traceback.print_exc()
            return False
    def fillShape(self,Shape=None,r=0,g=0,b=0,ColorName=None,Visible=1,Transparency=0):
        '''
        Shape: fill
        '''
        try:
            Shape = self.initShape(Shape=Shape)
            if not Shape: return False
            self.setMSOobj(obj=Shape.Fill,r=r,g=g,b=b,ColorName=ColorName,Visible=Visible,Transparency=Transparency,)
            return True
        except:
            traceback.print_exc()
            return False        
    def lineShape(self,Shape=None,r=0,g=0,b=0,ColorName=None,Visible=1,Transparency=0,Weight=1,DashStyle=1):
        '''
        Shape: line
        '''
        try:
            Shape = self.initShape(Shape=Shape)
            if not Shape: return False
            if DashStyle == -2: return False
            self.setMSOobj(obj=Shape.Line,r=r,g=g,b=b,ColorName=ColorName,
                           Visible=Visible,Transparency=Transparency,Weight=Weight,DashStyle=DashStyle)
            return True
        except:
            traceback.print_exc()
            return False
    def hyperlinkShape(self,Shape=None,Address=None,SubAddress=None):
        '''
        Shape:Hyperlink
        '''
        try:
            Shape = self.initShape(Shape=Shape)
            if not Shape: return False
            ActonSettings = Shape.ActionSettings(1) # ppMouseClick
            ActonSettings.Action = 7 # ppActionHyperlink
            if Address:
                ActonSettings.Hyperlink.Address = Address
                print('%s Text Hyperlink %s setted!'%(Shape.Name,Address))
                return True
            elif SubAddress:
                ActonSettings.Hyperlink.SubAddress = SubAddress
                print('%s Text Hyperlink %s setted!'%(Shape.Name,SubAddress))
                return True
            else:
                return False
        except:
            traceback.print_exc()
            return False
    def initShape(self,Shape=None):
        if not Shape:
            Shape = self.pptShape
        if not Shape:
            return False
        return Shape
    def ungroupShape(self):
        '''
        Shapes:Ungroup
        '''
        for SlideIndex in range(1,self.presInfo['SlidesCount']+1):
            pptSlide = self.pptSlides.Item(SlideIndex)
            while True:
                Shapes = pptSlide.Shapes
                groups = [Shape for Shape in Shapes if Shape.Type == 6]
                if groups:
                    for group in groups:
                        group.Ungroup()
                else:
                    break
        print('Shapes ungrouped!')
    def sortShapes(self,TableForward=True):
        '''
        根据Shape的上、左排序
        '''
        for SlideIndex in range(1,self.presInfo['SlidesCount']+1):
            pptSlide = self.pptSlides.Item(SlideIndex)
            Shapes = pptSlide.Shapes
            positions = [(Shape,-Shape.HasTable,Shape.Top,Shape.Left,) for Shape in Shapes]
            if TableForward:
                positions = sorted(positions,key=itemgetter(1,2,3))
                msg = 'Table and '
            else:
                positions = sorted(positions,key=itemgetter(2,3))
                msg = ''
            for Shape in positions:
                Shape[0].ZOrder(0) # 第一个移到最后，
        print('Shapes sorted by %sTop and Left!'%msg)
        
    def getShapeHyperlinksDict(self):
        '''
        HyperlinksDict
        '''
        HyperlinksDict = {}
        for SlideIndex in range(1,self.presInfo['SlidesCount']+1):
            pptSlide = self.pptSlides.Item(SlideIndex)
            if pptSlide.Hyperlinks:
                for hl in pptSlide.Hyperlinks:
                    Address =  hl.Address
                    SubAddress = hl.SubAddress
                    ZOrderPosition = MSOPowerpointFunc.funcGetShapeFromHyperlink(hl)
                    if ZOrderPosition:
                        HyperlinksDict[(SlideIndex,ZOrderPosition)] = (True,Address,SubAddress)
        return HyperlinksDict
    
    def addLabel(self,Text=None,Orientation=1,Left=0,Top=0,Width=100,Height=20,Initialized=True):
        '''
        Label: add text label
        Orientation: 1:horizon,6:Chinese vertical
        '''
        try:
            assert Orientation == 1 or Orientation == 6, 'Orientation must be 1 or 6'
            if self.presInfo['SlidesCount'] == 0: self.addSlide()
            self.pptShape = self.pptSlide.Shapes.AddLabel(Orientation=Orientation,Left=Left,Top=Top,Width=Width,Height=Height)
            self.addText(Text=Text)
            self.setShape()
            if Initialized:
                self.setTextFrame()
            print('Shape Label:%s added!'%self.pptShape.Name)
            return True
        except:
            traceback.print_exc()
            return False     
    def addText(self,Shape=None,Text=None):
        '''
        Text: add text
        '''
        try:
            Shape = self.initShape(Shape=Shape)
            if not Shape: return False
            if not Text:
                Text = ''
            if Shape.HasTextFrame:
                Shape.TextFrame.TextRange.Text = Text
                print('Text:%s added in Shape:%s!'%(Text,Shape.Name))
                return True
            return False
        except:
            traceback.print_exc()
            return False      
    def setTextFrame(self,Shape=None,AutoSize=0,HorizontalAnchor=1,VerticalAnchor=3,
                     MarginLeft=10,MarginTop=10,MarginRight=10,MarginBottom=10):
        '''
        Text: Aitosize and alignment
        AutoSize: 0:不自动调整,1:根据文本调整
        HorizontalAnchor: 1:左对齐，2:右对齐
        VerticalAnchor: 1:垂直上，3:垂直中，4:垂直下
        MarginLeft,MarginTop,MarginRight,MarginBottom: pixel?
        '''
        try:
            Shape = self.initShape(Shape=Shape)
            if not Shape: return False
            if Shape.HasTextFrame:
                TextFrame = Shape.TextFrame
                TextFrame.AutoSize = AutoSize  # 0:不自动调整,1:根据文本调整
                TextFrame.TextRange.ParagraphFormat.Alignment = HorizontalAnchor # 1:左对齐，2:中对齐
                TextFrame.VerticalAnchor = VerticalAnchor # 1:垂直上，3:垂直中，4:垂直下
                TextFrame.MarginLeft = MarginLeft 
                TextFrame.MarginTop = MarginTop
                TextFrame.MarginRight = MarginRight
                TextFrame.MarginBottom = MarginBottom
                print('Label %s setted!'%(Shape.Name))
                return True
            return False
        except:
            traceback.print_exc()
            return False  
    def fillText(self,Shape=None,r=0,g=0,b=0,ColorName=None,
                Size=None,Name=None,Bold=0,Italic=0,Underline=0,Strike=0):
        '''
        Text: Size,Name,Bold,Italic,Underline,Strike
        '''
        try:
            Shape = self.initShape(Shape=Shape)
            if not Shape: return False
            if Shape.HasTextFrame:
                for Font in [Shape.TextFrame.TextRange.Font,Shape.TextFrame2.TextRange.Font]:
                    self.setMSOobj(obj=Font,r=r,g=g,b=b,ColorName=ColorName,
                                   Size=Size,Name=Name,Bold=Bold,Italic=Italic,Underline=Underline,Strike=Strike)
                print('%s Text Range setted!'%(Shape.Name))
                return True
            return False
        except:
            traceback.print_exc()
            return False   
    def setTextCharacters(self,Shape=None,r=0,g=0,b=0,ColorName=None,
                          Start=1,Length=0,Size=0,Name=None,Bold=0,Italic=0,Underline=0):
        '''
        CharactersRange:Start,Length,Size,Name,Bold,Italic,Underline,Strike,Color
        '''
        try:
            Shape = self.initShape(Shape=Shape)
            if not Shape: return False
            if Shape.HasTextFrame:
                Font = Shape.TextFrame.TextRange.Characters(Start=Start, Length=Length).Font
                self.setMSOobj(obj=Font,r=r,g=g,b=b,ColorName=ColorName,
                               Size=Size,Name=Name,Bold=Bold,Italic=Italic,Underline=Underline)
                print('%s Text Characters (Start:%s, Length:%s) setted!'%(Shape.Name,Start,Start+Length-1))
                return True
            return False
        except:
            traceback.print_exc()
            return False     
    def addPicture(self,Picture=None,LinkToFile=0,SaveWithDocument=-1,Left=0,Top=0,Width=-1,Height=-1,AdjustDpi=True):
        '''
        Picture: Insert Picture
        '''
        try:
            assert os.path.exists(Picture),'%s not found!'%Picture
            if AdjustDpi: Picture = MSOPowerpointFunc.funcPictureDpi(Picture,self.presInfo['DPI'])
            if self.presInfo['SlidesCount'] == 0: self.addSlide()
            self.pptShape = self.pptSlide.Shapes.AddPicture(FileName=Picture,
                                           LinkToFile=LinkToFile,
                                           SaveWithDocument=SaveWithDocument,
                                           Left=Left,Top=Top,Width=Width,Height=Height)
            self.setShape()
            print('Shape Picture:%s added!'%self.pptShape.Name)
            return self.pptShape
        except:
            traceback.print_exc()
            return False
    def addLine(self,BeginX=0,BeginY=0,EndX=100,EndY=100):
        '''
        Line: add Line
        '''
        try:
            if self.presInfo['SlidesCount'] == 0: self.addSlide()
            self.pptShape = self.pptSlide.Shapes.AddLine(BeginX=BeginX,BeginY=BeginY,EndX=EndX,EndY=EndY)
            self.setShape()
            print('Shape Line:%s added!'%self.pptShape.Name)
            return True
        except:
            traceback.print_exc()
            return False       
        
    def addTable(self,NumRows=10,NumColumns=5,Left=0,Top=0,Width=600,Height=400,Initialized=True,StyleName=None):
        '''
        Table: addTable
        AutoAdjust: True:Adjusted
        '''
        try:
            if self.presInfo['SlidesCount'] == 0: self.addSlide()
            self.pptShape = self.pptSlide.Shapes.AddTable(NumRows=NumRows,NumColumns=NumColumns, 
                                                          Left=Left,Top=Top,Width=Width,Height=Height)
            self.setShape()
            if Initialized:
                self.setTableStyle()
                self.fillTableText(ColorName='black') # 表格所有单元格
                self.fillTableText(Cells=self.pptTable.Rows(1).Cells,ColorName='black',Bold=True) # 表格第一行
                self.fillTable(Cells=self.pptTable.Rows(1).Cells,ColorName='lightblue') # 表格第一行
                self.setTableWidthHeight()
                for Row in range(1,self.pptTable.Rows.Count+1):
                    for Col in range(1,self.pptTable.Columns.Count+1):
                        for attr in ['MarginLeft','MarginRight','MarginTop','MarginBottom']:
                            setattr(self.pptTable.Cell(Row,Col).Shape.TextFrame2,attr,0)
            elif StyleName:
                self.setTableStyle(StyleName=StyleName)
            print('Shape Table:%s added!'%self.pptShape.Name)
            return True
        except:
            traceback.print_exc()
            return False    
    def setTableStyle(self,StyleName=None):
        '''
        Table:Style
        '''
        try:
            assert self.pptTable,'self.pptShape has not Table!'
            if not StyleName:
                StyleName = '无样式，网格型'
            StyleID = self.msoTableStype.get(StyleName,None)
            self.pptTable.ApplyStyle(StyleID)
            print('Table in Shape:%s applied %s!'%(self.pptShape.Name,StyleName))
            return True
        except:
            traceback.print_exc()
            return False
    def fillTableBorder(self,Cells=None,r=0,g=0,b=0,ColorName=None,DashStyle=1,Weight=0.5):
        '''
        Table: Table Border
        '''
        try:
            assert self.pptTable,'self.pptShape has not Table!'
            if Cells:
                CellsRange = iter([Cells])
            else:
                CellsRange = iter([Row.Cells for Row in self.pptTable.Rows])
            for Cells in CellsRange:
                for i in (1,2,3,4):
                    self.setMSOobj(obj=Cells.Borders(i),r=r,g=g,b=b,ColorName=ColorName,ShowMessage=False,
                                   Visible=1,DashStyle=DashStyle,Weight=Weight)
            print('Table in Shape:%s added border!'%self.pptShape.Name)
            return True
        except:
            traceback.print_exc()
            return False   
    def fillTable(self,Cells=None,r=0,g=0,b=0,ColorName=None,Visible=1):
        '''
        Table: Table Color
        '''    
        try:
            assert self.pptTable,'self.pptShape has not Table!'
            if not Cells:
                Cells = iter(chain.from_iterable([Row.Cells for Row in self.pptTable.Rows]))
            for Cell in Cells:
                self.setMSOobj(obj=Cell.Shape.Fill,r=r,g=g,b=b,ColorName=ColorName,ShowMessage=False,Visible=Visible)
            print('Table in Shape:%s Cell filled!'%self.pptShape.Name)
            return True
        except:
            traceback.print_exc()
            return False  
    def fillTableText(self,Cells=None,r=0,g=0,b=0,ColorName=None,HorizontalAnchor=2,VerticalAnchor=3,Name='微软雅黑',Size=10,Bold=0,):
        '''
        Table: Table Text Format
        HorizontalAnchor: 1:左对齐，2:中对齐
        VerticalAnchor: 1:垂直上，3:垂直中，4:垂直下
        '''    
        try:
            assert self.pptTable,'self.pptShape has not Table!'
            if not Cells:
                Cells = iter(chain.from_iterable([Row.Cells for Row in self.pptTable.Rows]))
            for Cell in Cells:
                Cell.Shape.TextFrame.HorizontalAnchor = HorizontalAnchor # 1:左对齐，2:中对齐
                Cell.Shape.TextFrame.VerticalAnchor = VerticalAnchor # 1:垂直上，3:垂直中，4:垂直下
                self.setMSOobj(obj=Cell.Shape.TextFrame.TextRange.Font,r=r,g=g,b=b,ColorName=ColorName,ShowMessage=False,
                               Size=Size,Name=Name,Bold=Bold)
            print('Table in Shape:%s text format setted!'%self.pptShape.Name)
            return True
        except:
            traceback.print_exc()
            return False
    def setTableWidthHeight(self,ColRow='row',Index=1,Point=None):
        '''
        Table:ColRow,Height,Width
        ColRow: col or row
        Index: col or row number
        '''
        try:
            assert self.pptTable,'self.pptShape has not Table!'
            cr = {'col':['Columns','Width'],
                'row':['Rows','Height']}
            colrow = cr.get(ColRow.lower()[:3],None)
            if colrow:
                colrowParent = getattr(self.pptTable,colrow[0])
                if not Point:
                    Point = getattr(self.pptShape,colrow[1])/colrowParent.Count
                if Index:
                    setattr(colrowParent(Index),colrow[1],Point)
                    print('Table in Shape:%s %s %s width setted!'%(self.pptShape.Name,colrow[0],Index))
                else:
                    for _ in colrow[0]:
                        setattr(_,colrow[1],Point)
                    print('Table in Shape:%s %s %s setted!'%self.pptShape.Name,colrow[0],colrow[1])
                return True
            else:
                return False
        except:
            traceback.print_exc()
            return False     
    def setTableFromPandas(self,df,Left=0,Top=0,CellWidth=200,CellHeight=40,Initialized=True):
        '''
        Table: from Pandas
        '''
        try:
            rs,cs = df.shape
            headers = df.columns.tolist()
            for col in df.columns: # 数据格式调整
                if np.issubdtype(df[col],np.datetime64):
                    df[col] = df[col].apply(lambda x: datetime.datetime.strftime(x,'%Y/%m/%d'))
                if np.issubdtype(df[col],np.integer):
                    df[col] = df[col].apply(lambda x: '{:,}'.format(x))
                if np.issubdtype(df[col],np.number):
                    df[col] = df[col].apply(lambda x: '{:,.2f}'.format(x))
            mx = df.values.tolist()
            NumRows,NumColumns = rs+1,cs
            self.addTable(NumRows,NumColumns,Left,Top,CellWidth*NumColumns,CellHeight*NumRows,Initialized=Initialized)
            for Row in range(1,self.pptTable.Rows.Count+1):
                for Col in range(1,self.pptTable.Columns.Count+1):
                    if Row == 1:
                        self.pptTable.Cell(Row,Col).Shape.TextFrame.TextRange.Text = headers[Col-1]
                    else:
                        self.pptTable.Cell(Row,Col).Shape.TextFrame.TextRange.Text = mx[Row-2][Col-1]
            print('Table in Shape:%s text from pandas setted!'%self.pptShape.Name)
            return True
        except:
            traceback.print_exc()
            return False
        
    def addChart(self,xlChartType=1,Left=0,Top=0,Width=-1,Height=-1,PlotBy=1,Initialized=True):
        '''
        Chart: add Chart
        xlChartType: 数值,可以通过self.xlChartType查阅
        PlotBy: 1:xlRows,2:xlColumns
        '''
        assert type(xlChartType) == int, '%s must be integer!'
        try:
            if self.presInfo['SlidesCount'] == 0: self.addSlide()
            self.pptShape = self.pptSlide.Shapes.AddChart(xlChartType,Left,Top,Width,Height)
            self.setShape()
            print('Shape:%s Chart:%s added!'%(self.pptShape.Name,self.pptChart.Name))
            pptWorkbook,pptWorksheet = self.initChartData() # 初始化
            if Initialized:
                self.pptChart.PlotBy = PlotBy # 1:xlRows,2:xlColumns
                for i in range(self.pptSeriesCollection().Count):
                    self.pptSeriesCollection(1).Delete() # 清空Series
                pptWorksheet.Cells.Clear() # 清空图表数据
                print('Shape:%s Chart:%s initialized!'%(self.pptShape.Name,self.pptChart.Name))
                self.pptChart.Refresh()
                time.sleep(self.seconds)
            pptWorkbook.Close()
            time.sleep(self.seconds)
            return True
        except:
            traceback.print_exc()
            return False  
    def initChartData(self):
        self.pptChartData.Activate() # 打开Excel，需要关闭
        pptWorkbook = self.pptChartData.Workbook
        pptWorkbook.Parent.DisplayAlerts = False
        pptWorksheet = pptWorkbook.ActiveSheet
        return pptWorkbook,pptWorksheet   
    def changeChartType(self,xlChartType,Index=None):
        '''
        Chart: ChangeType
        xlChartType: Integer: use self.xlChartType for reference
        '''
        assert type(xlChartType) == int, '%s must be integer!'
        try:
            _xlChartType = self.pptChart.ChartType # 返回图表类型
            pptWorkbook,pptWorksheet = self.initChartData()  # 初始化
            if Index:
                self.pptSeriesCollection(Index).ChartType = xlChartType
                print('Shape:%s Chart:%s Index:%s change type from %s to %s!'%(self.pptShape.Name,self.pptChart.Name,Index,
                    MSOPowerpointFunc.funcLookupKey(_xlChartType,self.xlChartType),MSOPowerpointFunc.funcLookupKey(xlChartType,self.xlChartType)))
            else:
                self.pptChart.ChartType = xlChartType
                print('Shape:%s Chart:%s change type from %s to %s!'%(self.pptShape.Name,self.pptChart.Name,
                    MSOPowerpointFunc.funcLookupKey(_xlChartType,self.xlChartType),MSOPowerpointFunc.funcLookupKey(xlChartType,self.xlChartType)))
            self.pptChart.Refresh()
            time.sleep(self.seconds)
            pptWorkbook.Close()
            time.sleep(self.seconds)
            return True
        except:
            traceback.print_exc()
            return False       
    def setChartTitle(self,ChartTitle=None,Size=None,Name=None,Bold=True,Italic=None):
        '''
        Chart: ChartTitle
        '''
        try:
            if ChartTitle:
                if not self.pptChart.HasTitle:
                    self.pptChart.HasTitle = True
                self.pptChart.ChartTitle.Text = ChartTitle
                self.setMSOobj(obj=self.pptChart.ChartTitle.Font,Size=Size,Name=Name,Bold=Bold,Italic=Italic)
            else:
                self.pptChart.HasTitle = False
            print('Shape:%s Chart:%s ChartTitle setted!'%(self.pptShape.Name,self.pptChart.Name))
            return True
        except:
            traceback.print_exc()
            return False
    def setChartLegend(self,Legend=True,Size=None,Name=None,Bold=None,Italic=None):
        '''
        Chart: Chart Legend
        '''
        try:
            if Legend:
                self.pptChart.HasLegend = True
                self.setMSOobj(obj=self.pptChart.Legend.Font,Size=Size,Name=Name,Bold=Bold,Italic=Italic)
            else:
                self.pptChart.HasLegend = False
            print('Shape:%s Chart:%s Legend adjusted!'%(self.pptShape.Name,self.pptChart.Name))
            return True
        except:
            traceback.print_exc()
            return False
    def setChartAxes(self,Axis=None,AxisGroup=1,Size=None,Name=None,Bold=None,Italic=None
                     ,NumberFormatLocal=None,MajorGridlines=None,MinorGridlines=None,Delete=None):
        '''
        Chart: Chart Axes
        '''
        try:
            if AxisGroup == 2:
                Axis = 2
            if not Axis:
                Axes = [1,2] # 1:xlCategory,2:xlValue
            else:
                Axes = [Axis]
            AxesD = {1:349,2:353} # 349:msoElementPrimaryCategoryAxisShow,353:msoElementPrimaryValueAxisShow
            for _ in Axes:
                self.pptChart.SetElement(AxesD[_])
                _Axes = self.pptChart.Axes(_,AxisGroup)
                TickLabels = _Axes.TickLabels
                self.setMSOobj(obj=TickLabels.Font,Size=Size,Name=Name,Bold=Bold,Italic=Italic)
                if NumberFormatLocal:
                    TickLabels.NumberFormatLocal = NumberFormatLocal
                if MajorGridlines:
                    _Axes.HasMajorGridlines = True
                else:
                    _Axes.HasMajorGridlines = False
                if MinorGridlines:
                    _Axes.HasMinorGridlines = True
                else:
                    _Axes.HasMinorGridlines = False                    
                if Delete:
                    TickLabels.Delete()
            print('Shape:%s Chart:%s Axes %s adjusted!'%(self.pptShape.Name,self.pptChart.Name,Axes))
            return True
        except:
            traceback.print_exc()
            return False 
    def changeChartAxisGroup(self,Index=None,AxisGroup=2):
        '''
        Chart: AxisGroup
        '''
        try:
            if Index:
                self.pptSeriesCollection(Index).AxisGroup = AxisGroup
                print('Shape:%s Chart:%s Index:%s AxisGroup changed!'%(self.pptShape.Name,self.pptChart.Name,Index))
            return True
        except:
            traceback.print_exc()
            return False
    def setChartDataLabel(self,Index=None,Size=None,Name=None,Bold=None,Italic=None,NumberFormatLocal=None):
        '''
        Chart: DataLabel
        '''
        assert Index is None or type(Index) == int, 'Index must be None or Integer'
        try:
            if Index:
                Indexes = [Index]
            else:
                Indexes = range(1,self.pptSeriesCollection().Count+1)
            for Index in Indexes:
                SeriesCollection = self.pptSeriesCollection(Index)
                SeriesCollection.ApplyDataLabels() # 应用DataLables
                DataLabels = SeriesCollection.DataLabels()
                self.setMSOobj(obj=DataLabels.Format.TextFrame2.TextRange.Font,Size=Size,Name=Name,Bold=Bold,Italic=Italic)
                if NumberFormatLocal:
                    DataLabels.NumberFormatLocal = NumberFormatLocal
            print('Shape:%s Chart:%s Index %s, DataLabel setted!'%(self.pptShape.Name,self.pptChart.Name,Index))
            return True
        except:
            traceback.print_exc()
            return False
    def showChartDataLabel(self,Index=None,ShowSeriesName=False,ShowCategoryName=False,ShowValue=True,ShowPercentage=False):
        '''
        Chart: DataLabel ShowSeriesName,ShowCategoryName,ShowValue,ShowPercentage
        '''
        assert Index is None or type(Index) == int, 'Index must be None or Integer'
        try:
            if Index:
                Indexes = [Index]
            else:
                Indexes = list(range(1,self.pptSeriesCollection().Count+1))
            for Index in Indexes:
                SeriesCollection = self.pptSeriesCollection(Index)
                SeriesCollection.ApplyDataLabels() # 应用DataLables
                DataLabels = SeriesCollection.DataLabels()
                DataLabels.ShowSeriesName = ShowSeriesName
                DataLabels.ShowCategoryName = ShowCategoryName
                DataLabels.ShowValue = ShowValue
                DataLabels.ShowPercentage = ShowPercentage
            print('Shape:%s Chart:%s Index %s, DataLabel showed!'%(self.pptShape.Name,self.pptChart.Name,Index))
            return True
        except:
            traceback.print_exc()
            return False     
    def replaceChartDataLabel(self,Index=None,Replacements=[]):
        '''
        Chart: replace DataLabel
        '''
        assert type(Index) == int, 'Index must be Integer'
        assert type(Replacements) == list, 'Replacements must be List'
        try:
            SeriesCollection = self.pptSeriesCollection(Index)
            SeriesCollection.ApplyDataLabels() # 应用DataLables
            DataLabels = SeriesCollection.DataLabels()
            for datalabel,replacement in zip(range(1,DataLabels.Count+1),Replacements):
                DataLabels(datalabel).Format.TextFrame2.TextRange.Characters.Text = replacement
            print('Shape:%s Chart:%s Index %s, DataLabel replaced!'%(self.pptShape.Name,self.pptChart.Name,Index))
            return True
        except:
            traceback.print_exc()
            return False  
    def showChartDataTable(self,HasDataTable=True):
        '''
        Chart: show DataTable
        '''        
        try:
            self.pptChart.HasDataTable = HasDataTable
            print('Shape:%s Chart:%s DataTable showed!'%(self.pptShape.Name,self.pptChart.Name))
            return True
        except:
            traceback.print_exc()
            return False
    def setChartDataFromPandas(self,df=None,PlotBy=1):
        '''
        Chart: set data from df
        '''
        try:
            rs,cs = df.shape
            headers = df.columns
            for col in df.columns:
                if np.issubdtype(df[col],np.datetime64): # 日期格式
                    df[col] = df[col].apply(lambda x: datetime.datetime.strftime(x,'%Y/%m/%d'))
            mx = df.values.tolist()
            tries = self.tries
            while tries:
                try:
                    pptWorkbook,pptWorksheet = self.initChartData() # 初始化
                    break
                except:
                    tries -= 1
                    time.sleep(self.seconds*1.5)
            for Row in range(1,rs+2):
                for Col in range(1,cs+1):
                    if Row == 1:
                        pptWorksheet.Cells(Row,Col).Value = headers[Col-1]
                    else:
                        pptWorksheet.Cells(Row,Col).Value = mx[Row-2][Col-1]
            self.pptChart.SetSourceData(Source='=Sheet1!%s'% pptWorksheet.Usedrange.Address,PlotBy=PlotBy)
            self.pptChart.Refresh()
            time.sleep(self.seconds)
            pptWorkbook.Close() # 关闭表格
            time.sleep(self.seconds)
            print('Shape:%s ChartData value from pandas setted!'% self.pptShape.Name)
            return True
        except:
            traceback.print_exc()
            return False   
    def adjustChartSplitValue(self,SplitValue=None):
        '''
        Chart: BarOfPie SplitValue
        '''
        assert type(SplitValue) == int, 'SplitValue must be integer!'
        assert self.pptChart.ChartType in [self.xlChartType['xlBarOfPie']], "ChartType must in xlChartType['xlBarOfPie']"
        try:
            self.pptChart.ChartGroups(1).SplitValue = SplitValue
            print('Shape:%s Chart:%s SplitValue adjusted!'%(self.pptShape.Name,self.pptChart.Name))
            return True
        except:
            traceback.print_exc()
            return False
    
    def fillChart(self,Index=None,r=0,g=0,b=0,ColorName=None,Visible=1,Transparency=0):
        '''
        Chart: fill Chart
        '''
        assert type(Index) == int, 'Index must be Integer'
        try:    
            Fill = self.pptSeriesCollection(Index).Format.Fill
            self.setMSOobj(obj=Fill,r=r,g=g,b=b,ColorName=ColorName,Visible=Visible,Transparency=0)
            print('Shape:%s Chart:%s Index %s, chart filled!'%(self.pptShape.Name,self.pptChart.Name,Index))
            return True
        except:
            traceback.print_exc()
            return False 
    def fillChartPoint(self,Index=None,Point=None,r=0,g=0,b=0,ColorName=None,Visible=1,Transparency=0):
        '''
        Chart: fill Chart point
        '''
        assert type(Index) == int, 'Index must be Integer'
        assert type(Point) == int, 'Point must be Integer'
        try:    
            Fill = self.pptSeriesCollection(Index).Points(Point).Format.Fill
            self.setMSOobj(obj=Fill,r=r,g=g,b=b,ColorName=ColorName,Visible=Visible,Transparency=0)
            print('Shape:%s Chart:%s Index:%s Point:(%s), chart filled!'%(self.pptShape.Name,self.pptChart.Name,Index,Point))
            return True
        except:
            traceback.print_exc()
            return False  
    def lineChart(self,Index=None,r=0,g=0,b=0,ColorName=None,Visible=1,Transparency=0,Weight=1,DashStyle=1):
        '''
        Chart: line chart
        '''
        assert type(Index) == int, 'Index must be Integer'
        try:
            Line = self.pptSeriesCollection(Index).Format.Line
            if DashStyle == -2: return False
            self.setMSOobj(obj=Line,r=r,g=g,b=b,ColorName=ColorName,
                           Visible=Visible,Transparency=Transparency,Weight=Weight,DashStyle=DashStyle)
            print('Shape:%s Chart:%s Index %s, chart lined!'%(self.pptShape.Name,self.pptChart.Name,Index))
            return True
        except:
            traceback.print_exc()
            return False  
    def lineChartPoint(self,Index=None,Point=None,r=0,g=0,b=0,ColorName=None,Visible=1,Transparency=0,Weight=1,DashStyle=1):
        '''
        Chart: line Chart point
        '''
        assert type(Index) == int, 'Index must be Integer'
        assert type(Point) == int, 'Point must be Integer'
        try:    
            if DashStyle == -2: return False
            Line = self.pptSeriesCollection(Index).Points(Point).Format.Line
            self.setMSOobj(obj=Line,r=r,g=g,b=b,ColorName=ColorName,
                           Visible=Visible,Transparency=Transparency,Weight=Weight,DashStyle=DashStyle)
            print('Shape:%s Chart:%s Index:%s Point:(%s), chart lined!'%(self.pptShape.Name,self.pptChart.Name,Index,Point))
            return True
        except:
            traceback.print_exc()
            return False  
    

