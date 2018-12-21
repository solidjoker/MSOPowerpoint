
# coding: utf-8

# # PowerPointBase
# update: 2018-12-20
# 
# to be fix: 
# <li>export 需要处理负值<li>
# <li>Create 需要精细化</li>
# <li>看看MSOPowerpointFunc中，有什么可以放到CustomFunc</li>

# In[4]:


import os, time, datetime, pprint, traceback, tempfile
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


# In[5]:


class MSOPowerpoint(MSOPowerpointBase,MSOPowerpointElement):
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

    def exportShapesInfo(self):
        '''
        get ShapesInfo
        '''
        try:
            self.ungroupShape() # 去除组合
            self.HyperlinksDict = self.getShapeHyperlinksDict()
            self.shapesInfo = defaultdict(dict) # 初始化
            # shapesCats = ['Table','Chart','SmartArt','Text','Shape',]
            shapesCats = {
                'Table':{'Test':'HasTable','Func':self.exportShapeInfoTable},
                'Chart':{'Test':'HasChart','Func':self.exportShapeInfoChart},
                'SmartArt':{'Test':'HasSmartArt','Func':self.exportShapeInfoSmartArt},
                'Text':{'Test':'HasTextFrame','Func':self.exportShapeInfoText},
                'Shape':{'Test':None,'Func':self.exportShapeInfoShape},
                         }
            for cat in shapesCats:
                self.shapesInfo[cat] = defaultdict(dict)
            self.shapesCount = 0
            for SlideIndex in range(1,self.presInfo['SlidesCount']+1):
                pptSlide = self.pptSlides.Item(SlideIndex)
                print('Slide:%s is scanning!'%SlideIndex)
                Shapes = pptSlide.Shapes
                for Shape in Shapes:
                    self.shapesCount += 1
                    shapesCount = self.shapesCount
                    shapeInfo = {shapesCount:defaultdict(dict)}
                    self.exportShapeInfoBase(shapeInfo,Shape,shapesCount,SlideIndex) # 基本信息
                    for cat in ['Table','Chart','SmartArt','Text']:
                        if getattr(Shape,shapesCats[cat]['Test']):
                            shapesCats[cat]['Func'](shapeInfo,Shape,shapesCount,SlideIndex)
                            break
                        else:
                            cat = 'Shape'
                            shapesCats[cat]['Func'](shapeInfo,Shape,shapesCount,SlideIndex)
                    if shapeInfo[shapesCount]:
                        self.shapesInfo[cat].update(shapeInfo)
            for cat in shapesCats:
                if not self.shapesInfo[cat]:
                    self.shapesInfo.pop(cat)
                # RGB 修复
                for k,v in self.shapesInfo[cat].items():
                    for kk in v:
                        if kk.endswith('RGB'):
                            if self.shapesInfo[cat][k][kk] and self.shapesInfo[cat][k][kk] < 0:
                                self.shapesInfo[cat][k][kk] = np.nan
            self.getPresInfo()
            self.presInfo['ShapesCount'] = self.shapesCount
            return True
        except:
            traceback.print_exc()
            return False
    def exportShapeInfoBase(self,shapeInfo,Shape,shapesCount,SlideIndex):
        # 页面信息
        shapeInfo[shapesCount]['SlideIndex'] = SlideIndex
        shapeInfo[shapesCount]['ShapeName'] = Shape.Name
        # 基本信息
        shapeInfo[shapesCount]['ShapeType'] = MSOPowerpointFunc.funcLookupKey(Shape.Type,self.msoShapeType)
        shapeInfo[shapesCount]['AutoShapeType'] = MSOPowerpointFunc.funcLookupKey(Shape.AutoShapeType,self.msoAutoShapeType)
        shapeInfo[shapesCount]['ShapeZOrderPosition'] = Shape.ZOrderPosition
        attrs = ['Visible','Left','Top','Width','Height']
        for attr in attrs:
            shapeInfo[shapesCount][attr] = int(getattr(Shape,attr))
        Hyperlinks,Address,SubAddress = self.HyperlinksDict.get((SlideIndex,Shape.ZOrderPosition),(None,None,None))
        shapeInfo[shapesCount]['HyperLink'] = Hyperlinks
        shapeInfo[shapesCount]['Address'] = Address
        shapeInfo[shapesCount]['SubAddress'] = SubAddress
        self.exportShapeInfoFill(shapeInfo,Shape.Fill,shapesCount,SlideIndex)
        self.exportShapeInfoLine(shapeInfo,Shape.Line,shapesCount,SlideIndex)       
    def exportShapeInfoFill(self,shapeInfo,Fill,shapesCount,SlideIndex):
        # 底色
        shapeInfo[shapesCount]['Fill'] = None
        shapeInfo[shapesCount]['FillRGB'] = None
        try:
            if Fill.Visible:
                shapeInfo[shapesCount]['Fill'] = True
                shapeInfo[shapesCount]['FillRGB'] = int(Fill.ForeColor.__str__())
        except:
            traceback.print_exc(limit=1)
    def exportShapeInfoLine(self,shapeInfo,Line,shapesCount,SlideIndex):
        # 边框
        shapeInfo[shapesCount]['Line'] = None
        shapeInfo[shapesCount]['LineRGB'] = None
        shapeInfo[shapesCount]['Weight'] = None
        shapeInfo[shapesCount]['DashStyle'] = None
        try:
            if Line.Parent.HasTable:
                Line = Line.Parent.Table.Rows(1).Cells(1).Borders(1)
                # Table的Line属性特别奇怪
                if Line.Visible:
                    shapeInfo[shapesCount]['Line'] = True
                    shapeInfo[shapesCount]['LineRGB'] = int(Line.ForeColor.__str__())
                    shapeInfo[shapesCount]['Weight'] = Line.Weight
                    shapeInfo[shapesCount]['DashStyle'] = Line.DashStyle
        except:
            traceback.print_exc(limit=1)
    def exportShapeInfoTable(self,shapeInfo,Shape,shapesCount,SlideIndex):
        # 表格
        Table = Shape.Table
        shapeInfo[shapesCount]['Rows'] = Table.Rows.Count
        shapeInfo[shapesCount]['Columns'] = Table.Columns.Count
        try:
            shapeInfo[shapesCount]['StyleId'] = Table.Style.Id
        except:
            shapeInfo[shapesCount]['StyleId'] = None
        mxheader = np.zeros(shape=[Table.Columns.Count])
        mxheader = mxheader.astype(np.str)
        if Table.Rows.Count > 1:
            mx = np.zeros(shape=[Table.Rows.Count-1, Table.Columns.Count])
            mx = mx.astype(np.str)
        for Row in range(1,Table.Rows.Count+1):
            for Column in range(1,Table.Columns.Count+1):
                if Row ==1:
                    mxheader[Column-1] = Table.Cell(Row,Column).Shape.TextFrame.TextRange.Text
                else:
                    mx[Row-2,Column-1] = Table.Cell(Row,Column).Shape.TextFrame.TextRange.Text
        if Table.Rows.Count == 1:
            df = pd.DataFrame(columns=mxheader)
        else:
            df = pd.DataFrame(data=mx,columns=mxheader)
        shapeInfo[shapesCount]['TableData'] = df
    def exportShapeInfoChart(self,shapeInfo,Shape,shapesCount,SlideIndex):
        # 图表
        Chart = Shape.Chart
        shapeInfo[shapesCount]['ChartType'] = Chart.ChartType
        shapeInfo[shapesCount]['ChartTitle'] = Chart.ChartTitle.Text if Chart.HasTitle else False
        shapeInfo[shapesCount]['Legend'] = True if Chart.HasLegend else False
        shapeInfo[shapesCount]['DataTable'] = True if Chart.HasDataTable else False
        shapeInfo[shapesCount]['SeriesCollectionCount'] = Chart.SeriesCollection().Count
        # 坐标轴
        attrs = {'PrimaryCategoryAxis':(1,1),'PrimaryValueAxis':(2,1),'SecondaryCategoryAxis':(1,2),'SecondaryValueAxis':(2,2)}
        for attr in attrs:
            shapeInfo[shapesCount][attr] = False
            try:
                Chart.Axes(*attrs[attr]).Select()
                shapeInfo[shapesCount][attr] = True
            except:
                pass
        # 网格线
        shapeInfo[shapesCount]['PrimaryCategoryGridLinesMajor'] = False
        try:
            Chart.Axes(1,1).MajorGridlines.Select()
            shapeInfo[shapesCount]['PrimaryCategoryGridLinesMajor'] = True
        except:
            pass   
        shapeInfo[shapesCount]['PrimaryCategoryGridLinesMinorMajor'] = False
        try:
            Chart.Axes(1,1).MinorGridlines.Select()
            shapeInfo[shapesCount]['PrimaryCategoryGridLinesMinorMajor'] = True
        except:
            pass
        shapeInfo[shapesCount]['PrimaryValueGridLinesMajor'] = False
        try:
            Chart.Axes(2,1).MajorGridlines.Select()
            shapeInfo[shapesCount]['PrimaryValueGridLinesMajor'] = True
        except:
            pass  
        shapeInfo[shapesCount]['PrimaryValueGridLinesMinorMajor'] = False
        try:
            Chart.Axes(2,1).MinorGridlines.Select()
            shapeInfo[shapesCount]['PrimaryValueGridLinesMinorMajor'] = True
        except:
            pass        
        # Label
        shapeInfo[shapesCount]['ChartLabelSetting'] = self.exportLabelChart(Chart)
        # Series, Points
        shapeInfo[shapesCount]['ChartSeriesFill'] = self.exportFillChart(Chart)
        shapeInfo[shapesCount]['ChartSeriesLine'] = self.exportLineChart(Chart)
        shapeInfo[shapesCount]['ChartPointsFill'] = self.exportFillChartPoint(Chart)
        shapeInfo[shapesCount]['ChartPointsLine'] = self.exportLineChartPoint(Chart)
        ChartData = Chart.ChartData
        tries = self.tries
        while tries:
            try:
                ChartData.Activate()
                Workbook = ChartData.Workbook
                Worksheet = Workbook.Activesheet
                break
            except:
                tries -= 1
                time.sleep(self.seconds)
        tempf = os.path.join(os.path.dirname(self.presInfo['FileName']), 'temp%s.csv'%time.strftime("%Y%m%d%H%M%S", time.localtime()))
        try:
            Worksheet.SaveAs(tempf,FileFormat=6) # csv,sep=',',encoding='gbk'
            df = pd.read_csv(tempf,sep=',',encoding=self.encoding)
            df = df.set_index(df.columns[0])
            Workbook.Close()
            os.unlink(tempf)
            shapeInfo[shapesCount]['PlotBy'] = Chart.PlotBy # 1:xlRows, 2:xlColumns
            shapeInfo[shapesCount]['ChartData'] = df
            Chart.Refresh()
            return True
        except:
            print('Error in %s!\nPlease MustMustMust Delete It!'%SlideIndex)
            traceback.print_exc()
            return False
    def exportLabelChart(self,Chart=None):
        if not Chart:
            Chart = self.pptShape.Chart
        if not Chart:
            return False
        LabelChartDict = defaultdict(dict)
        for i in range(1, Chart.SeriesCollection().Count+1):
            LabelChartDict[i]['HasDataLabels'] = None
            attrs = ['ShowBubbleSize','ShowCategoryName','ShowLegendKey','ShowPercentage','ShowRange','ShowSeriesName','ShowValue',
                     'Position','NumberFormatLocal']
            fontattrs = ['Name','Size','Color']
            for attr in attrs:
                LabelChartDict[i][attr] = None
            for fontattr in fontattrs:
                LabelChartDict[i][fontattr] = None   
            if Chart.SeriesCollection(i).HasDataLabels:
                LabelChartDict[i]['HasDataLabels'] = True
                for attr in attrs:
                    try:
                        LabelChartDict[i][attr] = getattr(Chart.SeriesCollection(i).DataLabels(),attr)
                    except:
                        pass #traceback.print_exc(limit=1)
                fontattrs = ['Name','Size','Color']
                for fontattr in fontattrs:         
                    try:
                        LabelChartDict[i][fontattr] = getattr(Chart.SeriesCollection(i).DataLabels().Font,fontattr)
                    except:
                        pass #traceback.print_exc(limit=1)             
        df = pd.DataFrame(LabelChartDict).transpose()
        columns = ['HasDataLabels','ShowBubbleSize','ShowCategoryName','ShowLegendKey','ShowPercentage','ShowRange',
                   'ShowSeriesName','ShowValue','Position','NumberFormatLocal','Name','Size','Color']
        df = df[columns]
        return df
    def exportFillChart(self,Chart=None):
        if not Chart:
            Chart = self.pptShape.Chart
        if not Chart:
            return False
        FillChartDict = defaultdict(dict)
        for i in range(1, Chart.SeriesCollection().Count+1):
            FillChartDict[i]['Fill'] = None
            FillChartDict[i]['FillRGB'] = None
            try:
                Fill = Chart.SeriesCollection(i).Format.Fill
                if Fill.Visible:
                    FillChartDict[i]['Fill'] = True
                    FillChartDict[i]['FillRGB'] = int(Fill.ForeColor.__str__())
            except:
                traceback.print_exc(limit=1)
        df = pd.DataFrame(FillChartDict).transpose() 
        return df
    def exportLineChart(self,Chart=None):
        if not Chart:
            Chart = self.pptShape.Chart
        if not Chart:
            return False
        FillChartDict = defaultdict(dict)
        for i in range(1, Chart.SeriesCollection().Count+1):
            FillChartDict[i]['Line'] = None
            FillChartDict[i]['LineWeight'] = None
            FillChartDict[i]['LineDashStyle'] = None
            FillChartDict[i]['LineRGB'] = None
            try:
                Line = Chart.SeriesCollection(i).Format.Line
                if Line.Visible:
                    FillChartDict[i]['Line'] = True
                    FillChartDict[i]['LineWeight'] = Line.Weight if Line.Weight > 0 else 0
                    FillChartDict[i]['LineDashStyle'] = Line.DashStyle
                    FillChartDict[i]['LineRGB'] = int(Line.ForeColor.__str__())
            except:
                traceback.print_exc(limit=1)
        df = pd.DataFrame(FillChartDict).transpose() 
        return df
    def exportFillChartPoint(self,Chart=None):
        if not Chart:
            Chart = self.pptShape.Chart
        if not Chart:
            return False
        FillChartDict = defaultdict(dict)
        if Chart.SeriesCollection().Count > 0:
            SeriesCollection = Chart.SeriesCollection(1)
            for i in range(1, SeriesCollection.Points().Count+1):
                FillChartDict[i]['FillPoint'] = None
                FillChartDict[i]['FillPointRGB'] = None               
                try:
                    Fill = SeriesCollection.Points(i).Format.Fill
                    if Fill.Visible:
                        FillChartDict[i]['FillPoint'] = True
                        FillChartDict[i]['FillPointRGB'] = int(Fill.ForeColor.__str__())
                except:
                    traceback.print_exc(limit=1)
        df = pd.DataFrame(FillChartDict).transpose() 
        return df
    def exportLineChartPoint(self,Chart=None):
        if not Chart:
            Chart = self.pptShape.Chart
        if not Chart:
            return False
        FillChartDict = defaultdict(dict)
        if Chart.SeriesCollection().Count > 0:
            SeriesCollection = Chart.SeriesCollection(1)
            for i in range(1, SeriesCollection.Points().Count+1):
                FillChartDict[i]['LinePoint'] = None
                FillChartDict[i]['LinePointWeight'] = None
                FillChartDict[i]['LinePointDashStyle'] = None
                FillChartDict[i]['LinePointRGB'] = None
                try:
                    Line = SeriesCollection.Points(i).Format.Line
                    if Line.Visible:
                        FillChartDict[i]['LinePoint'] = True
                        FillChartDict[i]['LinePointWeight'] = Line.Weight if Line.Weight > 0 else 0
                        FillChartDict[i]['LinePointDashStyle'] = Line.DashStyle
                        FillChartDict[i]['LinePointRGB'] = int(Line.ForeColor.__str__())
                except:
                    traceback.print_exc(limit=1)
        df = pd.DataFrame(FillChartDict).transpose() 
        return df
    def exportShapeInfoSmartArt(self,shapeInfo,Shape,shapesCount,SlideIndex):
        # SmartArt
        pass
    def exportShapeInfoText(self,shapeInfo,Shape,shapesCount,SlideIndex):
        TextFrame = Shape.TextFrame
        Font = TextFrame.TextRange.Font
        # 对齐信息: 'HorizontalAnchor' 1:左对齐，2:右对齐;'VerticalAnchor'1:垂直上，3:垂直中，4:垂直下   
        attrs = ['AutoSize','VerticalAnchor']  
        for attr in attrs:
            shapeInfo[shapesCount][attr] = getattr(TextFrame,attr)
        shapeInfo[shapesCount]['HorizontalAnchor'] = TextFrame.TextRange.ParagraphFormat.Alignment
        # 四边距
        attrs = ['MarginLeft','MarginTop','MarginRight','MarginBottom']
        for attr in attrs:
            shapeInfo[shapesCount][attr] = getattr(TextFrame,attr)
        # 文本信息
        attrs = ['Name','Size','Bold','Italic','Underline'] 
        shapeInfo[shapesCount]['Text'] = TextFrame.TextRange.Text
        shapeInfo[shapesCount]['ColorRGB'] = int(Font.Color.__str__())
        for attr in attrs:
            shapeInfo[shapesCount][attr] = getattr(Font,attr) 
        # 段落标记
        shapeInfo[shapesCount]['BulletType'] = TextFrame.TextRange.ParagraphFormat.Bullet.Type
        if shapeInfo[shapesCount]['BulletType'] in [0,-2,3]:
            shapeInfo[shapesCount]['BulletType'] = 0
        shapeInfo[shapesCount]['BulletCharacter'] = TextFrame.TextRange.ParagraphFormat.Bullet.Character
    def exportShapeInfoShape(self,shapeInfo,Shape,shapesCount,SlideIndex):
        # Shape
        try:
            if shapeInfo[shapesCount]['ShapeType'] == 'msoAutoShape':
                shapeInfo[shapesCount]['Source'] = None
            elif shapeInfo[shapesCount]['ShapeType'] == 'msoPicture':
                dirname =  self.presInfo['FileName']
                dirname = dirname[:dirname.rfind('.')]
                dirname = os.path.join(dirname,'pictures')
                if os.path.exists(dirname):
                    os.mkdir(dirname)
                Source = os.path.join(dirname,'%s.gif'%shapesCount)
                Shape.Export(Source,Filter=0)
                shapeInfo[shapesCount]['Source'] = Source
            else:
                shapeInfo[shapesCount]['Source'] = None
            return True
        except:
            print('Error in %s!\nPlease MustMustMust Delete It!'%SlideIndex)
            traceback.print_exc()
            return False   
    def exportPPTtoExcel(self):
        return MSOPowerpointFunc.exportPPTtoExcel(self)
    
    def createPPT(self,filename):
        createFuncs = {'Summary':self.createSummary,
                       'Table':self.createTable,
                       'Chart':self.createChart,
                       'SmartArt':self.createSmartArt,
                       'Text':self.createText,
                       'Shape':self.createShape,
                      }
        #filename = MSOPowerpointFunc.getExcelTextFile(filename)
        dfs = MSOPowerpointFunc.sheetsToDfs(filename)
        records = []
        # Summary
        createfunc = 'Summary'
        newfilename,recordfilename  = createFuncs[createfunc](createfunc,dfs,records)
        for createfunc in createFuncs:
            if createfunc in dfs:
                createFuncs[createfunc](createfunc,dfs,newfilename,records) 
        with open(recordfilename,'w') as f:
            for record in records:
                f.write('%s\n'%record)
        if records:
            os.startfile(recordfilename)
        self.savePPT(newfilename)
        infomation = 'PPT Create >>> Powerpoint:%s is created!'%(os.path.basename(newfilename))
        for i in infomation:
            print(i,end='')
            time.sleep(self.seconds/100)
        print('')
        print('<'*40)
        return newfilename
    def createSummary(self,createfunc,dfs,records):
        '''
        Create PPT Summary
        '''
        print('%s is creating!'%(createfunc))
        credic = dfs.pop(createfunc).transpose()[0].to_dict()
        if os.path.exists('template.pptx'):
            self.openPPT('template.pptx')
        else:
            self.openPPT(credic['FileName'])
        newfilepath = os.path.dirname(credic['FileName'])
        newfilename = os.path.basename(credic['FileName'])
        newfilename = '%s%s.pptx'%(newfilename[:newfilename.rfind('.')],time.strftime("%Y%m%d%H%M%S", time.localtime()))
        newfilename = os.path.join(newfilepath,newfilename)
        self.savePPT(newfilename)
        recordfilename = '%s.txt'%newfilename[:newfilename.rfind('.')]
        recordfilename = os.path.join(newfilepath,recordfilename)
        # 生成页面
        SlidesCounts = self.presInfo['SlidesCount']
        for i in range(SlidesCounts):
            self.delSlide(Index=1)
        SlidesCounts = credic['SlidesCount']
        for i in range(SlidesCounts):
            self.addSlide()
        self.setupPage(SlideWidth=MSOPowerpointFunc.funcPixelInch(credic['Width']),SlideHeight=MSOPowerpointFunc.funcPixelInch(credic['Height']))   
        infomation = 'PPT Create >>> Powerpoint:%s is creating!'%(os.path.basename(credic['FileName']))
        for i in infomation:
            print(i,end='')
            time.sleep(self.seconds/100)
        print('')
        print('>'*40)
        return newfilename,recordfilename 
    def createTable(self,createfunc,dfs,newfilename,records):
        '''
        Create PPT Table
        '''
        print('%s is creating!'%(createfunc))
        df = dfs.pop(createfunc)
        for i in df.index.tolist():
            credic = df.transpose()[i].to_dict()
            SlideIndex = credic['SlideIndex']
            self.selectSlide(SlideIndex)
            dftable = dfs.pop('%s_%s'%(createfunc,i)).replace(np.nan,'')
            NumRows = credic['Rows']
            NumColumns = credic['Columns']
            Left = credic['Left']
            Top = credic['Top']
            Width = credic['Width']
            Height = credic['Height']
            self.setTableFromPandas(df=dftable,Left=Left,Top=Top,CellWidth=Width/NumColumns,Initialized=True)
            StyleId = credic['StyleId']
            StyleName = MSOPowerpointFunc.funcLookupKey(StyleId,self.msoTableStype)
            if not self.setTableStyle(StyleName=StyleName):
                try:
                    Fill = credic['Fill']
                    if Fill:
                        r,g,b = MSOPowerpointFunc.getRGBback(credic['FillRGB'])
                        self.fillTable(r=r,g=g,b=b)
                except:
                    pass
                try:
                    Line = credic['Line']
                    if Line:
                        r,g,b = MSOPowerpointFunc.getRGBback(credic['LineRGB'])
                        Weight = credic['Weight']
                        DashStyle = credic['DashStyle']
                        self.fillTableBorder(r=r,g=g,b=b,Weight=Weight,DashStyle=DashStyle)
                except:
                    pass                                
            self.createHyperlink(self.pptShape,credic)
            print('PPT Create >>> SlideIndex:%s,%s,%s created!'%(SlideIndex,createfunc,i))
    def createChart(self,createfunc,dfs,newfilename,records):
        '''
        Create PPT Chart
        '''
        print('%s is creating!'%(createfunc))
        df = dfs.pop(createfunc)
        for i in df.index.tolist():
            credic = df.transpose()[i].to_dict()
            SlideIndex = credic['SlideIndex']
            self.selectSlide(SlideIndex)
            dfchart = dfs.pop('%s_%s'%(createfunc,i)).replace(np.nan,'')
            Left = credic['Left']
            Top = credic['Top']
            Width = credic['Width']
            Height = credic['Height']
            xlChartType = credic['ChartType']
            PlotBy = credic['PlotBy']
            self.addChart(xlChartType=xlChartType,Left=Left,Top=Top,Width=Width,Height=Height,PlotBy=PlotBy,Initialized=True)
            self.setChartDataFromPandas(df=dfchart,PlotBy=PlotBy)
            ChartTitle = credic['ChartTitle']
            if ChartTitle:
                self.setChartTitle(ChartTitle=ChartTitle)
            else:
                self.setChartTitle(ChartTitle=False)
            Legend = credic['Legend']
            if Legend:
                self.setChartLegend(Legend=True)  
            else:
                self.setChartLegend(Legend=False)  
            DataTable = credic['DataTable']
            if DataTable:
                self.showChartDataTable(HasDataTable=True)
            else:
                self.showChartDataTable(HasDataTable=False) 
            # Label
            self.createChartLabel(createfunc,dfs,newfilename,records,SlideIndex,dfchart,i)
            SeriesCollection = credic['SeriesCollectionCount']            
            # to be fixed
            if SeriesCollection > 1:
                self.createChartFillLine(createfunc,dfs,newfilename,records,SlideIndex,dfchart,i)
            elif SeriesCollection == 1:
                self.createChartPointFillLine(createfunc,dfs,newfilename,records,SlideIndex,dfchart,i)
            # 坐标轴    
            attrs = {'PrimaryCategoryAxis':(1,1),'PrimaryValueAxis':(2,1),'SecondaryCategoryAxis':(1,2),'SecondaryValueAxis':(2,2)}
            _attrs = {'PrimaryCategoryAxis':349,'PrimaryValueAxis':353,'SecondaryCategoryAxis':359,'SecondaryValueAxis':363}
            for attr in attrs:
                if credic[attr]:
                    print('Axis %s Created'%attr)
                    try:
                        self.pptChart.SetElement(_attrs[attr])
                    except:
                        pass
                else:
                    print('Axis %s Deleted'%attr)
                    try:
                        self.pptChart.Axes(*attrs[attr]).Delete()
                    except:
                        pass
            # 网格线
            PrimaryCategoryGridLinesMajor = credic['PrimaryCategoryGridLinesMajor']
            if PrimaryCategoryGridLinesMajor:
                try:
                    self.pptChart.SetElement(334)
                except:
                    pass
            else:
                try:
                    self.pptChart.Axes(1,1).MajorGridlines.Delete()
                except:
                    pass              
            PrimaryCategoryGridLinesMinorMajor = credic['PrimaryCategoryGridLinesMinorMajor']
            if PrimaryCategoryGridLinesMinorMajor:
                try:
                    self.pptChart.SetElement(335)
                except:
                    pass
            else:
                try:
                    self.pptChart.Axes(1,1).MinorGridlines.Delete()
                except:
                    pass  
            PrimaryValueGridLinesMajor = credic['PrimaryValueGridLinesMajor']
            if PrimaryValueGridLinesMajor:
                try:
                    self.pptChart.SetElement(330)
                except:
                    pass
            else:
                try:
                    self.pptChart.Axes(2,1).MajorGridlines.Delete()
                except:
                    pass              
            PrimaryValueGridLinesMinorMajor = credic['PrimaryValueGridLinesMinorMajor']
            if PrimaryCategoryGridLinesMinorMajor:
                try:
                    self.pptChart.SetElement(331)
                except:
                    pass
            else:
                try:
                    self.pptChart.Axes(2,1).MinorGridlines.Delete()
                except:
                    pass     
            self.createHyperlink(self.pptShape,credic)
            self.pptChart.Refresh()
            print('PPT Create >>> SlideIndex:%s,%s,%s created!'%(SlideIndex,createfunc,i))     
    def createChartLabel(self,createfunc,dfs,newfilename,records,SlideIndex,dfchart,i):
        fl = 'LabelSetting'
        df = dfs.pop('%s%s_%s'%(createfunc,fl,i)).replace(np.nan,'')
        attrs = ['ShowBubbleSize','ShowCategoryName','ShowLegendKey','ShowPercentage','ShowRange','ShowSeriesName','ShowValue',
                 'Position','NumberFormatLocal']
        fontattrs = ['Name','Size','Color']
        for j in df.index.tolist():
            credic = df.transpose()[j].to_dict()
            DataLabels = credic['HasDataLabels']
            if DataLabels:
                Chart = self.pptChart
                Chart.SeriesCollection(j).ApplyDataLabels()
                for attr in attrs:
                    try:
                        setattr(Chart.SeriesCollection(j).DataLabels(),attr,credic[attr])
                    except:
                        records.append('SlideIndex:%s,%s,Series:%s,%s,%s,failed!'%(SlideIndex,createfunc,j,fl,attr))
                for fontattr in fontattrs:
                    try:
                        setattr(Chart.SeriesCollection(j).DataLabels().Font,fontattr,credic[fontattr])
                    except:
                        records.append('SlideIndex:%s,%s,Series:%s,%s,%s,failed!'%(SlideIndex,createfunc,j,fl,fontattrs))                     
        print('PPT Create >>> SlideIndex:%s,%s,Series:%s,%s,created!'%(SlideIndex,createfunc,j,fl))
    def createChartFillLine(self,createfunc,dfs,newfilename,records,SlideIndex,dfchart,i):
        fl = 'SeriesFill'
        df = dfs.pop('%s%s_%s'%(createfunc,fl,i)).replace(np.nan,'')
        for j in df.index.tolist():
            credic = df.transpose()[j].to_dict()
            Fill = credic['Fill']
            try:
                if Fill:
                    r,g,b = MSOPowerpointFunc.getRGBback(credic['FillRGB'])
                    self.fillChart(Index=j,r=r,g=g,b=b)
            except:
                records.append('SlideIndex:%s,%s,Series:%s,%s,failed!'%(SlideIndex,createfunc,j,fl))
                traceback.print_exc(limit=1) 
        print('PPT Create >>> SlideIndex:%s,%s,Series:%s,%s,created!'%(SlideIndex,createfunc,j,fl))

        fl = 'SeriesLine'
        df = dfs.pop('%s%s_%s'%(createfunc,fl,i)).replace(np.nan,'')
        for j in df.index.tolist():
            credic = df.transpose()[j].to_dict()
            Line = credic['Line']
            try:        
                if Line:
                    r,g,b = MSOPowerpointFunc.getRGBback(credic['LineRGB'])
                    DashStyle = credic['LineDashStyle']
                    Weight = credic['LineWeight']
                    self.lineChart(Index=j,DashStyle=DashStyle,Weight=Weight,r=r,g=g,b=b)
            except:
                records.append('SlideIndex:%s,%s,Series:%s,%s,failed!'%(SlideIndex,createfunc,j,fl))
                traceback.print_exc(limit=1)
            print('PPT Create >>> SlideIndex:%s,%s,Series:%s,%s,created!'%(SlideIndex,createfunc,j,fl))  
    def createChartPointFillLine(self,createfunc,dfs,newfilename,records,SlideIndex,dfchart,i):
        fl = 'PointsFill'
        df = dfs.pop('%s%s_%s'%(createfunc,fl,i)).replace(np.nan,'')
        for j in df.index.tolist():
            credic = df.transpose()[j].to_dict()
            Fill = credic['FillPoint']
            try:
                if Fill:
                    r,g,b = MSOPowerpointFunc.getRGBback(credic['FillPointRGB'])
                    self.fillChartPoint(Index=1,Point=j,r=r,g=g,b=b)
            except:
                records.append('SlideIndex:%s,%s,Series:%s,Point:%s,%s,failed!'%(SlideIndex,createfunc,1,j,fl))
                traceback.print_exc(limit=1)
        print('PPT Create >>> SlideIndex:%s,%s,Series:%s,Point:%s,%s,created!'%(SlideIndex,createfunc,1,j,fl))

        fl = 'PointsLine'
        df = dfs.pop('%s%s_%s'%(createfunc,fl,i)).replace(np.nan,'')
        for j in df.index.tolist():
            credic = df.transpose()[j].to_dict()
            Line = credic['LinePoint']
            try:
                if Line:
                    r,g,b = MSOPowerpointFunc.getRGBback(credic['LinePointRGB'])
                    DashStyle = credic['LinePointDashStyle']
                    Weight = credic['LinePointWeight']
                    self.lineChartPoint(Index=1,Point=j,r=r,g=g,b=b,DashStyle=DashStyle,Weight=Weight)
            except:
                records.append('SlideIndex:%s,%s,Series:%s,Point:%s,%s,failed!'%(SlideIndex,createfunc,1,j,fl))
                traceback.print_exc(limit=1)
        print('PPT Create >>> SlideIndex:%s,%s,Series:%s,Point:%s,%s,created!'%(SlideIndex,createfunc,1,j,fl))
    def createSmartArt(self,createfunc,dfs,newfilename,records):
        '''
        Create PPT Table
        '''
        print('%s is creating!'%(createfunc))
        df = dfs.pop(createfunc).replace(np.nan,'')
        for i in df.index.tolist():
            credic = df.transpose()[i].to_dict()
            SlideIndex = credic['SlideIndex']
            ShapeType=credic['ShapeType']
            records.append('PPT Create >>> SlideIndex:%s,%s,%s need to be created manually!'%(SlideIndex,createfunc,ShapeType))
            print('PPT Create >>> SlideIndex:%s,%s,%s need to be created manually!'%(SlideIndex,createfunc,ShapeType))
    def createShape(self,createfunc,dfs,newfilename,records):
        '''
        Create PPT Shape
        '''
        print('%s is creating!'%(createfunc))
        df = dfs.pop(createfunc).replace(np.nan,'')
        for i in df.index.tolist():
            credic = df.transpose()[i].to_dict()
            SlideIndex = credic['SlideIndex']
            self.selectSlide(SlideIndex)
            Left = credic['Left']
            Top = credic['Top']
            Width = credic['Width']
            Height = credic['Height']
            ShapeType = credic['ShapeType']
            if ShapeType == 'msoAutoShape':
                AutoShapeType = credic['AutoShapeType']
                AutoShapeType = self.msoAutoShapeType.get(AutoShapeType,None)
                AutoShapeType = 1 if AutoShapeType in [138] else AutoShapeType
                if AutoShapeType and AutoShapeType != -2:
                    self.addShape(AutoShapeType=AutoShapeType,Left=Left,Top=Top,Width=Width,Height=Height)
                    Fill = credic['Fill']
                    if Fill:
                        r,g,b = MSOPowerpointFunc.getRGBback(credic['FillRGB'])
                        self.fillShape(r=r,g=g,b=b)
                    Line = credic['Line']
                    if Line:
                        r,g,b = MSOPowerpointFunc.getRGBback(credic['LineRGB'])
                        Weight = credic['Weight']
                        DashStyle = credic['DashStyle']
                        self.lineShape(Weight=Weight,DashStyle=DashStyle,r=r,g=g,b=b)
                    self.createHyperlink(self.pptShape,credic)
                    print('PPT Create >>> SlideIndex:%s,%s,%s created!'%(SlideIndex,createfunc,i))
                else:
                    records.append('PPT Create >>> SlideIndex:%s,%s,%s need to be created manually!'%(SlideIndex,createfunc,ShapeType))
                    print('PPT Create >>> SlideIndex:%s,%s,%s need to be created manually!'%(SlideIndex,createfunc,ShapeType))
            elif ShapeType == 'msoPicture':
                Source = credic['Source']
                self.addPicture(Picture=Source,LinkToFile=0,SaveWithDocument=-1,Left=Left,Top=Top,Width=Width,Height=Height,AdjustDpi=False)                
                self.createHyperlink(self.pptShape,credic)
                print('PPT Create >>> SlideIndex:%s,%s,%s created!'%(SlideIndex,createfunc,i))
            else:
                records.append('PPT Create >>> SlideIndex:%s,%s,%s need to be created manually!'%(SlideIndex,createfunc,ShapeType))
                print('PPT Create >>> SlideIndex:%s,%s,%s need to be created manually!'%(SlideIndex,createfunc,ShapeType))
    def createText(self,createfunc,dfs,newfilename,records):
        '''
        Create PPT Text
        '''
        print('%s is creating!'%(createfunc))
        df = dfs.pop(createfunc).replace(np.nan,'')
        for i in df.index.tolist():
            credic = df.transpose()[i].to_dict()
            SlideIndex = credic['SlideIndex']
            self.selectSlide(SlideIndex)
            Left = credic['Left']
            Top = credic['Top']
            Width = credic['Width']
            Height = credic['Height']
            AutoShapeType = credic['AutoShapeType']
            AutoShapeType = self.msoAutoShapeType.get(AutoShapeType,None)
            AutoShapeType = 1 if AutoShapeType in [138] else AutoShapeType
            if AutoShapeType:
                self.addShape(AutoShapeType=AutoShapeType,Left=Left,Top=Top,Width=Width,Height=Height)
            else:
                self.addLabel(Orientation=1,Left=Left,Top=Top,Width=Width,Height=Height,Initialized=False)
            AutoSize = credic['AutoSize']
            HorizontalAnchor = credic['HorizontalAnchor']
            VerticalAnchor = credic['VerticalAnchor']
            MarginLeft = credic['MarginLeft']
            MarginTop = credic['MarginTop']
            MarginRight = credic['MarginRight']
            MarginBottom = credic['MarginBottom']
            self.setTextFrame(AutoSize=AutoSize,HorizontalAnchor=HorizontalAnchor,VerticalAnchor=VerticalAnchor,
                              MarginLeft=MarginLeft,MarginTop=MarginTop,MarginRight=MarginRight,MarginBottom=MarginBottom)
            Text = credic['Text']
            self.addText(Text=Text)     
            r,g,b = MSOPowerpointFunc.getRGBback(credic['ColorRGB'])
            Name = credic['Name']
            Size = credic['Size']
            Bold = credic['Bold']
            Italic = credic['Italic']
            Underline = credic['Underline']
            self.fillText(r=r,g=g,b=b,Name=Name,Size=Size,Bold=Bold,Italic=Italic,Underline=Underline)
            Fill = credic['Fill']
            if Fill:
                r,g,b = MSOPowerpointFunc.getRGBback(credic['FillRGB'])
                self.fillShape(r=r,g=g,b=b)
            Line = credic['Line']
            if Line:
                r,g,b = MSOPowerpointFunc.getRGBback(credic['LineRGB'])
                Weight = credic['Weight']
                DashStyle = credic['DashStyle']
                self.lineShape(Weight=Weight,DashStyle=DashStyle,r=r,g=g,b=b)    
            # 段落标记
            BulletType = credic['BulletType']
            if BulletType:
                self.pptShape.TextFrame.TextRange.ParagraphFormat.Bullet.Type = BulletType
                if BulletType == 1:
                    self.pptShape.TextFrame.TextRange.ParagraphFormat.Bullet.Character = credic['BulletCharacter']
            self.createHyperlink(self.pptShape,credic)
            print('PPT Create >>> SlideIndex:%s,%s,%s created!'%(SlideIndex,createfunc,i))      
    def createHyperlink(self,Shape,credic):
        HyperLink = credic.get('HyperLink',None)
        Address = credic.get('Address',None)
        SubAddress = credic.get('SubAddress',None)
        if HyperLink:
            self.hyperlinkShape(Shape,Address,SubAddress)

    


# #### Test_runExportPPTtoExcel

# In[6]:


if __name__ == '__main__':
    ## export
    FileName = r'C:/SmithYe/PythonProject3/OfficeApi/MSOPowerpoint/testfiles/template_shape.pptx'
    MP = MSOPowerpoint(FileName=FileName,WithWindow=1)
    MP.exportShapesInfo()
    print(MP.exportPPTtoExcel())


# #### Test_CreatePPT

# In[13]:


if __name__ == '__main__':
    # create
    MP = MSOPowerpoint(Blank=None)
    filename = r'C:/SmithYe/PythonProject3/OfficeApi/MSOPowerpoint/testfiles/template_Chart.xlsx'
    MP.createPPT(filename)


# #### Test_All

# In[14]:


if __name__ == '__main__': 
    ## initial
    FileName = r'./testfiles/template.pptx'
    MP = MSOPowerpoint(FileName=FileName,WithWindow=1)
    def test():
        print(MP.presInfo)
        ## select Slide
        Index = 1
        MP.selectSlide(Index=Index)
        ## shapesInfo
        MP.getShapesText()
        print(MP.shapesText)
        MP.exportShapesInfo()
        print(MP.shapesInfo)
        ## Select Shape
        Index = 1
        MP.selectShape(Index=Index)
        SlideIndex = 2
        ShapeName = 'TextBox 1'
        MP.selectShape(SlideIndex=SlideIndex,ShapeName=ShapeName) 
        ## Saveas
        newFileName = 'new.pptx'
        MP.savePPT(newFileName=newFileName,Close=True)          
    def test_slide():
        ## 设置页面
        MP.setupPage()
        ## 增加一页
        MP.addSlide()
        ## 删除一页
        MP.delSlide()  
    def test_shape():
        FileName = r'./testfiles/template.pptx'
        MP = MSOPowerpoint(FileName=FileName,WithWindow=1)
        MP.addShape(Height=300,Width=100)
        MP.fillShape(r=0,g=0,b=255)
        MP.lineShape(r=255,g=0,b=0,Weight=3)
        MP.hyperlinkShape(Address=r'https://www.baidu.com')   
    def test_text():
        FileName = None
        MP = MSOPowerpoint(FileName=FileName,WithWindow=1,GetDpi=0,Blank=1)  
        ## 增加文本框
        Orientation,Left,Top,Width,Height = 1,0,0,100,10
        Text='Solidjoker says:\nHello World!'
        MP.addLabel(Text=Text,Orientation=Orientation,Left=Left,Top=Top,Width=Width,Height=Height)
        MP.setTextFrame(AutoSize=1)
        MP.fillText(r=0,g=255,b=0,Size=50)
        ## Add Hyperlink
        Address = 'https://www.baidu.com'
        MP.hyperlinkShape(Address=Address)
        ## Set Text Characters Start and Length
        MP.setTextCharacters(Start=3,Length=8,Size=100,b=255)
    def test_picture():
        ## add picture
        Picture=r'./testfiles/test.png'
        Left=0
        Top=0
        MP.addPicture(Picture=Picture,Left=Left,Top=Top,AdjustDpi=1)
    def test_line():
        ## add Line
        MP.addLine()
        MP.lineShape()
    def test_table():
        ## table: add, without AutoAdjust
        NumRows,NumColumns,Left,Top,Width,Height=5,4,0,0,400,200
        PT.addTable(NumRows,NumColumns,Left,Top,Width,Height,Initialized=0)
        PT.setTableBorder(Weight=3,DashStyle=2)
        PT.setTableColor(r=0,g=128,b=128)
        PT.setTableWidthHeight(ColRow='row',Index=1,Point=50)
        ## table: add, with AutoAdjust
        #NumRows,NumColumns,Left,Top,Width,Height=5,4,0,0,400,200        
        #PT.addTable(NumRows,NumColumns,Left,Top,Width,Height,Initialized=1)
        ## table: title
        PT.setTableWidthHeight()
        ## table: text
        HorizontalAnchor = 2
        VerticalAnchor = 3
        Size = 10
        Name = '微软雅黑'
        Bold = 1
        ColorName = 'red'
        PT.setTableTextFormat(HorizontalAnchor=HorizontalAnchor,VerticalAnchor=VerticalAnchor,
                              Size=Size,Name=Name,Bold=Bold,ColorName='red')
        ## table，create from pandas
        filename = r'./testfiles/excel.xlsx'
        sheet_name='Sheet1'
        df = pd.read_excel(filename,sheet_name=sheet_name)
        MP.setTableFromPandas(df,Initialized=True)  
        print(PT.pptShape.Name)
    def test_chart():
        # add chart
        xlColumnClustered = MP.xlChartType['xlColumnClustered']
        MP.addChart(xlColumnClustered,Initialized=False)
        # change chart type
        Index = None
        xlLine = MP.xlChartType['xlLine']
        MP.changeChartType(xlChartType=xlLine,Index=Index)
        # set chart title
        MP.setChartTitle(ChartTitle='Title',Name='微软雅黑',Size=16)
        # set chart legend
        MP.setChartLegend(Legend=True,Name='微软雅黑',Size=16,Italic=True)
        # set chart axes
        Axis=None
        AxisGroup=1
        Size=10
        Name='微软雅黑'
        NumberFormatLocal='#,##'
        MP.setChartAxes(Axis=Axis,AxisGroup=AxisGroup,Size=Size,Name=Name,NumberFormatLocal=NumberFormatLocal)
        # change chart axisgroup
        Index=2
        AxisGroup=2
        MP.changeChartAxisGroup(Index=Index,AxisGroup=AxisGroup)
        # set chart secondary axes 
        Axis=2
        AxisGroup=2
        Size=20
        Name='微软雅黑'
        NumberFormatLocal='#,##'
        MP.setChartAxes(Axis=Axis,AxisGroup=AxisGroup,Size=Size,Name=Name,NumberFormatLocal=NumberFormatLocal)
        # Set DataLabel
        Index = None
        Name='微软雅黑'
        Size=10
        NumberFormatLocal='#,##'
        MP.setChartDataLabel(Index=Index,Size=Size,Name=Name,NumberFormatLocal=NumberFormatLocal)
        # show chart data label
        Index = 1
        MP.showChartDataLabel(Index=Index,ShowCategoryName=True)  
        # replace chart data label
        Index = 1
        Replacements = ['a','b','c']
        MP.replaceChartDataLabel(Index=Index,Replacements=Replacements)
        # show chart datatable
        HasDataTable = True
        MP.showChartDataTable(HasDataTable=HasDataTable) 
        # set chart data from pandas
        excel = './testfiles/chart.xlsx'
        df = pd.read_excel(excel)
        MP.setChartDataFromPandas(df=df)
        # BarOfBie SplitValue
        xlBarOfPie = MP.xlChartType['xlBarOfPie']
        MP.addChart(xlBarOfPie,Initialized=False)
        # df = pd.read_excel(excel)
        # PT.setChartDataFromPandas(df)
        SplitValue = 3
        MP.adjustChartSplitValue(SplitValue)        
        ## fill chart
        Index = 1
        Point = 1
        MP.fillChart(Index=Index)
        MP.fillChartPoint(Index=Index,Point=Point,r=128)
        ## line chart
        Index = 1
        Weight = 4
        r,g,b = 0,128,0
        MP.lineChart(Index=Index,Weight=Weight,r=0,g=128,b=0) 
        MP.lineChartPoint(Index=Index,Point=Point,Weight=Weight,r=0,g=128,b=256) 

