
# coding: utf-8

# # MSOPowerpointFunc
# 
# update: 2018-12-10
# 
# including:
# <li>sheetsToDfs</li>
# 
# <li>exportPPTtoExcel</li>
# <li>exportDfsToExcel</li>
# <li>getDfsFromTable</li>
# <li>getFillLineFromChart</li>
# <li>getExcelTextFile</li>
# 
# <li>getRGB</li>
# <li>getRGBBack</li>
# 
# <li>funcLookupKey</li>
# <li>funcCmInch</li>
# <li>funcPixelInch</li>
# <li>funcPictureDpi</li>
# <li>funcGetShapeFromHyperlink</li>
# 
# <li>getdir</li>

# In[1]:


import os,tempfile,traceback
import xlrd
import xlwings as xw
from PIL import Image
import pandas as pd
import numpy as np


# In[2]:


def sheetsToDfs(filename):
    # 返回字典{sheet_name:sheet_data}
    assert os.path.exists(filename),'% not found!'%filename
    dfs = {}
    workbook = xlrd.open_workbook(filename)
    sheet_names = workbook.sheet_names()
    del workbook
    for sheet_name in sheet_names:
        dfs[sheet_name] = pd.read_excel(filename,sheet_name=sheet_name).replace(np.nan,'')
    return dfs

def exportPPTtoExcel(MP):
    assert MP.shapesInfo, 'please run PowerpointBase.genGetShapesInfo first!'
    try:
        # 文件名称
        dirname =  MP.presInfo['FileName']
        dirname = dirname[:dirname.rfind('.')]
        if not os.path.exists(dirname):
            os.mkdir(dirname)
        _filename = os.path.basename(MP.presInfo['FileName'])
        _filename = _filename[:_filename.rfind('.')]
        filename = os.path.join(dirname,'%s.xlsx'%_filename)

        # 初始化
        shapesInfo = {k:v for k,v in MP.shapesInfo.items() if v}
        sheetnames = ['Summary'] + list(shapesInfo.keys())
        
        # 初始化
        SummaryColumns = list(MP.presInfo)
        FuncColumns =[
        'SlideIndex','ShapeName','ShapeType','AutoShapeType','ShapeZOrderPosition','Visible','Left','Top','Width','Height',
        'HyperLink','Address','SubAddress',
        'Fill','FillRGB',
        'Line','LineRGB','Weight','DashStyle']
        dfColumns = {
            'Summary':SummaryColumns,
            'Chart':FuncColumns+['ChartType','ChartTitle','Legend','DataTable','SeriesCollectionCount','PlotBy',
                                 'PrimaryCategoryAxis','PrimaryValueAxis','SecondaryCategoryAxis','SecondaryValueAxis',
                                 'PrimaryCategoryGridLinesMajor','PrimaryCategoryGridLinesMinorMajor',
                                 'PrimaryValueGridLinesMajor','PrimaryValueGridLinesMinorMajor',],
            'Table':FuncColumns+['Rows','Columns','StyleId'],
            'SmartArt':FuncColumns, 
            'Text':FuncColumns+['AutoSize','HorizontalAnchor','VerticalAnchor','MarginLeft','MarginTop','MarginRight','MarginBottom',
                                'Text','Name','Size','ColorRGB','Bold','Italic','Underline','BulletType','BulletCharacter',],
            'Shape':FuncColumns+['Source'],}
        dfs = []
        # 生成Summary
        dfname = 'Summary'
        presInfo = {k:v for k,v in MP.presInfo.items() if k in dfColumns[dfname]}
        df = pd.DataFrame(presInfo,index=[0])
        dfs.append(df)
        # 生成 Table,Chart,ext,SmartArt,Shape
        for dfname in sheetnames[1:]:
            df = pd.DataFrame(shapesInfo[dfname]).transpose()[dfColumns[dfname]]
            dfs.append(df)
        # 生成 Table,Chart 明细
        dfsTC,sheetsnameTC = getDfsFromTable(shapesInfo,dfColumns)
        dfs.extend(dfsTC)
        sheetnames.extend(sheetsnameTC)
        # 生成Chart Label,Fill,Line
        dfsFL,sheetsnameFL = getLabelFillLineFromChart(shapesInfo,dfColumns)
        dfs.extend(dfsFL)
        sheetnames.extend(sheetsnameFL)
        # 写入Excel
        exportDfsToExcel(dfs,sheetnames,filename,index=True)
        # 打开Excel
        os.startfile(filename)
        return filename
    except PermissionError as e:
        print('%s\nPlease close the excel:%s'%(e,filename))
        return False 

def exportDfsToExcel(dfs,sheetnames,filename,index=True):
    # 写入Excel
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    for df,sheetname in zip(dfs,sheetnames):
        df.to_excel(excel_writer=writer,sheet_name=sheetname,index=index)
    writer.save()
    writer.close()
    
def getDfsFromTable(shapesInfo,dfColumns):
    # 读取Table,Chart数据
    dfs = []
    sheetsname = []
    for _ in list(set(['Table','Chart']).intersection(set(shapesInfo.keys()))):
        df = pd.DataFrame(shapesInfo[_]).transpose()[dfColumns[_]+['%sData'%_]]
        for i in df.index:
            dfs.append(df.loc[i,'%sData'%_])
            sheetsname.append('%s_%s'%(_,i))
    return dfs,sheetsname

def getLabelFillLineFromChart(shapesInfo,dfColumns):
    # 读取Chart Label,Fill,Line
    dfs = []
    sheetsname = []
    for _ in list(set(['Chart']).intersection(set(shapesInfo.keys()))):
        for fl in ['LabelSetting','SeriesFill','SeriesLine','PointsFill','PointsLine']:
            df = pd.DataFrame(shapesInfo[_]).transpose()[dfColumns[_]+['%s%s'%(_,fl)]]
            for i in df.index:
                dfs.append(df.loc[i,'%s%s'%(_,fl)])
                sheetsname.append('%s%s_%s'%(_,fl,i))
    return dfs,sheetsname

def getExcelTextFile(filename):
    '''
    保存临时文件xlsx，所有值为text
    '''
    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(filename)
    wb = xw.Book(filename) 
    for sht in wb.sheets:
        for cell in sht.range(sht.api.Usedrange.Address):
            cell.value = cell.api.Text
    tmp = tempfile.mktemp('.xlsx')  
    wb.save(tmp)
    wb.close()
    app.quit()
    return tmp

def getRGB(**kwargs):
    '''
    getRGB
    '''
    colordic = {
        'white':{'r':255,'g':255,'b':255},
        'black':{'r':0,'g':0,'b':0},
        'red':{'r':255,'g':0,'b':0},
        'green':{'r':0,'g':255,'b':0},
        'blue':{'r':0,'g':0,'b':255},
        'cyan':{'r':0,'g':255,'b':255},
        'purple':{'r':255,'g':0,'b':255},
        'yellow':{'r':255,'g':255,'b':0},
        'lightblue':{'r':234,'g':239,'b':247},
        'middleblue':{'r':91,'g':155,'b':213}
    }
    alert = 'not valid colorname,should be in\n%s'%list(colordic.keys())
    if not kwargs:
        return False
    flag = False
    for v in kwargs.values():
        if not v == None:
            flag = True
            break
    if not flag:
        return False
    
    try:   
        _kwargs = {k[:5].lower():v for k,v in kwargs.items() if len(k)>=5}
        c = _kwargs.get('color',None)
        if c:
            _color = colordic.get(c.strip().lower(),None)
            if _color:
                return _color['r'] + _color['g']*256+_color['b']*256**2
            else:
                print(alert)
                return False
        else:
            _color = {'r':0,'g':0,'b':0}
            # 对于r,b,g
            _color.update({k.lower():(lambda v: 255 if v > 255 else 0 if v< 0 else v)(v) for k,v in kwargs.items() if not v is None}) 
            # 对于red,green,blue
            _rgb = ['red','green','blue'] 
            _kwargs = {k.lower():(lambda v: 255 if v > 255 else 0 if v< 0 else v)(v) for k,v in kwargs.items() if k.lower() in _rgb} 
            _color.update({k[0].lower():(lambda v: 255 if v > 255 else 0 if v< 0 else v)(v) for k,v in _kwargs.items() if k[0] in _color})
            return _color['r'] + _color['g']*256+_color['b']*256**2
    except:
        traceback.print_exc()
        return False
    
def getRGBback(RGB):
    '''
    根据RGB获取r,g,b
    '''
    try:
        r = RGB%(256)
        g = RGB%(256**2)//256
        b = RGB//(256**2)
        return r,g,b
    except:
        return None,None,None

def funcLookupKey(LookupType,LookupDict):
    '''
    lookup key with value
    '''
    for k,v in LookupDict.items():
        if v == LookupType:
            return k
    return False    

def funcCmInch(data,reverse=None):
    '''
    cm to/from Inch
    '''
    if reverse:
        return data*2.54 
    return data/2.54    

def funcPixelInch(data,reverse=None):
    '''
    pixel to/from Inch
    '''
    if reverse:
        return int(data*72)
    return data/72 

def funcPictureDpi(picture,dpi):
    '''
    adjust Picture DPI
    picture: filename
    '''
    if os.path.exists(picture):
        suffix = picture[picture.rfind('.'):]
        im = Image.open(picture)
        tempf = tempfile.mkstemp(suffix=suffix)[1]
        im.save(tempf,dpi=(dpi,dpi))
        im.close()
        return tempf
    return False

def funcGetShapeFromHyperlink(Hyperlink,maxtries=10):
    Shape = Hyperlink.Parent
    while maxtries:
        maxtries -= 1
        if hasattr(Shape,'ZOrderPosition'):
            return Shape.ZOrderPosition
        else:
            return funcGetShapeFromHyperlink(Shape,maxtries)
    return None

def getdir(pptSel):
    dir(pptSel)
    dir(pptSel._dispobj_)


# In[ ]:


if __name__ == '__main__':
    filename = 'test.xlsx'
    tmp = getExcelTextFile(filename)
    wb = xw.Book(tmp)


# In[4]:


getRGB(r=255,g=255,b=255)

