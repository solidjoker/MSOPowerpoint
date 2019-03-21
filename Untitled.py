
# coding: utf-8

# In[66]:


import os, time, datetime, pprint, traceback, tempfile
from collections import defaultdict
from PIL import Image
import pandas as pd
import numpy as np

import win32com
from win32com.client import Dispatch,Constants

import MSOPowerpointConfig
import MSOPowerpointFunc


# In[89]:


appExcel = 'Excel.Application'
ExcelApp = win32com.client.Dispatch(appExcel)
ExcelApp.Visible = 1
ExcelApp.DisplayAlerts = 1 # set DisplayAlerts to ppAlertsNone
wkbs = ExcelApp.Workbooks
# wkb = wkbs.Add()
filename = 'C:/SmithYe/PythonProject3/OfficeApi/MSOPowerpoint/testfiles/template_Table_20190321120214/powerpoint_data.xlsx'
wkb = wkbs.Open(filename)
shts = wkb.WorkSheets


# In[90]:


def linkExcelTableInformation(shts):
    '''
    get table information
    '''
    shtnames = [i.name for i in shts]
    if 'Table_Summary' in shtnames:
        shttablesum = shts['Table_Summary']
        cells1 = shttablesum.Cells(1,1)
        cells2 = shttablesum.Cells(shttablesum.UsedRange.Rows.Count,1)
        rng = shttablesum.Range(cells1,cells2)
        tabledic = defaultdict(list)
        for i in rng:
            tabledic[i.text].append(i.row)
        del tabledic['']
        return {k:(min(v),max(v)) for k,v in tabledic.items()}  


# In[91]:


tablesummary = linkExcelTableInformation(shts)


# In[92]:


tablesummary


# In[93]:


shtTableData = shts['Table_Data']   
cells1 = shtTableData.Cells(1,1)


# In[96]:


cells1.CurrentRegion.Copy()


# In[79]:


from collections import defaultdict
tabledic = defaultdict(list)

for i in rng:
    tabledic[i.text].append(i.row)
    print(i.text,i.row)
    
del tabledic['']
tablesummarydic = {k:(min(v),max(v)) for k,v in tabledic.items()}


# In[86]:


tablesummarydic

