
# coding: utf-8

# In[27]:


import os, time, datetime, pprint, traceback, tempfile
from collections import defaultdict
from PIL import Image
import pandas as pd
import numpy as np

import win32com
from win32com.client import Dispatch,Constants

import MSOPowerpointConfig
import MSOPowerpointFunc


# In[56]:


appExcel = 'Excel.Application'
ExcelApp = win32com.client.Dispatch(appExcel)
ExcelApp.Visible = 1
ExcelApp.DisplayAlerts = 1 # set DisplayAlerts to ppAlertsNone
wkbs = ExcelApp.Workbooks
# wkb = wkbs.Add()
filename = 'C:/SmithYe/PythonProject3/OfficeApi/MSOPowerpoint/testfiles/template_Table_20190321120214/powerpoint_data.xlsx'
wkb = wkbs.Open(filename)
shts = wkb.WorkSheets


# In[57]:


shtsum = shts['Summary']


# In[65]:


shtsum.Cells(2,15).hyperlinks(1).SubAddress


# In[54]:


rng.Address

