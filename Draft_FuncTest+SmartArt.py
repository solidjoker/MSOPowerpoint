
# coding: utf-8

# In[1]:


get_ipython().system('jupyter qtconsole --generate-config  ')


# In[30]:


def func(**kwargs):
    print(kwargs)


# In[31]:


func(a=1,b=1)


# In[32]:


def func(a=None,b=None):
    print(a,b)


# In[33]:


dic = {'a':1}
func(**dic)


# In[9]:


def func(a=None,b=None):
    print(a)


# In[10]:


func(a=1,b=1)


# In[ ]:


Layout = 'urn:microsoft.com/office/officeart/2005/8/layout/default'
Layout = '基本列表'
Left, Top, Width, Height = 0,0,100,100
PT.pptSlide.Shapes.AddSmartArt(Layout, Left, Top, Width, Height)

for i in range(1,135): 
    print(PT.pptApp.SmartArtLayouts.Item(i).Name,PT.pptApp.SmartArtLayouts.Item(i).Id)
    # PT.pptApp.SmartArtLayouts.Count

