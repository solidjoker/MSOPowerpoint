
# coding: utf-8

# # MSOPowerpointConfig
# update: 2018-12-05
# 
# including:
# <li>msoShapeType</li>
# <li>msoAutoShapeType</li>
# <li>msoTableStyle</li>
# <li>xlChartType</li>

# In[ ]:


customInfo ={
    'tries':5,
    'seconds':2,
    'encoding':'gbk',
}


# In[1]:


# msoShapeType https://docs.microsoft.com/zh-CN/office/vba/api/Office.MsoShapeType
msoShapeType = {
    'msoAutoShape':1,
    'msoCallout':2,
    'msoCanvas':20,
    'msoChart':3,
    'msoComment':4,
    'msoContentApp':27,
    'msoDiagram':21,
    'msoEmbeddedOLEObject':7,
    'msoFormControl':8,
    'msoFreeform':5,
    'msoGraphic':28,
    'msoGroup':6,
    'msoIgxGraphic':24,
    'msoInk':22,
    'msoInkComment':23,
    'msoLine':9,
    'msoLinkedGraphic':29,
    'msoLinkedOLEObject':10,
    'msoLinkedPicture':11,
    'msoMedia':16,
    'msoOLEControlObject':12,
    'msoPicture':13,
    'msoPlaceholder':14,
    'msoScriptAnchor':18,
    'msoShapeTypeMixed':-2,
    'msoTable':19,
    'msoTextBox':17,
    'msoTextEffect':15,
    'msoWebVideo':26,
    '自选图形':1,
    '标注':2,
    '画布':20,
    '图':3,
    '批注':4,
    '内容的 Office 加载项':27,
    '图表':21,
    '嵌入的 OLE 对象':7,
    '窗体控件':8,
    '任意多边形':5,
    '图形':28,
    '组合':6,
    'SmartArt 图形':24,
    '墨迹':22,
    '墨迹批注':23,
    '线条':9,
    '链接的图形':29,
    '链接 OLE 对象':10,
    '链接图片':11,
    '媒体':16,
    'OLE 控件对象':12,
    '图片':13,
    '占位符':14,
    '脚本定位标记':18,
    '混和形状类型':-2,
    '表':19,
    '文本框':17,
    '文本效果':15,
    'Web 视频':26,
}


# In[2]:


# msoAutoShapeType https://docs.microsoft.com/zh-CN/office/vba/api/Office.MsoAutoShapeType
msoAutoShapeType = {
    'msoShape10pointStar': 149,
    'msoShape12pointStar': 150,
    'msoShape16pointStar': 94,
    'msoShape24pointStar': 95,
    'msoShape32pointStar': 96,
    'msoShape4pointStar': 91,
    'msoShape5pointStar': 92,
    'msoShape6pointStar': 147,
    'msoShape7pointStar': 148,
    'msoShape8pointStar': 93,
    'msoShapeActionButtonBackorPrevious': 129,
    'msoShapeActionButtonBeginning': 131,
    'msoShapeActionButtonCustom': 125,
    'msoShapeActionButtonDocument': 134,
    'msoShapeActionButtonEnd': 132,
    'msoShapeActionButtonForwardorNext': 130,
    'msoShapeActionButtonHelp': 127,
    'msoShapeActionButtonHome': 126,
    'msoShapeActionButtonInformation': 128,
    'msoShapeActionButtonMovie': 136,
    'msoShapeActionButtonReturn': 133,
    'msoShapeActionButtonSound': 135,
    'msoShapeArc': 25,
    'msoShapeBalloon': 137,
    'msoShapeBentArrow': 41,
    'msoShapeBentUpArrow': 44,
    'msoShapeBevel': 15,
    'msoShapeBlockArc': 20,
    'msoShapeCan': 13,
    'msoShapeChartPlus': 182,
    'msoShapeChartStar': 181,
    'msoShapeChartX': 180,
    'msoShapeChevron': 52,
    'msoShapeChord': 161,
    'msoShapeCircularArrow': 60,
    'msoShapeCloud': 179,
    'msoShapeCloudCallout': 108,
    'msoShapeCorner': 162,
    'msoShapeCornerTabs': 169,
    'msoShapeCross': 11,
    'msoShapeCube': 14,
    'msoShapeCurvedDownArrow': 48,
    'msoShapeCurvedDownRibbon': 100,
    'msoShapeCurvedLeftArrow': 46,
    'msoShapeCurvedRightArrow': 45,
    'msoShapeCurvedUpArrow': 47,
    'msoShapeCurvedUpRibbon': 99,
    'msoShapeDecagon': 144,
    'msoShapeDiagonalStripe': 141,
    'msoShapeDiamond': 4,
    'msoShapeDodecagon': 146,
    'msoShapeDonut': 18,
    'msoShapeDoubleBrace': 27,
    'msoShapeDoubleBracket': 26,
    'msoShapeDoubleWave': 104,
    'msoShapeDownArrow': 36,
    'msoShapeDownArrowCallout': 56,
    'msoShapeDownRibbon': 98,
    'msoShapeExplosion1': 89,
    'msoShapeExplosion2': 90,
    'msoShapeFlowchartAlternateProcess': 62,
    'msoShapeFlowchartCard': 75,
    'msoShapeFlowchartCollate': 79,
    'msoShapeFlowchartConnector': 73,
    'msoShapeFlowchartData': 64,
    'msoShapeFlowchartDecision': 63,
    'msoShapeFlowchartDelay': 84,
    'msoShapeFlowchartDirectAccessStorage': 87,
    'msoShapeFlowchartDisplay': 88,
    'msoShapeFlowchartDocument': 67,
    'msoShapeFlowchartExtract': 81,
    'msoShapeFlowchartInternalStorage': 66,
    'msoShapeFlowchartMagneticDisk': 86,
    'msoShapeFlowchartManualInput': 71,
    'msoShapeFlowchartManualOperation': 72,
    'msoShapeFlowchartMerge': 82,
    'msoShapeFlowchartMultidocument': 68,
    'msoShapeFlowchartOfflineStorage': 139,
    'msoShapeFlowchartOffpageConnector': 74,
    'msoShapeFlowchartOr': 78,
    'msoShapeFlowchartPredefinedProcess': 65,
    'msoShapeFlowchartPreparation': 70,
    'msoShapeFlowchartProcess': 61,
    'msoShapeFlowchartPunchedTape': 76,
    'msoShapeFlowchartSequentialAccessStorage': 85,
    'msoShapeFlowchartSort': 80,
    'msoShapeFlowchartStoredData': 83,
    'msoShapeFlowchartSummingJunction': 77,
    'msoShapeFlowchartTerminator': 69,
    'msoShapeFoldedCorner': 16,
    'msoShapeFrame': 158,
    'msoShapeFunnel': 174,
    'msoShapeGear6': 172,
    'msoShapeGear9': 173,
    'msoShapeHalfFrame': 159,
    'msoShapeHeart': 21,
    'msoShapeHeptagon': 145,
    'msoShapeHexagon': 10,
    'msoShapeHorizontalScroll': 102,
    'msoShapeIsoscelesTriangle': 7,
    'msoShapeLeftArrow': 34,
    'msoShapeLeftArrowCallout': 54,
    'msoShapeLeftBrace': 31,
    'msoShapeLeftBracket': 29,
    'msoShapeLeftCircularArrow': 176,
    'msoShapeLeftRightArrow': 37,
    'msoShapeLeftRightArrowCallout': 57,
    'msoShapeLeftRightCircularArrow': 177,
    'msoShapeLeftRightRibbon': 140,
    'msoShapeLeftRightUpArrow': 40,
    'msoShapeLeftUpArrow': 43,
    'msoShapeLightningBolt': 22,
    'msoShapeLineCallout1': 109,
    'msoShapeLineCallout1AccentBar': 113,
    'msoShapeLineCallout1BorderandAccentBar': 121,
    'msoShapeLineCallout1NoBorder': 117,
    'msoShapeLineCallout2': 110,
    'msoShapeLineCallout2AccentBar': 114,
    'msoShapeLineCallout2BorderandAccentBar': 122,
    'msoShapeLineCallout2NoBorder': 118,
    'msoShapeLineCallout3': 111,
    'msoShapeLineCallout3AccentBar': 115,
    'msoShapeLineCallout3BorderandAccentBar': 123,
    'msoShapeLineCallout3NoBorder': 119,
    'msoShapeLineCallout4': 112,
    'msoShapeLineCallout4AccentBar': 116,
    'msoShapeLineCallout4BorderandAccentBar': 124,
    'msoShapeLineCallout4NoBorder': 120,
    'msoShapeLineInverse': 183,
    'msoShapeMathDivide': 166,
    'msoShapeMathEqual': 167,
    'msoShapeMathMinus': 164,
    'msoShapeMathMultiply': 165,
    'msoShapeMathNotEqual': 168,
    'msoShapeMathPlus': 163,
    'msoShapeMixed': -2,
    'msoShapeMoon': 24,
    'msoShapeNonIsoscelesTrapezoid': 143,
    'msoShapeNoSymbol': 19,
    'msoShapeNotchedRightArrow': 50,
    'msoShapeNotPrimitive': 138,
    'msoShapeOctagon': 6,
    'msoShapeOval': 9,
    'msoShapeOvalCallout': 107,
    'msoShapeParallelogram': 2,
    'msoShapePentagon': 51,
    'msoShapePie': 142,
    'msoShapePieWedge': 175,
    'msoShapePlaque': 28,
    'msoShapePlaqueTabs': 171,
    'msoShapeQuadArrow': 39,
    'msoShapeQuadArrowCallout': 59,
    'msoShapeRectangle': 1,
    'msoShapeRectangularCallout': 105,
    'msoShapeRegularPentagon': 12,
    'msoShapeRightArrow': 33,
    'msoShapeRightArrowCallout': 53,
    'msoShapeRightBrace': 32,
    'msoShapeRightBracket': 30,
    'msoShapeRightTriangle': 8,
    'msoShapeRound1Rectangle': 151,
    'msoShapeRound2DiagRectangle': 157,
    'msoShapeRound2SameRectangle': 152,
    'msoShapeRoundedRectangle': 5,
    'msoShapeRoundedRectangularCallout': 106,
    'msoShapeSmileyFace': 17,
    'msoShapeSnip1Rectangle': 155,
    'msoShapeSnip2DiagRectangle': 157,
    'msoShapeSnip2SameRectangle': 156,
    'msoShapeSnipRoundRectangle': 154,
    'msoShapeSquareTabs': 170,
    'msoShapeStripedRightArrow': 49,
    'msoShapeSun': 23,
    'msoShapeSwooshArrow': 178,
    'msoShapeTear': 160,
    'msoShapeTrapezoid': 3,
    'msoShapeUpArrow': 35,
    'msoShapeUpArrowCallout': 55,
    'msoShapeUpDownArrow': 38,
    'msoShapeUpDownArrowCallout': 58,
    'msoShapeUpRibbon': 97,
    'msoShapeUTurnArrow': 42,
    'msoShapeVerticalScroll': 101,
    'msoShapeWave': 103,
    '10 角星': 149,
    '12 角星': 150,
    '十六角星': 94,
    '二十四角星': 95,
    '三十二角星': 96,
    '四角星': 91,
    '五角星': 92,
    '6 角星': 147,
    '7 角星': 148,
    '八角星': 93,
    '返回或上一步按钮 支持鼠标单击和鼠标移过操作': 129,
    '开始按钮 支持鼠标单击和鼠标移过操作': 131,
    '与没有默认图片或文本的按钮 支持鼠标单击和鼠标移过操作': 125,
    '文档的按钮 支持鼠标单击和鼠标移过操作': 134,
    '结束按钮 支持鼠标单击和鼠标移过操作': 132,
    '转接或下一步按钮 支持鼠标单击和鼠标移过操作': 130,
    '帮助按钮 支持鼠标单击和鼠标移过操作': 127,
    '主页按钮 支持鼠标单击和鼠标移过操作': 126,
    '信息按钮 支持鼠标单击和鼠标移过操作': 128,
    '影片按钮 支持鼠标单击和鼠标移过操作': 136,
    '返回按钮 支持鼠标单击和鼠标移过操作': 133,
    '声音按钮 支持鼠标单击和鼠标移过操作': 135,
    '弧形': 25,
    '气球': 137,
    '带 90 度圆角的箭头': 41,
    'Sharp 90 度圆角的箭头 默认情况下安装点': 44,
    '凹凸效果': 15,
    '空心弧': 20,
    '圆柱形': 13,
    '方形划分为四个季度的垂直和水平': 182,
    '方形划分六个部件沿垂直和对角线': 181,
    '方形分为四个部分沿对角线': 180,
    'V 形': 52,
    '通过在这个圆; 的内部连接两个点上外围一条线的圆弦的圆': 161,
    '带 180 度圆角的箭头': 60,
    '云形状': 179,
    '云形标注': 108,
    '具有矩形形状孔效果的矩形': 162,
    '沿矩形路径; 对齐的四个右三角形四个段角': 169,
    '十字形': 11,
    '立方': 14,
    '上弧形箭头': 48,
    '下凸弯带形横幅': 100,
    '右弧形箭头': 46,
    '左弧形箭头': 45,
    '下弧形箭头': 47,
    '上凸弯带形': 99,
    'Decagon': 144,
    '具有两个三角形的形状中删除; 效果的矩形斜向条带': 141,
    '菱形': 4,
    'Dodecagon': 146,
    '环形': 18,
    '双大括号': 27,
    '双括号': 26,
    '双波形': 104,
    '下箭头': 36,
    '带下箭头的标注': 56,
    '中心区域位于弯带末端下方的弯带形': 98,
    '爆炸形': 90,
    '其他过程流程图符号': 62,
    '资料卡流程图符号': 75,
    '对照流程图符号': 79,
    '联系流程图符号': 73,
    '数据流程图符号': 64,
    '决策流程图符号': 63,
    '延期流程图符号': 84,
    '磁鼓流程图符号': 87,
    '显示流程图符号': 88,
    '文档流程图符号': 67,
    '摘录流程图符号': 81,
    '内部贮存流程图符号': 66,
    '磁盘流程图符号': 86,
    '手动输入流程图符号': 71,
    '手动操作流程图符号': 72,
    '合并流程图符号': 82,
    '多文档流程图符号': 68,
    '脱机存储区流程图符号': 139,
    '离页连接符流程图符号': 74,
    '“或者”流程图符号': 78,
    '预定义过程流程图符号': 65,
    '准备流程图符号': 70,
    '过程流程图符号': 61,
    '资料带流程图符号': 76,
    '磁带流程图符号': 85,
    '排序流程图符号': 80,
    '库存数据流程图符号': 83,
    '汇总连接流程图符号': 77,
    '终止流程图符号': 69,
    '折角形': 16,
    '矩形的图片框': 158,
    '漏斗': 174,
    '与六个齿齿轮': 172,
    '与九个齿齿轮': 173,
    '矩形的图片框的一半': 159,
    '心形': 21,
    'Heptagon': 145,
    '六边形': 10,
    '横卷形': 102,
    '等腰三角形': 7,
    '左箭头': 34,
    '带左箭头的标注': 54,
    '左大括号': 31,
    '左括号': 29,
    '向逆时针方向的环形箭头': 176,
    '左右双向箭头': 37,
    '带左右双向箭头的标注': 57,
    '指向顺时针方向和逆时针; 的环形箭头点两端都与一个曲线的箭头': 177,
    '带有两端箭头的功能区': 140,
    '左右上三向箭头': 40,
    '左上双向箭头': 43,
    '闪电形': 22,
    '带边框和水平标注线的标注': 109,
    '带水平强调线的标注': 113,
    '带边框和水平强调线的标注': 121,
    '带水平线的标注': 117,
    '带对角直线的标注': 110,
    '带对角标注线和强调线的标注': 114,
    '带边框、对角直线和强调线的标注': 122,
    '不带边框和对角标注线的标注': 118,
    '带倾斜线的标注': 111,
    '带倾斜标注线和强调线的标注': 115,
    '带边框、倾斜标注线和强调线的标注': 123,
    '不带边框和倾斜标注线的标注': 119,
    '带 U 型标注线段的标注': 112,
    '带强调线和 U 型标注线段的标注': 116,
    '带边框、强调线和 U 型标注线段的标注': 124,
    '不带边框和 U 型标注线段的标注': 120,
    '行的反函数': 183,
    '除号 /': 166,
    '等价符号 =': 167,
    '减法符号-': 164,
    '乘法符号 x': 165,
    '非等价符号 ≠': 168,
    '添加符号 +': 163,
    '只返回值，表示其他状态的组合': -2,
    '新月形': 24,
    '与非对称非并行侧面梯形': 143,
    '禁止符': 19,
    '燕尾形右箭头': 50,
    '不支持': 138,
    '八边形': 6,
    '椭圆形': 9,
    '椭圆形标注': 107,
    '平行四边形': 2,
    '五边形': 12,
    "部分缺少圆 （' 饼图）": 142,
    '循环形状的季度': 175,
    '缺角矩形': 28,
    '四个季度-圆定义一个矩形形状': 171,
    '四向箭头': 39,
    '带四向箭头的标注': 59,
    '矩形': 1,
    '矩形标注': 105,
    '右箭头': 33,
    '带右箭头的标注': 53,
    '右大括号': 32,
    '右括号': 30,
    '直角三角形': 8,
    '与一个圆角矩形': 151,
    '具有两个沿对角线值相对的圆角矩形': 157,
    '带有共享一侧的两个舍入角的矩形': 152,
    '圆角矩形': 5,
    '圆角矩形标注': 106,
    '笑脸': 17,
    '与段的一个角的矩形': 155,
    '带有两段的角，沿对角线值相对的矩形': 157,
    '带有共享一侧的段的两个角的矩形': 156,
    '一个段角与一个圆的角矩形': 154,
    '定义一个矩形形状的四个小平方': 170,
    '尾部带条纹的右箭头': 49,
    '太阳': 23,
    '曲线的箭头': 178,
    '水快捷批处理': 160,
    '梯形': 3,
    '上箭头': 35,
    '带上箭头的标注': 55,
    '上下双向箭头': 38,
    '带上下双向箭头的标注': 58,
    '中心区域位于弯带末端上方的弯带形横幅': 97,
    'U 型箭头': 42,
    '竖卷形': 101,
    '波形': 103,
}     


# In[3]:


# msoTableStyle https://docs.microsoft.com/zh-cn/previous-versions/office/hh273476(v=office.14)
msoTableStyle={ 
    'No Style, No Grid':'{2D5ABB26-0587-4C30-8999-92F81FD0307C}',
    'Themed Style 1 - Accent 1':'{3C2FFA5D-87B4-456A-9821-1D502468CF0F}',
    'Themed Style 1 - Accent 2':'{284E427A-3D55-4303-BF80-6455036E1DE7}',
    'Themed Style 1 - Accent 3':'{69C7853C-536D-4A76-A0AE-DD22124D55A5}',
    'Themed Style 1 - Accent 4':'{775DCB02-9BB8-47FD-8907-85C794F793BA}',
    'Themed Style 1 - Accent 5':'{35758FB7-9AC5-4552-8A53-C91805E547FA}',
    'Themed Style 1 - Accent 6':'{08FB837D-C827-4EFA-A057-4D05807E0F7C}',
    'No Style, Table Grid':'{5940675A-B579-460E-94D1-54222C63F5DA}',
    'Themed Style 2 - Accent 1':'{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}',
    'Themed Style 2 - Accent 2':'{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}',
    'Themed Style 2 - Accent 3':'{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}',
    'Themed Style 2 - Accent 4':'{E269D01E-BC32-4049-B463-5C60D7B0CCD2}',
    'Themed Style 2 - Accent 5':'{327F97BB-C833-4FB7-BDE5-3F7075034690}',
    'Themed Style 2 - Accent 6':'{638B1855-1B75-4FBE-930C-398BA8C253C6}',
    'Light Style 1':'{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}',
    'Light Style 1 - Accent 1':'{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}',
    'Light Style 1 - Accent 2':'{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}',
    'Light Style 1 - Accent 3':'{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}',
    'Light Style 1 - Accent 4':'{D27102A9-8310-4765-A935-A1911B00CA55}',
    'Light Style 1 - Accent 5':'{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}',
    'Light Style 1 - Accent 6':'{68D230F3-CF80-4859-8CE7-A43EE81993B5}',
    'Light Style 2':'{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}',
    'Light Style 2 - Accent 1':'{69012ECD-51FC-41F1-AA8D-1B2483CD663E}',
    'Light Style 2 - Accent 2':'{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}',
    'Light Style 2 - Accent 3':'{F2DE63D5-997A-4646-A377-4702673A728D}',
    'Light Style 2 - Accent 4':'{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}',
    'Light Style 2 - Accent 5':'{5A111915-BE36-4E01-A7E5-04B1672EAD32}',
    'Light Style 2 - Accent 6':'{912C8C85-51F0-491E-9774-3900AFEF0FD7}',
    'Light Style 3':'{616DA210-FB5B-4158-B5E0-FEB733F419BA}',
    'Light Style 3 - Accent 1':'{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}',
    'Light Style 3 - Accent 2':'{5DA37D80-6434-44D0-A028-1B22A696006F}',
    'Light Style 3 - Accent 3':'{8799B23B-EC83-4686-B30A-512413B5E67A}',
    'Light Style 3 - Accent 4':'{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}',
    'Light Style 3 - Accent 5':'{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}',
    'Light Style 3 - Accent 6':'{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}',
    'Medium Style 1':'{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}',
    'Medium Style 1 - Accent 1':'{B301B821-A1FF-4177-AEE7-76D212191A09}',
    'Medium Style 1 - Accent 2':'{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}',
    'Medium Style 1 - Accent 3':'{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}',
    'Medium Style 1 - Accent 4':'{1E171933-4619-4E11-9A3F-F7608DF75F80}',
    'Medium Style 1 - Accent 5':'{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}',
    'Medium Style 1 - Accent 6':'{10A1B5D5-9B99-4C35-A422-299274C87663}',
    'Medium Style 2':'{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}',
    'Medium Style 2 - Accent 1':'{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}',
    'Medium Style 2 - Accent 2':'{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}',
    'Medium Style 2 - Accent 3':'{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}',
    'Medium Style 2 - Accent 4':'{00A15C55-8517-42AA-B614-E9B94910E393}',
    'Medium Style 2 - Accent 5':'{7DF18680-E054-41AD-8BC1-D1AEF772440D}',
    'Medium Style 2 - Accent 6':'{93296810-A885-4BE3-A3E7-6D5BEEA58F35}',
    'Medium Style 3':'{8EC20E35-A176-4012-BC5E-935CFFF8708E}',
    'Medium Style 3 - Accent 1':'{6E25E649-3F16-4E02-A733-19D2CDBF48F0}',
    'Medium Style 3 - Accent 2':'{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}',
    'Medium Style 3 - Accent 3':'{EB344D84-9AFB-497E-A393-DC336BA19D2E}',
    'Medium Style 3 - Accent 4':'{EB9631B5-78F2-41C9-869B-9F39066F8104}',
    'Medium Style 3 - Accent 5':'{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}',
    'Medium Style 3 - Accent 6':'{2A488322-F2BA-4B5B-9748-0D474271808F}',
    'Medium Style 4':'{D7AC3CCA-C797-4891-BE02-D94E43425B78}',
    'Medium Style 4 - Accent 1':'{69CF1AB2-1976-4502-BF36-3FF5EA218861}',
    'Medium Style 4 - Accent 2':'{8A107856-5554-42FB-B03E-39F5DBC370BA}',
    'Medium Style 4 - Accent 3':'{0505E3EF-67EA-436B-97B2-0124C06EBD24}',
    'Medium Style 4 - Accent 4':'{C4B1156A-380E-4F78-BDF5-A606A8083BF9}',
    'Medium Style 4 - Accent 5':'{22838BEF-8BB2-4498-84A7-C5851F593DF1}',
    'Medium Style 4 - Accent 6':'{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}',
    'Dark Style 1':'{E8034E78-7F5D-4C2E-B375-FC64B27BC917}',
    'Dark Style 1 - Accent 1':'{125E5076-3810-47DD-B79F-674D7AD40C01}',
    'Dark Style 1 - Accent 2':'{37CE84F3-28C3-443E-9E96-99CF82512B78}',
    'Dark Style 1 - Accent 3':'{D03447BB-5D67-496B-8E87-E561075AD55C}',
    'Dark Style 1 - Accent 4':'{E929F9F4-4A8F-4326-A1B4-22849713DDAB}',
    'Dark Style 1 - Accent 5':'{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}',
    'Dark Style 1 - Accent 6':'{AF606853-7671-496A-8E4F-DF71F8EC918B}',
    'Dark Style 2':'{5202B0CA-FC54-4496-8BCA-5EF66A818D29}',
    'Dark Style 2 - Accent 1/Accent 2':'{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}',
    'Dark Style 2 - Accent 3/Accent 4':'{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}',
    'Dark Style 2 - Accent 5/Accent 6':'{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}',
    '无样式，无网格':'{2D5ABB26-0587-4C30-8999-92F81FD0307C}', 
    '主题样式 1 - 强调 1':'{3C2FFA5D-87B4-456A-9821-1D502468CF0F}',
    '主题样式 1 - 强调 2':'{284E427A-3D55-4303-BF80-6455036E1DE7}',
    '主题样式 1 - 强调 3':'{69C7853C-536D-4A76-A0AE-DD22124D55A5}',
    '主题样式 1 - 强调 4':'{775DCB02-9BB8-47FD-8907-85C794F793BA}',
    '主题样式 1 - 强调 5':'{35758FB7-9AC5-4552-8A53-C91805E547FA}',
    '主题样式 1 - 强调 6':'{08FB837D-C827-4EFA-A057-4D05807E0F7C}',
    '无样式，网格型':'{5940675A-B579-460E-94D1-54222C63F5DA}',
    '主题样式 2 - 强调 1':'{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}',
    '主题样式 2 - 强调 2':'{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}',
    '主题样式 2 - 强调 3':'{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}',
    '主题样式 2 - 强调 4':'{E269D01E-BC32-4049-B463-5C60D7B0CCD2}',
    '主题样式 2 - 强调 5':'{327F97BB-C833-4FB7-BDE5-3F7075034690}',
    '主题样式 2 - 强调 6':'{638B1855-1B75-4FBE-930C-398BA8C253C6}',
    '浅色样式 1':'{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}',
    '浅色样式 1 - 强调 1':'{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}',
    '浅色样式 1 - 强调 2':'{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}',
    '浅色样式 1 - 强调 3':'{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}',
    '浅色样式 1 - 强调 4':'{D27102A9-8310-4765-A935-A1911B00CA55}',
    '浅色样式 1 - 强调 5':'{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}',
    '浅色样式 1 - 强调 6':'{68D230F3-CF80-4859-8CE7-A43EE81993B5}',
    '浅色样式 2':'{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}',
    '浅色样式 2 - 强调 1':'{69012ECD-51FC-41F1-AA8D-1B2483CD663E}',
    '浅色样式 2 - 强调 2':'{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}',
    '浅色样式 2 - 强调 3':'{F2DE63D5-997A-4646-A377-4702673A728D}',
    '浅色样式 2 - 强调 4':'{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}',
    '浅色样式 2 - 强调 5':'{5A111915-BE36-4E01-A7E5-04B1672EAD32}',
    '浅色样式 2 - 强调 6':'{912C8C85-51F0-491E-9774-3900AFEF0FD7}',
    '浅色样式 3':'{616DA210-FB5B-4158-B5E0-FEB733F419BA}',
    '浅色样式 3 - 强调 1':'{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}',
    '浅色样式 3 - 强调 2':'{5DA37D80-6434-44D0-A028-1B22A696006F}',
    '浅色样式 3 - 强调 3':'{8799B23B-EC83-4686-B30A-512413B5E67A}',
    '浅色样式 3 - 强调 4':'{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}',
    '浅色样式 3 - 强调 5':'{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}',
    '浅色样式 3 - 强调 6':'{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}',
    '中度样式 1':'{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}',
    '中度样式 1 - 强调 1':'{B301B821-A1FF-4177-AEE7-76D212191A09}',
    '中度样式 1 - 强调 2':'{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}',
    '中度样式 1 - 强调 3':'{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}',
    '中度样式 1 - 强调 4':'{1E171933-4619-4E11-9A3F-F7608DF75F80}',
    '中度样式 1 - 强调 5':'{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}',
    '中度样式 1 - 强调 6':'{10A1B5D5-9B99-4C35-A422-299274C87663}',
    '中度样式 2':'{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}',
    '中度样式 2 - 强调 1':'{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}',
    '中度样式 2 - 强调 2':'{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}',
    '中度样式 2 - 强调 3':'{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}',
    '中度样式 2 - 强调 4':'{00A15C55-8517-42AA-B614-E9B94910E393}',
    '中度样式 2 - 强调 5':'{7DF18680-E054-41AD-8BC1-D1AEF772440D}',
    '中度样式 2 - 强调 6':'{93296810-A885-4BE3-A3E7-6D5BEEA58F35}',
    '中度样式 3':'{8EC20E35-A176-4012-BC5E-935CFFF8708E}',
    '中度样式 3 - 强调 1':'{6E25E649-3F16-4E02-A733-19D2CDBF48F0}',
    '中度样式 3 - 强调 2':'{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}',
    '中度样式 3 - 强调 3':'{EB344D84-9AFB-497E-A393-DC336BA19D2E}',
    '中度样式 3 - 强调 4':'{EB9631B5-78F2-41C9-869B-9F39066F8104}',
    '中度样式 3 - 强调 5':'{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}',
    '中度样式 3 - 强调 6':'{2A488322-F2BA-4B5B-9748-0D474271808F}',
    '中度样式 4':'{D7AC3CCA-C797-4891-BE02-D94E43425B78}',
    '中度样式 4 - 强调 1':'{69CF1AB2-1976-4502-BF36-3FF5EA218861}',
    '中度样式 4 - 强调 2':'{8A107856-5554-42FB-B03E-39F5DBC370BA}',
    '中度样式 4 - 强调 3':'{0505E3EF-67EA-436B-97B2-0124C06EBD24}',
    '中度样式 4 - 强调 4':'{C4B1156A-380E-4F78-BDF5-A606A8083BF9}',
    '中度样式 4 - 强调 5':'{22838BEF-8BB2-4498-84A7-C5851F593DF1}',
    '中度样式 4 - 强调 6':'{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}',
    '深色样式 1':'{E8034E78-7F5D-4C2E-B375-FC64B27BC917}',
    '深色样式 1 - 强调 1':'{125E5076-3810-47DD-B79F-674D7AD40C01}',
    '深色样式 1 - 强调 2':'{37CE84F3-28C3-443E-9E96-99CF82512B78}',
    '深色样式 1 - 强调 3':'{D03447BB-5D67-496B-8E87-E561075AD55C}',
    '深色样式 1 - 强调 4':'{E929F9F4-4A8F-4326-A1B4-22849713DDAB}',
    '深色样式 1 - 强调 5':'{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}',
    '深色样式 1 - 强调 6':'{AF606853-7671-496A-8E4F-DF71F8EC918B}',
    '深色样式 2':'{5202B0CA-FC54-4496-8BCA-5EF66A818D29}',
    '深色样式 2 - 强调 1/强调 2':'{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}',
    '深色样式 2 - 强调 3/强调 4':'{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}',
    '深色样式 2 - 强调 5/强调 6':'{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}',
}


# In[4]:


# xlChartType https://docs.microsoft.com/zh-CN/office/vba/api/excel.xlcharttype
xlChartType = {
    'xl3DArea':-4098,
    'xl3DAreaStacked':78,
    'xl3DAreaStacked100':79,
    'xl3DBarClustered':60,
    'xl3DBarStacked':61,
    'xl3DBarStacked100':62,
    'xl3DColumn':-4100,
    'xl3DColumnClustered':54,
    'xl3DColumnStacked':55,
    'xl3DColumnStacked100':56,
    'xl3DLine':-4101,
    'xl3DPie':-4102,
    'xl3DPieExploded':70,
    'xlArea':1,
    'xlAreaStacked':76,
    'xlAreaStacked100':77,
    'xlBarClustered':57,
    'xlBarOfPie':71,
    'xlBarStacked':58,
    'xlBarStacked100':59,
    'xlBubble':15,
    'xlBubble3DEffect':87,
    'xlColumnClustered':51,
    'xlColumnStacked':52,
    'xlColumnStacked100':53,
    'xlConeBarClustered':102,
    'xlConeBarStacked':103,
    'xlConeBarStacked100':104,
    'xlConeCol':105,
    'xlConeColClustered':99,
    'xlConeColStacked':100,
    'xlConeColStacked100':101,
    'xlCylinderBarClustered':95,
    'xlCylinderBarStacked':96,
    'xlCylinderBarStacked100':97,
    'xlCylinderCol':98,
    'xlCylinderColClustered':92,
    'xlCylinderColStacked':93,
    'xlCylinderColStacked100':94,
    'xlDoughnut':-4120,
    'xlDoughnutExploded':80,
    'xlLine':4,
    'xlLineMarkers':65,
    'xlLineMarkersStacked':66,
    'xlLineMarkersStacked100':67,
    'xlLineStacked':63,
    'xlLineStacked100':64,
    'xlPie':5,
    'xlPieExploded':69,
    'xlPieOfPie':68,
    'xlPyramidBarClustered':109,
    'xlPyramidBarStacked':110,
    'xlPyramidBarStacked100':111,
    'xlPyramidCol':112,
    'xlPyramidColClustered':106,
    'xlPyramidColStacked':107,
    'xlPyramidColStacked100':108,
    'xlRadar':-4151,
    'xlRadarFilled':82,
    'xlRadarMarkers':81,
    'xlStockHLC':88,
    'xlStockOHLC':89,
    'xlStockVHLC':90,
    'xlStockVOHLC':91,
    'xlSurface':83,
    'xlSurfaceTopView':85,
    'xlSurfaceTopViewWireframe':86,
    'xlSurfaceWireframe':84,
    'xlXYScatter':-4169,
    'xlXYScatterLines':74,
    'xlXYScatterLinesNoMarkers':75,
    'xlXYScatterSmooth':72,
    'xlXYScatterSmoothNoMarkers':73,
    '三维面积图':-4098,
    '三维堆积面积图':78,
    '百分比堆积面积图':77,
    '三维簇状条形图':60,
    '三维堆积条形图':61,
    '三维百分比堆积条形图':62,
    '三维柱形图':-4100,
    '三维簇状柱形图':54,
    '三维堆积柱形图':55,
    '三维百分比堆积柱形图':56,
    '三维折线图':-4101,
    '三维饼图':-4102,
    '分离型三维饼图':70,
    '面积图':1,
    '堆积面积图':76,
    '簇状条形图':57,
    '复合条饼图':71,
    '堆积条形图':58,
    '百分比堆积条形图':59,
    '气泡图':15,
    '三维气泡图':87,
    '簇状柱形图':51,
    '堆积柱形图':52,
    '百分比堆积柱形图':53,
    '簇状条形圆锥图':102,
    '堆积条形圆锥图':103,
    '百分比堆积条形圆锥图':104,
    '三维柱形圆锥图':105,
    '簇状柱形圆锥图':92,
    '堆积柱形圆锥图':93,
    '百分比堆积柱形圆锥图':101,
    '簇状条形圆柱图':95,
    '堆积条形圆柱图':96,
    '百分比堆积条形圆柱图':97,
    '三维柱形圆柱图':98,
    '百分比堆积柱形圆柱图':94,
    '圆环图':-4120,
    '分离型圆环图':80,
    '折线图':4,
    '数据点折线图':65,
    '堆积数据点折线图':66,
    '百分比堆积数据点折线图':67,
    '堆积折线图':63,
    '百分比堆积折线图':64,
    '饼图':5,
    '分离型饼图':69,
    '复合饼图':68,
    '簇状条形棱锥图':109,
    '堆积条形棱锥图':110,
    '百分比堆积条形棱锥图':111,
    '三维柱形棱锥图':112,
    '簇状柱形棱锥图':106,
    '堆积柱形棱锥图':107,
    '百分比堆积柱形棱锥图':108,
    '雷达图':-4151,
    '填充雷达图':82,
    '数据点雷达图':81,
    '盘高-盘低-收盘图':88,
    '开盘-盘高-盘低-收盘图':89,
    '成交量-盘高-盘低-收盘图':90,
    '成交量-开盘-盘高-盘低-收盘图':91,
    '三维曲面图':83,
    '曲面图（俯视图）':85,
    '曲面图（俯视线框图）':86,
    '三维曲面图（线框）':84,
    '散点图':-4169,
    '折线散点图':74,
    '无数据点折线散点图':75,
    '平滑线散点图':72,
    '无数据点平滑线散点图':73,
}


# In[ ]:


# msoDashStyle
msoDashStyle = {
    'msoLineDash':4,
    'msoLineDashDot':5,
    'msoLineDashDotDot':6,
    'msoLineDashStyleMixed':-2,
    'msoLineLongDash':7,
    'msoLineLongDashDot':8,
    'msoLineRoundDot':3,
    'msoLineSolid':1,
    'msoLineSquareDot':2
}


 


# In[ ]:


# ppBulletType
ppBulletType = {
    'ppBulletMixed':-2, 
    'ppBulletNone':0,
    'ppBulletNumbered':2,
    'ppBulletPicture':3,
    'ppBulletUnnumbered':1,}


# In[ ]:


# xlLabelPosition
xlLabelPosition ={
    'xlLabelPositionAbove':0,
    'xlLabelPositionBelow':1,
    'xlLabelPositionBestFit':5,
    'xlLabelPositionCenter': -4108,
    'xlLabelPositionCustom':7,
    'xlLabelPositionInsideBase':4,
    'xlLabelPositionInsideEnd':3,
    'xlLabelPositionLeft':-4131,
    'xlLabelPositionMixed':6,
    'xlLabelPositionOutsideEnd':2,
    'xlLabelPositionRight':-4152,} 

