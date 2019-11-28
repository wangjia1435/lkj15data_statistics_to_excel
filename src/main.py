#!/usr/bin/python
# -*- coding:utf-8 -*-
# FileName:     main.py
# Description:  excel表规则检索,寻找出指定范围内符合规则的关键字
# Author:       Wangjia
# Version:      0.0.1
# Created:      2019-05-27
# Company:      None
# LastChange:   create 2019-05-27
# History:      
#----------------------------------------------------------------------------

import openpyxl
from src import xml_config_parse
from src import excel_parse
from src.rule_retrieve import  excel_rule

xmlObj=xml_config_parse()

g_UpInFirstRowDict=xmlObj.get_row_format(xmlObj.get_first_row_list(u"上行"))
g_excelObj=excel_parse("../input","outputFileName.xlsx")
g_excelObj.AddExcelRow(g_UpInFirstRowDict,u"上行")

g_DnInFirstRowDict=xmlObj.get_row_format(xmlObj.get_first_row_list(u"下行"))
g_excelObj=excel_parse("../input","outputFileName.xlsx")
g_excelObj.AddExcelRow(g_DnInFirstRowDict,u"下行")

UpIndex=1
DnIndex=1
Index=0
while True:
    #1 获得待解析的excel名
    inputExcelName=g_excelObj.GetIterFileName()
    print(inputExcelName)
    if inputExcelName==None:
        break
    #2 创建规则类
    rule=excel_rule(inputExcelName)  
    #3 判断解析的excel的上下行
    dir=rule.ElRu_GetDir()

    
    if dir==u"上行":
        UpIndex=UpIndex+1
        Index=UpIndex
    else:
        DnIndex=DnIndex+1
        Index=DnIndex
    #4 解析excel规则得到行
    Row=rule.ElRu_GetRowList()
    #5 获得行与列字典列表
    RowDict=xmlObj.get_row_format(Row,Index)
    g_excelObj.AddExcelRow(RowDict,dir)
