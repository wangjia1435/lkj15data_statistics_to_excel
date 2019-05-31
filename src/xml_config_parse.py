#!/usr/bin/python
# -*- coding:utf-8 -*-
#----------------------------------------------------------------------------
# FileName:     xml_config_parse.py
# Description:  see https://docs.python.org/3/library/xml.etree.elementtree.html?highlight=elementtree
#               字典操作 see https://blog.csdn.net/hhtnan/article/details/77164198
# Author:       Wangjia
# Version:      0.0.1
# Created:      2019-05-17
# Company:      None
# LastChange:   create 2019-05-17
# History:      
#----------------------------------------------------------------------------

try:
    import xml.etree.cElementTree as ET
except ImportError:
    import xml.etree.ElementTree as ET
class xml_config_parse():
    """description of class """
    #第一行表格
    s_InFirstRowDict={} 
    s_InputPath=""
    s_OutputPathFile=""
    s_CoordinateX=1 #ROW
    s_CoordinateY1="A"
    s_CoordinateY2=""
    s_Bool=0
    def __init__(self):
        pass
    #坐标循环累加test OK
    def get_coordinate(self,CoordinateX=1):
        
        y1=self.s_CoordinateY1
        y2=self.s_CoordinateY2

        if(y1=='Z' and y2=="Z"):
            print("out range\n")
            self.s_CoordinateY1="A"
            self.s_CoordinateY2=""
            return self.s_CoordinateY1+str(CoordinateX)

        if(self.s_Bool==0):
            self.s_Bool=1
            return y1+str(CoordinateX)

        if(y1=='Z' and y2==""):
            y1=ord('A')
            y2=ord('A')
            # chr 整数转换成字符
            y1=chr(y1)
            y2=chr(y2)
            self.s_CoordinateY1=y1
            self.s_CoordinateY2=y2
            return y2+y1+str(CoordinateX)
        elif(y1=='Z'):
            y1=ord('A')
            y2=ord(self.s_CoordinateY2)+1
            y1=chr(y1)
            y2=chr(y2)
            self.s_CoordinateY1=y1
            self.s_CoordinateY2=y2
            return y2+y1+str(CoordinateX) 
        elif(y1!="Z" and y2==""):
            y1=ord(self.s_CoordinateY1)+1
            y1=chr(y1)
            self.s_CoordinateY1=y1
            return y1+str(CoordinateX)
        else:
            #ord 字符转换成整数
            y1=ord(self.s_CoordinateY1)+1
            y1=chr(y1)
            self.s_CoordinateY1=y1
            return y2+y1+str(CoordinateX)     
        pass

    # 获取第一行列表内容
    def get_first_row_list(self,dir=u"上行"):
        s_colList=[]

        tree = ET.parse(r'config.xml')
        root = tree.getroot()

        """finding intersting elements"""
        for baseCol in root.iter('列'):
            if baseCol!=None:
                _dic={}
                for item,value in baseCol.attrib.items():
                    #print(_dic)
                    s_colList.append(item)
        print(s_colList)
        tempS=[]
        for StationTag in root.iter('station'):
            if StationTag!=None:
                #print('root-tag:',StationTag.tag,',root-attrib:',StationTag.attrib,',root-text:',StationTag.text)
                tempL=[]

                for subTag in StationTag:
                    _dic={}
                    _stationName=""
                    _stationSonDiscrip=""
                    for item,value in StationTag.attrib.items():
                        _stationName=value

                    #item代表key value代表value
                    for item,value in subTag.attrib.items():
                        _stationSonDiscrip=_stationName+item
                        
                    tempL.append(_stationSonDiscrip)
                tempS=tempS+tempL
        print(tempS)
        if dir == u"上行":
            s_colList=s_colList+(tempS) 
        else:
            tempS.reverse()
            # ['神池南机外停车', '神池南进站停车距离', '神池南侧线股道号'
            #    改成['神池南侧线股道号', '神池南进站停车距离',  '神池南机外停车', 和上行保持一致
            for i in range(0,int(len(tempS)/3)):
                temp=tempS[i*3]
                tempS[i*3]=tempS[i*3+2]
                tempS[i*3+2]=temp
                pass
            print(tempS)   
            s_colList=s_colList+(tempS) 
        return s_colList        
    
    
    # 获取站台名列表
    def get_station_list(self):
        s_stationList=[]
        tree = ET.parse(r'config.xml')
        root = tree.getroot()

        for allStationTag in root.iter('station'):
            if allStationTag!=None:
                #print('root-tag:',allStationTag.tag,',root-attrib:',allStationTag.attrib,',root-text:',allStationTag.text)
                for item,value in allStationTag.attrib.items():
                    _stationName=value
                s_stationList.append(_stationName)
        return s_stationList
    
    # 返回具体行字典格式列表
    def get_row_format(self,list=[],CoordinateX=1):
        '''
        param:list 行中填写的内容
        param: x   行号
        '''
        #纵坐标重新开始
        self.s_Bool=0
        self.s_CoordinateY1="A"
        self.s_CoordinateY2=""        
        _colDicList=[]
        for item in list:
            if item!=None:
                _dic={}
                _dic[self.get_coordinate(CoordinateX)]=item
                _colDicList.append(_dic)
        return (_colDicList)
def main():    
    l=xml_config_parse()
    print(l.get_row_format(l.get_first_row_list(u"下行"), CoordinateX=2))
    pass

if __name__ == '__main__':
    main()

