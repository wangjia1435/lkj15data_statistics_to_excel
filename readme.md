##  目的
1. 自动统计从http://222.240.61.122:8088/lmd-web/app/index.html 下载的当日 固定机车的LKJ数据，进行分析统计。
2. 下载对应记录后通过dmcl2012.exe软件解析成excel数据。

## 配置文件
config.xml:

![1574904572(1)](https://github.com/wangjia1435/lkj15data_statistics_to_excel/blob/master/src/picture/1574904572(1).png)

## 生成表格

![1574905642(1)](https://github.com/wangjia1435/lkj15data_statistics_to_excel/blob/master/src/picture/1574905642(1).png)

## 软件
**class类 excel_parse**
1. `AddExcelRow(self, RowList=[], Dir=u"上行")` 创建excel表格，并给最后一列添加内容


**class类 xml_config_parse**

1. `get_first_row_list(self,dir=u"上行")`解析xml格式获取第一行标题，上行与下行有区别
2. `get_coordinate(self,CoordinateX=1)`循环从CoordinateX行获取坐标，每调用一次列坐标加1
3. `get_station_list(self)`获取xml配置中的所有站台
4. `get_row_format(self,list=[],CoordinateX=1)`将坐标与列表元素组合成存储字典的列表.key为坐标，value为list元素值.

**class类 excel_rule**
该类主要定义每个输出标题内容的生成规则
1. `ElRu_GetKey(self,firstRow=0,lastRow=0, col=0, key=u"",outCol=0, frontToRear=True, fullPattern=True)`寻找关键内容的规则函数.
    - *firstRow*(从excel表的第firstRow行开始找)
    - *lastRow*（找到lastRow行结束）
    - *col*（找的是col列）
    - *key* 上述范围搜寻的关键字
    - *outCol*如果找到了，则输出该行的第outCol列内容，没找到输出not found
    - *frontToRear*是否从前往后找，或者从后往前找
    - *fullPattern*部分匹配还是全词匹配  

2. `ElRu_GetRowNum(self,firstRow=0,lastRow=0, col=0, key=u"",outCol=0, frontToRear=True, fullPattern=True)`获取关键字所在的行数，参数含义同上个函数
3. `ElRu_GetMaxRowNum(self)`获取excel表最大行数
4. `ElRu_GetRowList(self)`根据规则函数生成的关键字组成列表


## visual stdio下调试python代码
1. 利用shift+alt+f5直接在python交互中执行项目.这样结果很快出来.
2. 如果单步调试，启动会很慢.不知道怎么回事

xml格式下只有三种数据 tag(标签对象)， attri(字典对象-字典的key-value分别也是str对象)，text(str对象)
