##  目的
1. 自动统计从http://222.240.61.122:8088/lmd-web/app/index.html下载的当日 固定机车的LKJ数据，进行分析统计。
2. 下载对应记录后通过dmcl2012.exe软件解析成excel数据。

## 配置文件
config.xml:

## 生成表格
关键字：

日期：<事件：文件头> <其他：日期>

机车号：<事件:机车号><其他: 机车号> 机车号不要AB

车次:<事件：车次><其他：>

司机号：<事件：司机1+司机2> <其他：>

吨位：<事件：总重> <其他： >

计长：<事件：计长> <其他： >

## 软件流程
1. 解析xml，保存excel第1列所用数据，使用字典形式存储
。变量名`g_InFirstRowDict{"ColName":"Content"}`
其中ColName的规则为 `[A-Z][A-Z][n]`,所以程序设计最多28\*28列;

- 输入为input文件夹下的excel.每次新加的excel存放其中.已经解析过的文件不再解析.解析过的在
ouput文件夹下的excel中会有记录.
- 如果输出excel中已经有第一列，那么同`g_InFirstRowDict`比较.采用`g_InFirstRowDict`为准


2. 解析输入excel.更新输出新的一行excel内容.变量名为`g_OutLastRowDict{"ColName":"Contend"}`其中ColName的规则为 `[A-Z][A-Z][n]`。

3. 解析Input文件夹excel判断该文件是否符合解析

4. 变量名`g_InFirstRowDict{"ColName":"Content"}`的`ColName`设计规则,根据规则去解析Input下有效excel,
获得`content`.根据输出excel文件获得写入的`content`二维坐标.

5. 设计规则
-  列名 “关键字” 
- 行初始 2
- 行结束 last
- 寻找关键字 "绿灯"

6. 变量解释
- `g_stationList`存放上行方向顺序的所有站名,如果开车方向为下行,那么`g_stationList`=`g_stationList.reverse()`
- `g_stationDicList`存放输入excel解析的所有站名对应的行号
- 侧线股道号、进站停车距、机外停车这个三个数值需要通过`g_stationDicList`找到行号

7. 将变量`g_OutLastRowDict{"ColName":"Content"}`更新到output文件夹excel的最新一行.

## visual stdio下调试python代码
1. 利用shift+alt+f5直接在python交互中执行项目.这样结果很快出来.
2. 如果单步调试，启动会很慢.不知道怎么回事

xml格式下只有三种数据 tag(标签对象)， attri(字典对象-字典的key-value分别也是str对象)，text(str对象)