#!/usr/bin/python
# -*- coding:utf-8 -*-
#----------------------------------------------------------------------------
# FileName:     ftp_download.py
# Description:  
# Author:       Wangjia
# Version:      0.0.1
# Created:      2019-05-17
# Company:      None
# LastChange:   create 2019-05-17
# History:      
#----------------------------------------------------------------------------

from ftplib import FTP
import os
import time
from test.test_ftplib import RETR_DATA

ftp=FTP()
ftp.set_debuglevel(2)
ftp.connect("192.168.1.11", 21)
ftp.login("ziji","abc.1234")

print(ftp.getwelcome())
#ftp.cwd(".//projects//ziji")
#dirnames=ftp.nlst()
dirnames=ftp.dir()
for dirname in dirnames:
    print(dirname)


'''更改当前目录 '''
os.chdir("D://资料//103_HXD记录//ato记录")
print(os.getcwd())

#获取时间戳
newDir=time.strftime('%Y%m%d',time.localtime(time.time()))+"记录"

#创建今天日期目录
#os.path.exist()
#os.makedirs(path) 创建多层目录
#os.mkdir(path) 创建目录
if not os.path.exists(newDir):
    os.makedirs(newDir+"//上午//上行")
    os.makedirs(newDir+"//上午//下行")
    os.makedirs(newDir+"//下午//上行")
    os.makedirs(newDir+"//下午//下行")
else:
    print("%s exist dir\n" % newDir)

DnNumbers=input("请输入下载文件数:\n")
print("下载文件数%s" % DnNumbers)
DnDirNumbers=input("请输入下载文件夹:\n 1 //上午//上行\n 2 //上午//下行\n 3 //下午//上行\n 4 //下午//下行\n")
print("下载文件目录%s" % DnDirNumbers)

dic={'1':"//上午//上行",'2':"//上午//下行",'3':"//下午//上行",'4':"//下午//下行"}

os.chdir(newDir+dic[DnDirNumbers])
print(dirnames[-1])
for i in range(int(DnNumbers)):
    fp=open(dirnames[len(dirnames)-1-i],"wb")
    ftp.retrbinary("RETR %s" %(dirnames[len(dirnames)-1-i]), fp.write, 1024)
    fp.close()

ftp.quit()
    