# -*- coding:utf-8 -*-
import os

def get_dir_input(str):
    path = input(str)
    path = path.strip('''"''')
    if os.path.isdir(path):
        path_a = path
    else:
        path_a = os.path.normpath(path)
    return path_a
#下面这些变量需要您根据自己的具体情况选择
#biaotou = ['序号','村','组',	'户主姓名','户主身份证号','家庭人口数','家庭成员姓名','性别'	,'原身份证号码','与户主关系','备注']  
biaotou = []
#在哪里搜索多个表格

filelocation = 'C:\\ssd\\工作\\移民扶贫工作资料夹\\扶贫资料\\7———减贫脱贫\\2018年脱贫\\各村减贫计划\\'
#filelocation = input("请将包含合并文件的文件夹拖入")
#filelocation = get_dir_input("www")

#当前文件夹下搜索的文件名后缀
fileform = "xlsx"  
#将合并后的表格存放到的位置
filedestination = 'C:\\ssd\\工作\\移民扶贫工作资料夹\\扶贫资料\\7———减贫脱贫\\2018年脱贫\\各村减贫计划\\汇总\\'
#合并后的表格命名为file
file = "城头山镇2018年预脱贫人口花名册"  
  
#首先查找默认文件夹下有多少文档需要整合
import glob  
from numpy import *  
filearray = []  
#for filename in glob.glob(filelocation + "*." + fileform):  
for filename in glob.glob(filelocation + "*.xls*" ):  
    filearray.append(filename)  
#以上是从pythonscripts文件夹下读取所有excel表格，并将所有的名字存储到列表filearray
print("在默认文件夹下有%d个文档哦" % len(filearray))  
ge = len(filearray)  
matrix = [None] * ge  
#实现读写数据
  
#下面是将所有文件读数据到三维列表cell[][][]中（不包含表头）
import xlrd  
for i in range(ge):  
    fname = filearray[i]  
    bk = xlrd.open_workbook(fname) 
    try:  
        sh = bk.sheet_by_name('2017年预脱贫花名册1')  
    except:  
        print("在文件%s中没有找到sheet1，读取文件数据失败,要不你换换表格的名字？" % fname)  
    nrows = sh.nrows   
    matrix[i] = [0] * (nrows - 1)  
      
    ncols = sh.ncols  
    for m in range(nrows - 1):    
        matrix[i][m] = ["0"] * ncols  
  
    for j in range(1,nrows):  
        for k in range(0,ncols):  
            matrix[i][j - 1][k] = sh.cell(j,k).value  
#下面是写数据到新的表格test.xls中哦
import xlwt  
filename = xlwt.Workbook()  
sheet = filename.add_sheet("hel")  
#下面是把表头写上
for i in range(0,len(biaotou)):  
    sheet.write(0,i,biaotou[i])  
#求和前面的文件一共写了多少行
zh = 1  
for i in range(ge):  
    for j in range(len(matrix[i])):  
        for k in range(len(matrix[i][j])):  
            sheet.write(zh,k,matrix[i][j][k])  
        zh = zh + 1  
print("我已经将%d个文件合并成1个文件，并命名为%s.xls.快打开看看正确不？" % (ge,file))  
filename.save(filedestination + file + ".xls")
