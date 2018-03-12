# -*- coding:utf-8 -*-

# 将多个Excel文件合并成一个
import xlrd
import xlsxwriter


# 打开一个excel文件
def open_xls(file):
    fh = xlrd.open_workbook(file)
    return fh


# 获取excel中所有的sheet表
def getsheet(fh):
    return fh.sheets()


# 获取sheet表的行数
def getnrows(fh, sheet):
    table = fh.sheets()[sheet]
    return table.nrows


# 读取文件内容并返回行内容
def getFilect(file, shnum):
    fh = open_xls(file)
    table = fh.sheets()[shnum]
    num = table.nrows
    for row in range(num):
        rdata = table.row_values(row)
        datavalue.append(rdata)
    return datavalue


# 获取sheet表的个数
def getshnum(fh):
    x = 0
    sh = getsheet(fh)
    for sheet in sh:
        x += 1
    return x

# 处理路径字符串，将\转换为\\，避免转义错误,去掉多余的"".
def take_path(path_a):
    path_b = ''
    path_c = ''
    if path_a[0] == '\"':
        path_b = path_a.strip('''"''')
    else:
        path_b = path_a
    path_c = path_b.replace("\\", "\\\\")
    return path_c

# 采集要合并文件路径
def get_filepath():
    allxls = []
    active = True
    while active:
        temp = input("将要合并的excel文件拖入窗口,如需结束采集请输入“q”。")
        if temp == "q":
            active = False
        else:
            allxls.append(temp)
    return allxls


if __name__ == '__main__':
    # 定义要合并的excel文件列表
    # allxls=['F:/test/excel1.xlsx','F:/test/excel2.xlsx']
    # allxls = ["C:\ssd\hebing\新建文件夹\\1 车溪河社区民兵整组样表\\1复转退军人名册.xls",
    #           "C:\ssd\hebing\新建文件夹\\2大庙村民兵整组样表\\1复转退军人名册.xls",
    #           "C:\ssd\hebing\新建文件夹\\3大兴村民兵整组资料\\1复转退军人名册.xls"]
    allxls_old = get_filepath()
    allxls=[]
    for xls in allxls_old:
        xls = take_path(xls)
        allxls.append(xls)
        

    # 存储所有读取的结果
    datavalue = []
    for fl in allxls:
        fh = open_xls(fl)
        x = getshnum(fh)
        for shnum in range(x):
            print("正在读取文件：" + str(fl) + "的第" + str(shnum) + "个sheet表的内容...")
            rvalue = getFilect(fl, shnum)
    # 定义最终合并后生成的新文件
    endfile = take_path(input("请将新建的空文件拖入窗体，这个文件将是合并后的文件"))
    wb1 = xlsxwriter.Workbook(endfile)
    # 创建一个sheet工作对象
    ws = wb1.add_worksheet()
    for a in range(len(rvalue)):
        for b in range(len(rvalue[a])):
            c = rvalue[a][b]
            ws.write(a, b, c)
    wb1.close()
    print("文件合并完成，文件路径为" + endfile)
