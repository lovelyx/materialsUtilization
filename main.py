
import tkinter.messagebox as msgbox
import traceback
from tkinter.ttk import Style

from xml.dom import minidom
import tkinter as tk
from tkinter import filedialog
import os
import re
import datetime as dt
from page.My_sheet import My_sheet
from page.style import style
import xlwt

# 判断为文件还是文件夹
def panduan(fullPath):
    # 首先遍历当前目录所有文件及文件夹
    file_list = os.listdir(fullPath)
    # 循环判断每个元素是否是文件夹还是文件，是文件夹的话，递归
    for file in file_list:
        # 利用os.path.join()方法取得路径全名，并存入cur_path变量，否则每次只能遍历一层目录
        cur_path = os.path.join(fullPath, file)
        # 判断是否是文件夹
        if os.path.isdir(cur_path):
            panduan(cur_path)
        else:
            listPath2.append(cur_path)

# 读取数据
def readData(Folderpath):
    panduan(Folderpath)
    com = re.compile('.*_C.saw')
    for i in listPath2:
        if com.match(i.split('\\')[-1]):
            listPath3.append(i)
    print(listPath3)
    for i in listPath3:  # 遍历所有xml文件
        # 打开文件
        dist = {}
        datalist = []
        with open(i, "r", encoding="gb18030") as f:
            for line in f:
                list2 = []
                list2.append(i.split('\\')[-1])
                # 按照逗号切割
                stringTOOL = line.split(",")
                # 判断每一行第一个，是什么
                if stringTOOL[0] == "MAT2":
                    MaterialName = stringTOOL[1]
                    # 工件名称
                    list2.append(MaterialName)
                    workpiecesNum = stringTOOL[47]
                    # 工件个数
                    list2.append(int(workpiecesNum))
                    workpiecesArea = stringTOOL[48]
                    # 工件面积
                    list2.append(float(workpiecesArea))
                    SurplusNum = stringTOOL[50]
                    # 余料数
                    list2.append(int(SurplusNum))
                    SurplusArea = stringTOOL[51]
                    # 余料总面积
                    list2.append(float(SurplusArea))
                    Utilization = (float(workpiecesArea) / float(SurplusArea))
                    result = '{:.2%}'.format(Utilization)
                    list2.append(result)
                    datalist.append(list2)
                    datalist2.append(list2)

            dist[i.split('\\')[-1]] = datalist
        dist2.update(dist)


# 写入数据到单元格
def writeCell():
    col = ('时间','批次号','材料名称', '工件数', '工件总面积', '使用余料板件数', '使用余料总面积','余料利用率','日利用率')
    test.write_row(0, 0, col,styleCell.style1())
    test.write_rows(1,1,datalist2,styleCell.style0())


# 合并单元格
def mergeCell():
    z=0
    i=0
    for row1,data in enumerate(dist2.values()):
        listKey=list(dist2.keys())
        num = len(data)
        for row2,data2 in enumerate(data):
            i = i + 1
        test.write_merge(i - num + 1, i, 1, 1, listKey[z])
        test.write_merge(1,i,0,0,time)
        test.write_merge(1,i,8, 8, test.MaterialresultSum, style.style2(self=""))
        z+=1
     # msgbox.showinfo("运行失败", "运行失败第{0}个文件错误".format())
if __name__ == '__main__':
 try:
    # 1. 读取xml文件
    # 设置弹窗
    FileNum = 1
    root = tk.Tk()
    root.withdraw()
    # 设置弹窗标题
    # 输入路径
    Folderpath = filedialog.askdirectory(title="选择输入目录")
    # 保存数据到字典，用于单元格合并时使用
    dist2 = {}
    # 保存数据到数组，用于写入数据
    datalist2=[]
    listPath2 = []
    listPath3 = []
    # 获取当前时间
    # time = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    time = dt.datetime.now().strftime("%m-%d")
    # 读取数据
    readData(Folderpath)
    # 创建表格对象
    test = My_sheet()
    # 创建样式对象
    styleCell = style()
    # 写入数据
    writeCell()
    # 合并单元格
    mergeCell()
    # 保存数据
    test.save("D:/{}余料利用率.xls".format(time))

    # 弹窗提示
    msgbox.showinfo("结束", "一共处理{0}个文件".format(len(listPath3)))

 except:
    print(traceback.format_exc())
    # 写入到 tb.txt 文件中
    traceback.print_exc(file=open('tb.txt', 'w+', encoding="utf-8"))
    msgbox.showinfo("运行失败", "文件处理失败，请联系管理员")
    # msgbox.showinfo("运行失败", "运行失败第{0}个文件错误".format())

