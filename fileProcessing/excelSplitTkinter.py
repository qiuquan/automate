#!/usr/bin/python3
# -*- coding:utf-8 -*-
# @Time   : 2018/12/13 0013 13:05
# @Author : QiuQuan
# 本软件主要用于Excel文件的分割，文件的第一行为标题行，分割条件为Excel第一列元素,分割后的文件存储在软件放置的位置。
# 使用步骤：
# 1.点击“选择Excel”,选择成功文本框会显示文件路径；\
# 2.点击"分割Excel",开始分割；\n3.文本框显示“分割成功”则代表处理完成。'

import tkinter
from tkinter import filedialog
import os
import xlwt
import xlrd

def quit():
     root.destroy()

def selectPath():
     filename = tkinter.filedialog.askopenfilename()
     var.set(filename)

def help():
    top = tkinter.Toplevel()
    top.title('说明')
    msg = tkinter.Message(top, anchor="center",aspect=2000, borderwidth= 10, width=300, text='本软件主要用于Excel文件的分割，文件的第一行为标题行，分割条件为Excel第一列元素,分割后的文件存储在软件放置的位置。\n\n使用步骤：\n1.点击“选择Excel”,选择成功文本框会显示文件路径；\
                                                                                             2.点击"分割Excel",开始分割；\n3.文本框显示“分割成功”则代表处理完成。')
    msg.pack()

# 集合是一个无序不重复元素集，重复元素在set中自动被过滤
def cleanDuplicate1(onelist):
    return list(set(onelist))

def TransferOfPrivateData():
     # root1 = tkinter.Tk()
     # root1.withdraw()

     # print("待处理文件输入示例(路径+名称)：D:\\\**\\\**.xls")
     # inputFilePath = input("请输入待处理文件：")
     # outputFilePath = input("请输入处理完毕文件存储路径：")
     # print(filePath)
     bookData = xlrd.open_workbook(var.get())
     sheetData = bookData.sheets()[0]
     # print(sheetData)
     nrows = sheetData.nrows
     ncols = sheetData.ncols
     # print(nrows, ncols)
     title_row = sheetData.row_values(0)
     # print(title_row)
     zoneList = sorted(cleanDuplicate1(sheetData.col_values(0)[1:]))
     # print(len(zoneList))

     rowFlag = 0
     # print(sheetData.col_values(0)[1])
     for zone in zoneList:
          newbook = xlwt.Workbook()
          newsheet = newbook.add_sheet("zone")
          # print(newbook.add_sheet(zone).name)
          for title in range(0, len(title_row)):
               newsheet.write(0, title, title_row[title])
          for rown in range(1, nrows):
               if zone == sheetData.row_values(rown)[0]:
                    rowFlag = rowFlag + 1
                    for coln in range(0, ncols):
                         newsheet.write(rowFlag, coln, sheetData.cell_value(rown, coln))
          newbook.save(str(zone) + ".xls")
          rowFlag = 0

     var.set("分割成功！！")

if __name__ == '__main__':

     root = tkinter.Tk()
     root.title("Excel文件分割（QQ）")
     root.geometry("450x200+450+210")

     frame = tkinter.Frame()
     frame.pack()

     frm_L = tkinter.Frame(frame)
     tkinter.Label(frm_L).pack()
     var = tkinter.StringVar()
     E1 = tkinter.Entry(frm_L, textvariable=var, bd=1, width=50)
     E1.pack(expand=1)
     tkinter.Label(frm_L).pack()
     frm_L.pack()

     frm_R = tkinter.Frame(frame)
     tkinter.Button(frm_R, text="选择Excel", command=selectPath, height=1, width=10).pack(side=tkinter.LEFT)
     tkinter.Button(frm_R, text="分割Excel", command=TransferOfPrivateData, height=1, width=10).pack(side=tkinter.RIGHT)
     frm_R.pack()

     menubar = tkinter.Menu(root)
     menubar.add_command(label="说明", command=help)
     root["menu"] = menubar

     tkinter.Label(root, text="").pack()
     tkinter.Button(root, text="退出", command=quit, height=1, width=8).pack()

     root.mainloop()