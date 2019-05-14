#!/usr/bin/python3
# -*- coding:utf-8 -*-
# @Time   : 2018/8/15 0015 19:01
# @Author : QiuQuan
# 按照第一列元素将Excel文件分割成许多Excel文件

import xlwt
import xlrd
import os


# 集合是一个无序不重复元素集，重复元素在set中自动被过滤
def cleanDuplicate1(onelist):
    return list(set(onelist))

# 字典所有keys是不能重复的
def cleanDuplicate2(onelist):
    return {}.fromkeys(onelist).keys()

# 遍历列表
def cleanDuplicate3(onelist):
    templist = []
    for one in onelist:
        if one not in templist:
            templist.append(one)
    return templist

# 排序方法
def cleanDuplicate4(onelist):
    resultlist = []
    templist = sorted(onelist)
    i = 0
    while i < len(templist):
        if templist[i] not in resultlist:
            resultlist.append(templist[i])
        else:
            i += 1
    return resultlist

print("待处理文件输入示例(路径+名称)：D:\\\**\\\**.xls")
inputFilePath = input("请输入待处理文件：")
# outputFilePath = input("请输入处理完毕文件存储路径：")
# print(filePath)
bookData = xlrd.open_workbook(inputFilePath)
sheetData = bookData.sheets()[0]
# print(sheetData)
nrows = sheetData.nrows
ncols = sheetData.ncols
# print(nrows, ncols)
title_row = sheetData.row_values(0)
# print(title_row)
zoneList = sorted(cleanDuplicate3(sheetData.col_values(0)[1:]))
# print(len(zoneList))

rowFlag = 0
# print(sheetData.col_values(0)[1])
for zone in zoneList:
    # print(zone)
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
print("处理完毕！")
# os.system("pause")