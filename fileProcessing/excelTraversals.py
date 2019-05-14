#!/usr/bin/python3
# -*- coding:utf-8 -*-
# @Time   : 2018/8/15 0015 19:00
# @Author : QiuQuan
# 遍历Excel文件的数据

import xlrd
import xlwt

workbook = xlrd.open_workbook("D:\\pythonCode\\fileProcessing\\EXCEL\\traversals.xlsx")
worksheet = workbook.sheets()
# print(workbook, worksheet)

newbook = xlwt.Workbook()

for sheet in worksheet:
    # print sheet
    # print(sheet.name)
    newsheet = newbook.add_sheet(sheet.name)
    # print(newsheet.name)
    nrow = sheet.nrows
    ncol = sheet.ncols
    # print(nrow, ncol)
    # for i in range(nrow):
    #     print(sheet.row_values(i))
    # for j in range(ncol):
    #     print(sheet.col_values(j))
    for rowloop in range(nrow):
        for colloop in range(ncol):
            # print(sheet.cell_value(rowloop, colloop))
            newsheet.write(rowloop, colloop, sheet.cell_value(rowloop, colloop))
newbook.save("D:\\pythonCode\\fileProcessing\\EXCEL\\traversals_new.xls")


