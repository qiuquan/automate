#!/usr/bin/python3
# -*- coding:utf-8 -*-
# @Time   : 2018/8/15 0015 19:02
# @Author : QiuQuan
# Excel文件按照每8行切割为一个新的Excel文件

from __future__ import division
import xlrd
import xlwt

limit = 8
data = xlrd.open_workbook("D:\\Code\\Automate\\fileProcessing\\1.xls")
print(data.sheet_names())
table = data.sheet_by_index(0)
nrows = table.nrows
ncols = table.ncols
title_row = table.row_values(0)
print(nrows),
print(ncols)
sheets = nrows / limit
print(sheets)
if not sheets.is_integer():
    sheets = sheets + 1

for i in range(0, int(sheets)):
    workbook = xlwt.Workbook(encoding='ascii')
    worksheet = workbook.add_sheet(sheetname="i")
    for j in range(0, ncols):
        worksheet.write(0, j, label = title_row[j])
    for k in range(1, limit+1):
        newRow = k + i*limit
        if newRow < nrows:
            rowContent = table.row_values(newRow)
            for col in range(0, ncols):
                worksheet.write(k, col, rowContent[col])
    workbook.save(str(i) + ".xls")
