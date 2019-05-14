#!/usr/bin/python3
# -*- coding:utf-8 -*-
# @Time   : 2018/8/15 0015 18:55
# @Author : QiuQuan
# 大文本文件切割，可以处理csv、txt等格式文件

import os
import time

def mkSubFile(lines, head, srcName, sub):
    [des_filename, extname] = os.path.splitext(srcName)
    filename = des_filename + '_' + str(sub) + extname
    print('make file: %s' % filename)
    fout = open(filename, 'w')
    try:
        fout.writelines([head])
        fout.writelines(lines)
        return sub + 1
    finally:
        fout.close()


def splitByLineCount(filename, count):
    fin = open(filename, 'r')
    try:
        head = fin.readline()
        buf = []
        sub = 1
        for line in fin:
            buf.append(line)
            if len(buf) == count:
                sub = mkSubFile(buf, head, filename, sub)
                buf = []
        if len(buf) != 0:
            sub = mkSubFile(buf, head, filename, sub)
    finally:
        fin.close()


if __name__ == '__main__':
    print("功能：大文本文件切割，可以处理csv、txt等格式文件。")
    # begin = time.time()
    route = input("请输入文件路径（eg:C:\\\**\\\**\\\*.txt）：")
    rSplit = input("请输入切割的行数：")
    a = int(rSplit)
    splitByLineCount(route, a)
    end = time.time()
    # print('time is %d seconds ' % (end - begin))
    os.system("pause")


