#!/usr/bin/env python
# encoding: utf-8
"""
@author: zhoutao
@license: (C) Copyright 2013-2017, Node Supply Chain Manager Corporation Limited.
@contact: zhoutao970226@gmail.com
@software: pycharm
@file: excelReader.py
@time: 2019/1/16 14:51
@desc:
"""

import sys
import xlrd


# 生成conf_classInfo文件
def generateClassInfoConf():
    # 打开excel文件
    excel = xlrd.open_workbook("classInfo.xls")
    # 获取该课表信息sheet
    table = excel.sheets()[0]
    # 获取表格的行与列的个数
    rows = table.nrows
    cols = table.ncols
    # 遍历读取每个单元格
    classes_header = '{"classInfo"' + ": [\n"
    classes_tail = "]\n}"
    classes = ""
    for i in range(1, rows):
        className = table.cell(i, 0).value
        startWeek = table.cell(i, 1).value
        endWeek = table.cell(i, 2).value
        weekday = table.cell(i, 3).value
        classTime = table.cell(i, 4).value
        classRoom = table.cell(i, 5).value
        classes = classes + '{\n"className":"' + className + '",\n' + '"week":{\n"startWeek":' + startWeek \
                  + ",\n" + '"endWeek":' + endWeek + "\n},\n" + '"weekday":' + weekday + ",\n" \
                  + '"classTime":' + classTime + ",\n" + '"classRoom":"' + classRoom + '"\n}'
        if i != rows - 1:
            classes += ','
            pass
    classes = classes_header + classes + classes_tail
    return classes
    pass


def main():
    sys.path.append("..")
    print("Welcome to TimeTable Generating Tools V1.0 by ZhouTao!!!")
    option = input("Please Input 0 To Continue:")
    if option != '0':
        exit(0)
    print("Reading ""classInfo.xls"" File, Please Waiting")
    cla = generateClassInfoConf()
    print("Generating The ""conf_classInfo.json"" File ...")
    with open('conf_classInfo.json', 'w') as f:
        f.write(cla)
        f.close()
    print("\nALL DONE !")
    pass
