#!/usr/bin/env python
# encoding: utf-8
"""
@author: zhoutao
@license: (C) Copyright 2013-2017, Node Supply Chain Manager Corporation Limited.
@contact: zhoutao970226@gmail.com
@software: pycharm
@file: generateClassInfo.py
@time: 2019/1/14 14:35
@desc:
"""

import xlrd
import xlwt
import re
import sys


def get_TimeTable():
    sys.path.append("..")
    # 打开课表文件
    table = xlrd.open_workbook("kb.xlsx")
    # 读取第一个Sheet
    kb = table.sheets()[0]

    content = []
    col = kb.ncols
    # 循环遍历奈每一列,排除第一列，从第2列到第8列一次表示为每周周日到周六
    for i in range(1, 8):
        column = kb.col(i, 0, col)
        for one in column:
            # 获取每个单元格的内容
            cls = one.value
            if cls != '':
                # 以回车换行分割字符串，获得课程信息
                result = cls.split('\n')
                # 获取每个单元格的课程数目
                numOfCourses = int(len(result) / 4)
                # 没有课程。直接执行下一次循环
                if numOfCourses < 1:
                    continue
                # 取出每个单元格中的课程及其信息
                j = 0
                while j < numOfCourses:
                    # 课程名称
                    className = result[j * 4]
                    # 课程信息
                    class_info = result[j * 4 + 2]
                    # 处理课程信息,正则匹配课程周次、时间、教室等
                    rx = re.compile(r'(\d+-\d+)周\[(.*?)节\](.*)')
                    res = rx.findall(class_info)
                    # 上课周次、时间
                    Weekend = res[0][0]
                    ClassTime = res[0][1]
                    # 上课教室
                    classRoom = res[0][2]
                    # 处理上课周次、时间
                    weekend = handleWeekend(Weekend)
                    classTime = handleClassTime(ClassTime)
                    day = (i + 6) % 7
                    if day == 0:
                        day = 7
                    if type(weekend[0]) is list:
                        for k in range(0, len(weekend)):
                            content.append((className, weekend[0], weekend[1], str(day), classTime, classRoom))
                        if classTime == '7':
                            continue
                            pass
                    else:
                        content.append((className, weekend[0], weekend[1], str(day), classTime, classRoom))
                    # print(class_name, Weekend, classTime, classRoom)
                    j += 1
            pass
    return content
    pass


# 处理上课周次,包括1-16类型与1，2，3类型
def handleWeekend(weekend):
    result = weekend.split('-')
    if result[0] != weekend:
        handle_weekend = [result[0], result[1]]
    else:
        result = weekend.split(',')
        numOfWeek = len(result)
        handle_weekend = [[0, 0]] * numOfWeek
        for i in range(0, numOfWeek):
            handle_weekend[i] = [result[i], result[i]]
    return handle_weekend
    pass


# 处理上课时间，将其对应到相应的课程第几节
def handleClassTime(classtime):
    result = classtime.split('-')
    class_time = 0
    if result[0] != classtime:
        if int(result[1]) - int(result[0]) == 1:
            if result[0] == '01':
                class_time = 0
            if result[0] == '03':
                class_time = 1
            if result[0] == '05':
                class_time = 2
            if result[0] == '07':
                class_time = 3
            if result[0] == '09':
                class_time = 4
        if int(result[1]) - int(result[0]) == 3:
            if result[0] == '01':
                class_time = 5
            if result[0] == '05':
                class_time = 6
            if result[0] == '09':
                class_time = 7
        if int(result[1]) - int(result[0]) == 2:
            class_time = 7
    if result[0] == '091011':
        class_time = 7
    return str(class_time)
    pass


# 创建classInfo.xls文件
def get_classInfo():
    sys.path.append("..")
    title = [
        'className', 'startWeek', 'endWeek', 'weekday', 'classTime', 'classRoom'
    ]
    wbk = xlwt.Workbook()
    # 添加一个名为 课程表的sheet页
    sheet = wbk.add_sheet('sheet')
    for i in range(len(title)):  # 写入表头
        sheet.write(0, i, title[i])  # 写入每行,第一个值是行，第二个值是列，第三个是写入的值
    # 获取处理过的课程信息
    content = get_TimeTable()
    for i in range(len(content)):
        for j in range(len(content[0])):
            sheet.write(i + 1, j, content[i][j])
    # 保存数据到‘test.xls’文件中
    wbk.save('classInfo.xls')  # 保存excel必须使用后缀名是.xls的，不是能是.xlsx的


