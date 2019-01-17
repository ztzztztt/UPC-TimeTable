#!/usr/bin/env python
# encoding: utf-8
"""
@author: zhoutao
@license: (C) Copyright 2013-2017, Node Supply Chain Manager Corporation Limited.
@contact: zhoutao970226@gmail.com
@software: pycharm
@file: main.py
@time: 2019/1/16 16:18
@desc:
"""

from TimeTables import generateClassInfo, excelReader, getTimeTables


def main():
    while True:
        print("****************欢迎使用UPC课表生成工具-V1.0****************\nAuthor:zhoutao@s.upc.edu.cn")
        option = input("请选择功能:\n1):解析Excel课表文件\n2):生成课程信息JSON文件\n3):生成课表ICS文件\n4):退出")
        try:
            if option == '1':
                generateClassInfo.get_classInfo()
                pass
            elif option == '2':
                excelReader.main()
                pass
            elif option == '3':
                getTimeTables.main()
                pass
        except:
            exit(1)
        if option == '4':
            exit(0)
        pass


main()
