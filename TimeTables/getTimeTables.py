#!/usr/bin/env python
# encoding: utf-8
"""
@author: zhoutao
@license: (C) Copyright 2013-2017, Node Supply Chain Manager Corporation Limited.
@contact: zhoutao970226@gmail.com
@software: pycharm
@file: getTimeTables.py
@time: 2019/1/16 22:38
@desc:
"""

import sys
import json
import icalendar
import random
from datetime import datetime, timedelta

from pytz import UTC


# 初始化calendar对象
def initial(calendar):
    calendar.add("METHOD", "PUBLISH")
    calendar.add("VERSION", "2.0")
    calendar.add("X-WR-CALNAME", "课程表")
    calendar.add("PRODID", "-//Apple Inc.//Mac OS X 10.12//EN")
    calendar.add("X-APPLE-CALENDAR-COLOR", "#FC4208")
    calendar.add("X-WR-TIMEZONE", "Asia/Shanghai")
    calendar.add("CALSCALE", "GREGORIAN")
    # 创建Timezone对象
    timezone = icalendar.Timezone()
    timezone.add("TZID", "Asia/Shanghai")
    # 创建TimezoneStandard对象
    stander = icalendar.TimezoneStandard()
    # 实现添加TZOFFSETFROM：+0900
    stander.add("TZOFFSETFROM", timedelta(0.375))

    # 添加RRULE:FREQ=YEARLY;UNTIL=19910914T150000Z;BYMONTH=9;BYDAY=3SU
    arguements = {}
    arguements["FREQ"] = "YEARLY"
    arguements["UNTIL"] = datetime(1991, 9, 14, 15, 0, 0, tzinfo=UTC)
    arguements["BYMONTH"] = "9"
    arguements["BYDAY"] = "3SU"
    stander.add("RRULE", arguements)

    stander.add("DTSTART", datetime(1989, 9, 17, 0, 0, 0))
    stander.add("TZNAME", "GMT+8")
    # 实现添加TZOFFSETTO：+0800
    stander.add("TZOFFSETTO", timedelta(1 / 3))
    timezone.add_component(stander)
    daylight = icalendar.TimezoneDaylight()
    daylight.add("TZOFFSETFROM", timedelta(1 / 3))
    daylight.add("DTSTART", datetime(1991, 4, 14, 0, 0, 0))
    daylight.add("TZNAME", "GMT+8")
    daylight.add("TZOFFSETTO", timedelta(0.375))
    daylight.add("RDATE", datetime(1991, 4, 14, 0, 0, 0))
    timezone.add_component(daylight)
    calendar.add_component(timezone)
    return calendar
    pass


# 生成ics的文件内容
def generate(calendar, firstWeek):
    create_time = getCreateTime()
    with open("conf_classInfo.json", 'r') as f:
        data = json.load(f)
    classInfo = data["classInfo"]
    # 取出年、月、日
    year = int(firstWeek[0:4])
    month = int(firstWeek[4:6])
    day = int(firstWeek[6:8])
    print(len(classInfo))
    for i in range(len(classInfo)):
        # 生成初始时间
        first = datetime(year, month, day, 0, 0, 0)
        start = classInfo[i]["week"]["startWeek"]
        end = classInfo[i]["week"]["endWeek"]
        classTime = classInfo[i]["classTime"]
        weekday = classInfo[i]["weekday"]
        time = getClassTime(classTime)
        timeOfStart = time[0]
        timeOfEnd = time[1]
        for j in range(start-1, end):
            print(j)
            startTime = first + timedelta(days=weekday + j*7, hours=int(timeOfStart[0:2]), minutes=int(timeOfStart[2:4]))
            endTime = first + timedelta(days=weekday + j*7, hours=int(timeOfEnd[0:2]), minutes=int(timeOfEnd[2:4]))
            print(startTime, endTime)
            summary = classInfo[i]["className"] + "@" + classInfo[i]["classRoom"]
            # 创建事件
            uid = getUID()
            event = icalendar.Event()
            event['uid'] = uid
            event.add("CREATE", create_time)
            event.add("SUMMARY", summary)
            event.add("TRANSP", "OPAQUE")
            event.add("X-APPLE-TRAVEL-ADVISORY-BEHAVIOR", "AUTOMATIC")
            event.add('DTSTART', startTime)
            event.add('DTEND', endTime)
            event['DTSTAMP'] = create_time
            event.add("SEQUENCE", "0")
            # 创建ALARM
            alarmUID = getAlarmUID()
            alarm = icalendar.Alarm()
            alarm.add("X-WR-ALARMUID", alarmUID)
            alarm["TRIGGER"] = "-PT30M"
            alarm.add("DESCRIPTION", "事件提醒")
            alarm.add("ACTION", "DISPLAY")
            alarm.add("UID", "VfKKDbniQ8&zhoutao@s.upc.edu.cn")
            event.add_component(alarm)
            calendar.add_component(event)
            pass
        pass
    return calendar
    pass


# 获取文件生成的时间
def getCreateTime():
    return datetime.now().strftime("%Y%m%dT%H%M%SZ")
    pass


# 获取每个事件的唯一标识符uid
def getUID():
    ran = random.randrange(0, 999999999999, 11)
    uid = str(ran) + "zhoutao@s.upc.edu.cn"
    return uid
    pass


def getAlarmUID():
    alarmUID = random.randrange(0, 9999999999999999999999, 22)
    return str(alarmUID) + "&" + "zhoutao@s.upc.edu.cn"
    pass


# 获取到课程的时间
def getClassTime(classTime):
    with open("conf_classTime.json", "rb") as f:
        data = json.load(f)
    startTime = data["classTime"][classTime]["startTime"]
    endTime = data["classTime"][classTime]["endTime"]
    return [startTime, endTime]
    pass


def main():
    sys.path.append("..")
    firstWeek = input("请输入第一周的第一天(星期天) 例如20190224:")
    print(firstWeek)
    calendar = icalendar.Calendar()
    print("正在初始化...")
    calendar = initial(calendar)
    print("正在生成日历文件...")
    calendar = generate(calendar, firstWeek)
    with open('class.ics', 'wb') as f:
        f.write(calendar.to_ical())
    f.close()
    print("文件生成成功！！！")
    pass
