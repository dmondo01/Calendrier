# -*- coding: utf-8 -*-
#
# D.Mondou
# main.py
# 26/09/2018
#

from TimeTable import TimeTable
from datetime import datetime
from TeacherType import TeacherType


if __name__ == '__main__':
    #time_table = TimeTable("LOGIN_ULR", 192, TeacherType.EC, datetime(2021, 9, 1), datetime(2022,8,31))
    # time_table = TimeTable("LOGIN_ULR", 384, TeacherType.PRAG, datetime(2022, 9, 1))
    time_table = TimeTable("dmondo01", 186, TeacherType.EC, datetime(2022, 9, 1))
    time_table.create_time_table()
    time_table.create_excel_time_table()
