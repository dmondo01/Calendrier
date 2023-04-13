# -*- coding: utf-8 -*-
#
# D.Mondou
# main.py
# 26/09/2018
#

from TimeTable import TimeTable
from datetime import datetime
from src.TeacherType import TeacherType


if __name__ == '__main__':
    time_table = TimeTable("dmondo01", datetime(2022, 9, 1), 186, TeacherType.EC)
    time_table.create_time_table()
    time_table.create_excel_time_table()
