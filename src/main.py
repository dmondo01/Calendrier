# -*- coding: utf-8 -*-
#
# D.Mondou
# main.py
# 26/09/2018
#

from TimeTable import TimeTable
from datetime import datetime

from src.TypeTeacher import TypeTeacher


def main():
    time_table = TimeTable("dmondo01", datetime(2022, 9, 1), 186, TypeTeacher.EC)
    time_table.create_time_table()
    time_table.createExcelTimeTable()


if __name__ == '__main__':
    main()
