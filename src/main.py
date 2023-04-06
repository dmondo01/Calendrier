# -*- coding: utf-8 -*-
#
# D.Mondou
# main.py
# 26/09/2018
#

from TimeTable import TimeTable
from datetime import datetime


def main():
    time_table = TimeTable("jricha03", datetime(2022, 9, 1), 192)
    time_table.createTimeTable()
    time_table.createExcelTimeTable()


if __name__ == '__main__':
    main()
