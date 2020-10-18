# -*- coding: utf-8 -*-
#
# D.Mondou
# main.py
# 26/09/2018
#

from TimeTable import TimeTable
from datetime import datetime



def main():
    timeTable = TimeTable("dmondo01", datetime(2020, 9, 1), 192)
    timeTable.createTimeTable()
    timeTable.createExcelTimeTable()

if __name__ == '__main__':
    main()