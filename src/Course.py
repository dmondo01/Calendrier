# -*- coding: utf-8 -*-
#
# D.Mondou
# Course.py
# 26/09/2018
#


class Course(object):
    def __init__(self, module, type, date, beginHour, endHour, duration):
        self.m_beginHour = beginHour
        self.m_date = date
        self.m_duration = duration
        self.m_endHour = endHour
        self.m_module = module
        self.m_type = type

    def getBeginHour(self):
        return self.m_beginHour

    def getDate(self):
        return self.m_date

    def getDuration(self):
        return self.m_duration

    def getEndHour(self):
        return self.m_endHour

    def getModule(self):
        return self.m_module

    def getType(self):
        return self.m_type

    def __lt__(self, other):
        return self.m_beginHour <= other.getBeginHour()

