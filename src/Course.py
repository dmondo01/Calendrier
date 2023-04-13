# -*- coding: utf-8 -*-
#
# D.Mondou
# Course.py
# 26/09/2018
#


class Course(object):
    def __init__(self, module, type_course, date, begin_hour, end_hour, duration):
        self.m_begin_hour = begin_hour
        self.m_date = date
        self.m_duration = duration
        self.m_end_hour = end_hour
        self.m_module = module
        self.m_type = type_course

    def get_begin_hour(self):
        return self.m_begin_hour

    def get_date(self):
        return self.m_date

    def get_duration(self):
        return self.m_duration

    def get_end_hour(self):
        return self.m_end_hour

    def get_module(self):
        return self.m_module

    def get_type(self):
        return self.m_type

    def __lt__(self, other):
        return self.m_begin_hour <= other.get_begin_hour()
