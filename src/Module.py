# -*- coding: utf-8 -*-
#
# D.Mondou
# Module.py
# 26/09/2018
#
from src.CourseType import CourseType


class Module(object):
    def __init__(self, name, code):
        self.m_name = name
        self.m_code = code
        self.m_courses = []
        self.m_cm_hour = 0
        self.m_td_hour = 0
        self.m_tea_hour = 0
        self.m_tp_hour = 0

    def add_course(self, course):
        self.m_courses.append(course)
        match course.get_type():
            case CourseType.CM:
                self.m_cm_hour += course.get_duration()
            case CourseType.TD:
                self.m_td_hour += course.get_duration()
            case CourseType.TP:
                self.m_tp_hour += course.get_duration()
            case CourseType.TEA:
                self.m_tea_hour += course.get_duration()

    def get_cm_hour(self):
        return self.m_cm_hour

    def get_courses(self):
        return self.m_courses

    def get_name(self):
        return self.m_name

    def get_code(self):
        return self.m_code

    def get_td_hour(self):
        return self.m_td_hour

    def get_tea_hour(self):
        return self.m_tea_hour

    def get_tp_hour(self):
        return self.m_tp_hour
