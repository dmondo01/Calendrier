# -*- coding: utf-8 -*-
#
# D.Mondou
# Module.py
# 26/09/2018
#


class Module(object):
    def __init__(self, name):
        self.m_name = name
        self.m_courses = []
        self.m_CMHour = 0
        self.m_TDHour = 0
        self.m_TEAHour = 0
        self.m_TPHour = 0

    def addCourse(self, course):
        self.m_courses.append(course)
        if course.getType() == "CM":
            self.m_CMHour += course.getDuration()
        else:
            if course.getType() == "TD":
                self.m_TDHour += course.getDuration()
            else:
                if course.getType() == "TP":
                    self.m_TPHour += course.getDuration()
                else:
                    if course.getType() == "TEA":
                        self.m_TEAHour += course.getDuration();

    def getCMHour(self):
        return self.m_CMHour

    def getCourses(self):
        return self.m_courses

    def getName(self):
        return self.m_name

    def getTDHour(self):
        return self.m_TDHour

    def getTEAHour(self):
        return self.m_TEAHour;

    def getTPHour(self):
        return self.m_TPHour

