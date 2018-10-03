# -*- coding: utf-8 -*-
#
# D.Mondou
# TimeTable.py
# 29/09/2018
#

from ics import Calendar
from datetime import datetime
from urllib2 import urlopen
import xlsxwriter
import os
from Module import Module
from Course import Course


class TimeTable(object):
    """
    :param login: login ULR pour acceder a l'EDT
    :param beginYearDate: date de debut pour la prise en compte des heures (ex: 2018-09-01)
    :param nbHoursPerform: Nombre d'heures a effectuer pour un service complet
    """
    def __init__(self, login, beginYearDate, nbHoursPerform):

        self.m_beginYearDate = beginYearDate
        self.m_ics = "https://apps.univ-lr.fr/cgi-bin/WebObjects/ServeurPlanning.woa/wa/iCalendarOccupations?login=" + login
        self.m_nbHoursPerform = nbHoursPerform
        self.m_modules = []
        self.m_courses = {}

    """
    Lecture de l'ICS et extraction des informations
    """
    def createTimeTable(self):
        gcal = Calendar(urlopen(self.m_ics).read().decode('iso-8859-1'))

        for component in gcal.events:
            description = component.description.replace(u"Ã¨", "e").replace(u"Ã©", "e")
            type = ""
            if "Tp : " in component.description:
                type = "TP"
            else:
                if "Td : " in component.description:
                    type = "TD"
                else:
                    if "Cm : " in component.description:
                        type = "CM"

            if type != "":
                index = description.find(":")
                temp = description[index + 2:]
                endIndex = temp.find(":")

                temp = temp[:endIndex - 7].replace(u"Ã¨", "e").replace(u"Ã©", "e").replace(u"Ãª", "e")
                parentheseIndex = temp.find("(")

                date = str(component.begin)[:10].split("-")
                d = datetime(int(date[0]), int(date[1]), int(date[2]))



                if d >= self.m_beginYearDate:
                    print "---------------------------------"
                    if parentheseIndex != -1:
                        temp = temp[:parentheseIndex - 1]

                    comaIndex = temp.find(",")
                    if comaIndex != -1:
                        temp = temp[:comaIndex - 1]

                    index = temp.find("-")
                    if index != -1:
                        temp = temp[:index-1]

                    print temp
                    print description
                    print d
                    print str(component.begin)[:10]
                    print str(component.begin)[11:19]
                    print component.duration

                    if temp == "Recuperatio":
                        temp = "Analyse de Donnees Mobile"

                    module =  None
                    for m in self.m_modules:
                        if m.getName() == temp:
                            module = m
                    if module is None:
                        module = Module(temp)

                    duration = 1.5
                    if str(component.duration) == "3:00:00":
                        duration = 3

                    m = None
                    try:
                        m = self.m_courses[d]
                    except:
                        m = []

                    course = Course(module, type, str(d), str(component.begin)[11:19], str(component.end)[11:19], duration)
                    module.addCourse(course)
                    m.append(course)
                    self.m_courses[d] = m
                    if module not in self.m_modules:
                        self.m_modules.append(module)

    """
    Creation du fichier excel recapitulatif
    """
    def createExcelTimeTable(self):
        workbook = xlsxwriter.Workbook(".." + os.sep + "service.xlsx")
        worksheet = workbook.add_worksheet()

        worksheet.set_column(0, 0, 8.33)
        worksheet.set_column(1, 1, 9.83)
        worksheet.set_column(2, 2, 14.33)
        worksheet.set_column(3, 3, 18.33)
        worksheet.set_column(4, 4, 18.33)
        worksheet.set_column(5, 5, 45.50)
        worksheet.set_column(6, 6, 9.67)
        worksheet.set_column(7, 7, 9.67)
        worksheet.set_column(8, 8, 9.67)
        worksheet.set_column(10, 10, 45.50)
        worksheet.set_column(11, 11, 9.83)
        worksheet.set_column(12, 12, 9.83)
        worksheet.set_column(13, 13, 9.83)
        worksheet.set_column(14, 14, 9.83)
        worksheet.set_column(15, 15, 9.83)

        # Start from the first cell. Rows and columns are zero indexed.
        row = 1
        col = 1

        cell_format = workbook.add_format()
        cell_format.set_align('center')

        worksheet.write(0, 1, "Semaine", cell_format)
        worksheet.write(0, 2, "Date", cell_format)
        worksheet.write(0, 3, u"Créneau", cell_format)
        worksheet.write(0, 4, "Niveau", cell_format)
        worksheet.write(0, 5, "UE", cell_format)
        worksheet.write(0, 6, "Nature", cell_format)
        worksheet.write(0, 7, "Total", cell_format)
        worksheet.write(0, 8, "HETD", cell_format)
        worksheet.write(0, 10, "UE", cell_format)
        worksheet.write(0, 11, "CM", cell_format)
        worksheet.write(0, 12, "TD", cell_format)
        worksheet.write(0, 13, "TP", cell_format)
        worksheet.write(0, 14, "Total", cell_format)
        worksheet.write(0, 15, "HETD", cell_format)

        totalDuration = 0
        totalDurationHETD = 0
        lastWeek = int(sorted(self.m_courses.keys())[0].strftime("%V"))
        nbCoursesByDay = 0

        merge_format = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter'})

        # Iterate over the module.
        for d in sorted(self.m_courses.keys()):
            week = int(d.strftime("%V"))
            if week != lastWeek:
                if nbCoursesByDay > 1:
                    worksheet.merge_range(row-nbCoursesByDay, col, row-1, col,lastWeek, merge_format)
                else:
                    worksheet.write_number(row, col, lastWeek)
                lastWeek = week
                nbCoursesByDay = len(self.m_courses[d])
            else:
                nbCoursesByDay += len(self.m_courses[d])



            date = str(d)[:10].split("-")
            if(len(self.m_courses[d]) > 1):
                worksheet.merge_range(row, col+1, row+len(self.m_courses[d])-1, col+1,
                                      date[2]+"/"+date[1]+"/"+date[0], merge_format)
            else:
                worksheet.write(row, col+1, date[2]+"/"+date[1]+"/"+date[0])
            for course in self.m_courses[d]:
                worksheet.write(row, col + 2, course.getBeginHour()[:-3].replace(":","h") + " - " + course.getEndHour()[:-3].replace(":","h"))
                worksheet.write(row, col + 3, "")
                worksheet.write(row, col + 4, course.getModule().getName())
                worksheet.write(row, col + 5, course.getType())
                worksheet.write_number(row, col + 6, course.getDuration())
                if course.getType() == "CM":
                    worksheet.write_number(row, col + 7, course.getDuration() * 1.5)
                    totalDurationHETD += course.getDuration() * 1.5
                else:
                    if course.getType() == "TD":
                        worksheet.write_number(row, col + 7, course.getDuration())
                        totalDurationHETD += course.getDuration()
                    else:
                        if course.getType() == "TP":
                            worksheet.write_number(row, col + 7, course.getDuration() * 2 / 3)
                            totalDurationHETD += course.getDuration() * 2 / 3
                row += 1
                totalDuration += course.getDuration()

        if nbCoursesByDay > 1:
            worksheet.merge_range(row - nbCoursesByDay, col, row - 1, col, lastWeek, merge_format)
        else:
            worksheet.write_number(row-1, col, lastWeek)

        worksheet.write(row+1, 6, "Total")
        worksheet.write_number(row+1, 7, totalDuration)
        worksheet.write_number(row+1, 8, totalDurationHETD)
        worksheet.write(row+2, 7, "Reste")
        worksheet.write_number(row+2, 8, self.m_nbHoursPerform-totalDurationHETD)

        # Recapitulatif par UE
        row = 1
        CM = 0
        TD = 0
        TP = 0
        for module in self.m_modules:
            worksheet.write(row, 10, module.getName())
            worksheet.write_number(row, 11, module.getCMHour())
            worksheet.write_number(row, 12, module.getTDHour())
            worksheet.write_number(row, 13, module.getTPHour())
            worksheet.write_number(row, 14, module.getCMHour()+module.getTDHour()+module.getTPHour())
            worksheet.write_number(row, 15, module.getCMHour()*1.5+module.getTDHour()+module.getTPHour()*2/3)
            row += 1
            CM += module.getCMHour()
            TD += module.getTDHour()
            TP += module.getTPHour()

        worksheet.write(row+1, 10, "Total", cell_format)
        worksheet.write_number(row+1, 11, CM, cell_format)
        worksheet.write_number(row+1, 12, TD, cell_format)
        worksheet.write_number(row+1, 13, TP, cell_format)
        worksheet.write_number(row+1, 14, CM+TD+TP, cell_format)
        worksheet.write_number(row+1, 15, CM*1.5+TD+TP*2/3, cell_format)

        workbook.close()