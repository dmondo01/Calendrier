# -*- coding: utf-8 -*-
#
# D.Mondou
# TimeTable.py
# 29/09/2018
#

import csv
import ssl

from ics import Calendar
from datetime import datetime
from urllib.request import urlopen
import xlsxwriter
import os
import certifi
from Module import Module
from Course import Course


class TimeTable(object):
    """
    :param login: login ULR pour acceder a l'EDT
    :param begin_year_date: date de debut pour la prise en compte des heures (ex: 2018-09-01)
    :param nb_hours_perform: Nombre d'heures a effectuer pour un service complet
    :param type_teacher: Type d'enseignant (EC, PRAG, ATER, ...)
    """

    def __init__(self, login, begin_year_date, nb_hours_perform, type_teacher):
        self.m_begin_year_date = begin_year_date
        self.m_ics = "https://apps.univ-lr.fr/cgi-bin/WebObjects/ServeurPlanning.woa/wa/iCalendarOccupations?login=" + login
        self.m_nb_hours_perform = nb_hours_perform
        self.m_type_teacher = type_teacher
        self.m_modules = []
        self.m_courses = {}
        self.m_maquette = {}

        self._load_courses_from_csv()

    """
    Lecture du fichier csv pour charger la liste des EC de licence et master
    """

    def _load_courses_from_csv(self):
        with open("./csv/maquette.csv", "r") as file:
            csvreader = csv.reader(file, delimiter=';')
            # On ignore le nom des colonnes
            next(csvreader)
            for row in csvreader:
                if len(row) == 2:
                    self.m_maquette[row[0]] = row[1]

    """
    Lecture de l'ICS et extraction des informations
    """

    def create_time_table(self):
        gcal = Calendar(
            urlopen(self.m_ics, context=ssl.create_default_context(cafile=certifi.where())).read().decode('iso-8859-1'))

        for component in gcal.events:
            date = str(component.begin)[:10].replace(" ", "").split("-")
            d = datetime(int(date[0]), int(date[1]), int(date[2]))

            if d >= self.m_begin_year_date:
                description = component.description.replace(u"Ã¨", "e").replace(u"Ã©", "e")

                course_type = ""
                if "tp : " in component.description.lower():
                    course_type = "TP"
                else:
                    if "td : " in component.description.lower():
                        course_type = "TD"
                    else:
                        if "cm : " in component.description.lower():
                            course_type = "CM"
                        else:
                            if "tea : " in component.description.lower():
                                course_type = "TEA"

                if course_type != "":
                    splited_description = description.split()
                    code = splited_description[3]

                    # Retrouver EC a partir de son code
                    try:
                        name_ec = self.m_maquette[code]
                    except:
                        name_ec = None

                    if name_ec == None:
                        index = description.find(":")
                        name_ec = description[index + 2:]
                        endIndex = name_ec.find(",")

                        name_ec = name_ec[:endIndex].replace(u"Ã¨", "e").replace(u"Ã©", "e").replace(u"Ãª", "e")
                        parentheseIndex = name_ec.find("(")

                        if parentheseIndex != -1:
                            name_ec = name_ec[:parentheseIndex - 1]

                        comaIndex = name_ec.find(",")
                        if comaIndex != -1:
                            name_ec = name_ec[:comaIndex - 1]

                        index = name_ec.find("-")
                        if index != -1:
                            name_ec = name_ec[:index - 1]

                    print("---------------------------------")
                    print(code + " " + name_ec)
                    print(str(component.begin)[:10])
                    print(str(component.begin)[11:19])
                    print(component.duration)

                    module = None
                    for m in self.m_modules:
                        if m.getCode() == code:
                            module = m
                    if module is None:
                        module = Module(name_ec, code)

                    duration = 1.5
                    if str(component.duration) == "1:00:00":
                        duration = 1
                    else:
                        if str(component.duration) == "2:00:00":
                            duration = 2
                        else:
                            if str(component.duration) == "2:30:00":
                                duration = 2.5
                            else:
                                if str(component.duration) == "3:00:00":
                                    duration = 3
                                else:
                                    if str(component.duration) == "4:30:00":
                                        duration = 4.5

                    m = None
                    try:
                        m = self.m_courses[d]
                    except:
                        m = []

                    course = Course(module, course_type, str(d), str(component.begin)[11:19], str(component.end)[11:19],
                                    duration)
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
        worksheet.set_column(3, 3, 45.50)
        worksheet.set_column(4, 4, 11)
        worksheet.set_column(5, 5, 11)
        worksheet.set_column(6, 6, 11)
        worksheet.set_column(7, 7, 9.67)
        worksheet.set_column(8, 8, 9.67)
        worksheet.set_column(10, 10, 50.50)
        worksheet.set_column(11, 11, 9.83)
        worksheet.set_column(12, 12, 9.83)
        worksheet.set_column(13, 13, 9.83)
        worksheet.set_column(14, 14, 9.83)
        worksheet.set_column(15, 15, 12)
        worksheet.set_column(16, 16, 15)

        # Start from the first cell. Rows and columns are zero indexed.
        row = 1
        col = 0

        cell_format = workbook.add_format()
        cell_format.set_align('center')

        worksheet.write(0, 0, "Semaine", cell_format)
        worksheet.write(0, 1, "Date", cell_format)
        worksheet.write(0, 2, u"Créneau", cell_format)
        worksheet.write(0, 3, "EC", cell_format)
        worksheet.write(0, 4, "Nature", cell_format)
        worksheet.write(0, 5, "Total", cell_format)
        worksheet.write(0, 6, "HETD", cell_format)
        worksheet.write(0, 10, "UE", cell_format)
        worksheet.write(0, 11, "CM", cell_format)
        worksheet.write(0, 12, "TD", cell_format)
        worksheet.write(0, 13, "TP", cell_format)
        worksheet.write(0, 14, "TEA", cell_format)
        worksheet.write(0, 15, "Total sans TEA", cell_format)
        worksheet.write(0, 16, "HETD sans TEA", cell_format)

        totalDuration = 0
        totalDurationHETD = 0
        if self.m_courses.__len__() != 0:
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
                        worksheet.merge_range(row - nbCoursesByDay, col, row - 1, col, lastWeek, merge_format)
                    else:
                        worksheet.write_number(row, col, lastWeek)
                    lastWeek = week
                    nbCoursesByDay = len(self.m_courses[d])
                else:
                    nbCoursesByDay += len(self.m_courses[d])

                date = str(d)[:10].split("-")
                if (len(self.m_courses[d]) > 1):
                    worksheet.merge_range(row, col + 1, row + len(self.m_courses[d]) - 1, col + 1,
                                          date[2] + "/" + date[1] + "/" + date[0], merge_format)
                else:
                    worksheet.write(row, col + 1, date[2] + "/" + date[1] + "/" + date[0])
                for course in sorted(self.m_courses[d]):
                    worksheet.write(row, col + 2,
                                    course.getBeginHour()[:-3].replace(":", "h") + " - " + course.getEndHour()[
                                                                                           :-3].replace(":", "h"))

                    # worksheet.write(row, col + 3, "")

                    worksheet.write(row, col + 3, course.getModule().getName())
                    worksheet.write(row, col + 4, course.getType())
                    worksheet.write_number(row, col + 5, course.getDuration())
                    if course.getType() == "CM":
                        worksheet.write_number(row, col + 6, course.getDuration() * 1.5)
                        totalDurationHETD += course.getDuration() * 1.5
                    else:
                        if course.getType() == "TD":
                            worksheet.write_number(row, col + 6, course.getDuration())
                            totalDurationHETD += course.getDuration()
                        else:
                            if course.getType() == "TP":
                                # worksheet.write_number(row, col + 7, course.getDuration() * 2 / 3)
                                worksheet.write_number(row, col + 6, course.getDuration())
                                # totalDurationHETD += course.getDuration() * 2 / 3
                                totalDurationHETD += course.getDuration()
                            # else:
                            # if course.getType() == "TEA":
                            # worksheet.write_number(row, col + 7, course.getDuration() * 0.015 * 15)
                            # totalDurationHETD += course.getDuration() * 0.015
                    row += 1
                    if course.getType() != "TEA":
                        totalDuration += course.getDuration()

            if nbCoursesByDay > 1:
                worksheet.merge_range(row - nbCoursesByDay, col, row - 1, col, lastWeek, merge_format)
            else:
                worksheet.write_number(row - 1, col, lastWeek)

        worksheet.write(row + 1, 4, "Total sans TEA")
        worksheet.write_number(row + 1, 5, totalDuration)
        worksheet.write_number(row + 1, 6, totalDurationHETD)
        worksheet.write(row + 2, 5, "Reste")
        worksheet.write_number(row + 2, 6, self.m_nb_hours_perform - totalDurationHETD)

        # Recapitulatif par UE
        row = 1
        CM = 0
        TD = 0
        TP = 0
        TEA = 0
        for module in self.m_modules:
            worksheet.write(row, 10, module.getCode() + " " + module.getName())
            worksheet.write_number(row, 11, module.getCMHour())
            worksheet.write_number(row, 12, module.getTDHour())
            worksheet.write_number(row, 13, module.getTPHour())
            worksheet.write_number(row, 14, module.getTEAHour())
            # worksheet.write_number(row, 15, module.getCMHour()+module.getTDHour()+module.getTPHour()+module.getTEAHour())
            # worksheet.write_number(row, 16, module.getCMHour()*1.5+module.getTDHour()+module.getTPHour()*2/3+module.getTEAHour()*0.015*15)
            worksheet.write_number(row, 15, module.getCMHour() + module.getTDHour() + module.getTPHour())
            # worksheet.write_number(row, 16, module.getCMHour() * 1.5 + module.getTDHour() + module.getTPHour() * 2 / 3)
            worksheet.write_number(row, 16, module.getCMHour() * 1.5 + module.getTDHour() + module.getTPHour())
            row += 1
            CM += module.getCMHour()
            TD += module.getTDHour()
            TP += module.getTPHour()
            TEA += module.getTEAHour()

        worksheet.write(row + 1, 10, "Total", cell_format)
        worksheet.write_number(row + 1, 11, CM, cell_format)
        worksheet.write_number(row + 1, 12, TD, cell_format)
        worksheet.write_number(row + 1, 13, TP, cell_format)
        worksheet.write_number(row + 1, 14, TEA, cell_format)
        # worksheet.write_number(row+1, 15, CM+TD+TP+TEA, cell_format)
        # worksheet.write_number(row+1, 16, CM*1.5+TD+TP*2/3+TEA*0.015*15, cell_format)
        worksheet.write_number(row + 1, 15, CM + TD + TP, cell_format)
        # worksheet.write_number(row + 1, 16, CM * 1.5 + TD + TP * 2 / 3, cell_format)
        worksheet.write_number(row + 1, 16, CM * 1.5 + TD + TP, cell_format)

        workbook.close()
