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
from CourseType import CourseType
from TeacherType import TeacherType

"""
Récupération du nom de l'EC si non présent dans la maquette csv
:param description : Description d'un événement du calendrier
:return : Le nom de l'EC
"""


def _get_name_ec(description):
    index = description.find(":")
    name_ec = description[index + 2:]
    end_index = name_ec.find(",")

    name_ec = name_ec[:end_index]
    bracket_index = name_ec.find("(")

    if bracket_index != -1:
        name_ec = name_ec[:bracket_index - 1]

    coma_index = name_ec.find(",")
    if coma_index != -1:
        name_ec = name_ec[:coma_index - 1]

    index = name_ec.find("-")
    if index != -1:
        name_ec = name_ec[:index - 1]

    return name_ec


"""
Récupération du type du cours (CM, TD, TP ou TEA)
:param description : Description d'un événement du calendrier
:return : Le type du cours
"""


def _get_course_type(description):
    if "tp : " in description:
        return CourseType.TP
    elif "td : " in description:
        return CourseType.TD
    elif "cm : " in description:
        return CourseType.CM
    elif "tea : " in description:
        return CourseType.TEA
    return ""


class TimeTable(object):
    """
    :param login: login ULR pour acceder a l'EDT
    :param nb_hours_perform: Nombre d'heures a effectuer pour un service complet
    :param type_teacher: Type d'enseignant (EC, PRAG, ATER, ...)
    :param begin_date: date de debut pour la prise en compte des heures (ex: 2018-09-01)
    :param end_date: date de fin optionnelle pour la prise en compte des heures (ex: 2019-08-31)
    """

    def __init__(self, login, nb_hours_perform, type_teacher, begin_date, end_date=None):
        self.m_begin_date = begin_date
        self.m_end_date = end_date
        self.m_ics = "https://apps.univ-lr.fr/cgi-bin/WebObjects/ServeurPlanning.woa/wa/iCalendarOccupations?login=" + login
        self.m_nb_hours_perform = nb_hours_perform
        self.m_type_teacher = type_teacher
        self.m_modules = {}
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

            # Si (date de fin presente  et date fin >= d >= date debut) ou (pas de date de fin et d >= date debut)
            if (self.m_end_date is not None and self.m_end_date >= d >= self.m_begin_date) or (
                    self.m_end_date is None and d >= self.m_begin_date):
                description = component.description.replace(u"Ã¨", "e").replace(u"Ã©", "e").replace(u"Ã", "à")

                course_type = _get_course_type(component.description.lower())
                splited_description = description.split()

                if len(splited_description) >= 4:
                    code = splited_description[3]
                else:
                    code = ""

                if course_type != "" and code != "" and not code.__contains__("(") and not code.__contains__(
                        "PCM") and not code.__contains__("Examen"):
                    # Retrouver EC a partir de son code
                    try:
                        name_ec = self.m_maquette[code]
                    except KeyError:
                        name_ec = _get_name_ec(description)

                    print("---------------------------------")
                    print(code + " " + name_ec)
                    print(str(component.begin)[:10])
                    print(str(component.begin)[11:19])
                    print(component.duration)

                    try:
                        module = self.m_modules[code]
                    except KeyError:
                        module = Module(name_ec, code)

                    duration = component.duration.seconds / 3600

                    try:
                        m = self.m_courses[d]
                    except KeyError:
                        m = []

                    course = Course(module, course_type, str(d), str(component.begin)[11:19], str(component.end)[11:19],
                                    duration)
                    module.add_course(course)
                    m.append(course)
                    self.m_courses[d] = m
                    if code not in self.m_modules.keys():
                        self.m_modules[code] = module

    """
    Creation du fichier excel recapitulatif
    """

    def create_excel_time_table(self):
        workbook = xlsxwriter.Workbook(".." + os.sep + "service.xlsx")
        worksheet = workbook.add_worksheet()

        worksheet.set_column(0, 0, 8.33)
        worksheet.set_column(1, 1, 9.83)
        worksheet.set_column(2, 2, 14.33)
        worksheet.set_column(3, 3, 45.50)
        worksheet.set_column(4, 4, 12)
        worksheet.set_column(5, 5, 12)
        worksheet.set_column(6, 6, 12)
        worksheet.set_column(7, 7, 9.67)
        worksheet.set_column(8, 8, 9.67)
        worksheet.set_column(10, 10, 53.50)
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
        worksheet.write(0, 10, "EC", cell_format)
        worksheet.write(0, 11, "CM", cell_format)
        worksheet.write(0, 12, "TD", cell_format)
        worksheet.write(0, 13, "TP", cell_format)
        worksheet.write(0, 14, "TEA", cell_format)
        worksheet.write(0, 15, "Total (horsTEA)", cell_format)
        worksheet.write(0, 16, "HETD (hors TEA)", cell_format)

        total_duration = 0
        total_duration_hetd = 0
        full_service = False
        # Comptage ministere
        total_extra_hour = 0

        if self.m_courses.__len__() != 0:
            last_week = int(sorted(self.m_courses.keys())[0].strftime("%V"))
            nb_courses_by_day = 0

            merge_format = workbook.add_format({
                'align': 'left',
                'valign': 'vcenter'})

            # Iterate over the module.
            for d in sorted(self.m_courses.keys()):
                week = int(d.strftime("%V"))
                if week != last_week:
                    if nb_courses_by_day > 1:
                        worksheet.merge_range(row - nb_courses_by_day, col, row - 1, col, last_week, merge_format)
                    else:
                        worksheet.write_number(row - 1, col, last_week, merge_format)
                    last_week = week
                    nb_courses_by_day = len(self.m_courses[d])
                else:
                    nb_courses_by_day += len(self.m_courses[d])

                date = str(d)[:10].split("-")
                if len(self.m_courses[d]) > 1:
                    worksheet.merge_range(row, col + 1, row + len(self.m_courses[d]) - 1, col + 1,
                                          date[2] + "/" + date[1] + "/" + date[0], merge_format)
                else:
                    worksheet.write(row, col + 1, date[2] + "/" + date[1] + "/" + date[0], merge_format)
                for course in sorted(self.m_courses[d]):
                    worksheet.write(row, col + 2,
                                    course.get_begin_hour()[:-3].replace(":", "h") + " - " + course.get_end_hour()[
                                                                                             :-3].replace(":", "h"))
                    worksheet.write(row, col + 3, course.get_module().get_name())
                    worksheet.write(row, col + 4, str(course.get_type())[-3:].replace(".", ""))
                    worksheet.write_number(row, col + 5, course.get_duration())
                    if course.get_type() == CourseType.CM:
                        worksheet.write_number(row, col + 6, course.get_duration() * 1.5)
                        total_duration_hetd += course.get_duration() * 1.5
                        total_extra_hour += course.get_duration() * 1.5
                    else:
                        if course.get_type() == CourseType.TD:
                            worksheet.write_number(row, col + 6, course.get_duration())
                            total_duration_hetd += course.get_duration()
                            total_extra_hour += course.get_duration()
                        else:
                            if course.get_type() == CourseType.TP:
                                if self.m_type_teacher == TeacherType.EC or self.m_type_teacher == TeacherType.PRAG or self.m_type_teacher == TeacherType.PRCE:
                                    worksheet.write_number(row, col + 6, course.get_duration())
                                    total_duration_hetd += course.get_duration()
                                    total_extra_hour += course.get_duration() * 2 / 3
                                else:
                                    worksheet.write_number(row, col + 6, course.get_duration() * 2 / 3)
                                    total_duration_hetd += course.get_duration() * 2 / 3

                    if total_duration_hetd >= self.m_nb_hours_perform and not full_service:
                        color_format = workbook.add_format({'bold': True, 'bg_color': 'red'})
                        worksheet.write(row, 7, "Service du atteint (sans prise en compte du TEA)", color_format)
                        full_service = True
                        total_extra_hour = 0

                    row += 1
                    if course.get_type() != CourseType.TEA:
                        total_duration += course.get_duration()

            if nb_courses_by_day > 1:
                worksheet.merge_range(row - nb_courses_by_day, col, row - 1, col, last_week, merge_format)
            else:
                worksheet.write_number(row - 1, col, last_week, merge_format)

        worksheet.write(row + 1, 4, "Total (hors TEA)")
        worksheet.write_number(row + 1, 5, total_duration)
        worksheet.write_number(row + 1, 6, total_duration_hetd)

        if self.m_nb_hours_perform - total_duration_hetd > 0:
            worksheet.write(row + 3, 5, "Reste")
            worksheet.write_number(row + 3, 6, self.m_nb_hours_perform - total_duration_hetd)

        # Recapitulatif par UE
        row = 1
        cm = 0
        td = 0
        tp = 0
        tea = 0
        for module in self.m_modules.values():
            worksheet.write(row, 10, module.get_code() + " " + module.get_name())
            worksheet.write_number(row, 11, module.get_cm_hour())
            worksheet.write_number(row, 12, module.get_td_hour())
            worksheet.write_number(row, 13, module.get_tp_hour())
            worksheet.write_number(row, 14, module.get_tea_hour())
            worksheet.write_number(row, 15, module.get_cm_hour() + module.get_td_hour() + module.get_tp_hour())

            if self.m_type_teacher == TeacherType.EC or self.m_type_teacher == TeacherType.PRAG or self.m_type_teacher == TeacherType.PRCE:
                worksheet.write_number(row, 16,
                                       module.get_cm_hour() * 1.5 + module.get_td_hour() + module.get_tp_hour())
            else:
                worksheet.write_number(row, 16,
                                       module.get_cm_hour() * 1.5 + module.get_td_hour() + module.get_tp_hour() * 2 / 3)
            row += 1
            cm += module.get_cm_hour()
            td += module.get_td_hour()
            tp += module.get_tp_hour()
            tea += module.get_tea_hour()

        worksheet.write(row + 1, 10, "Total", cell_format)
        worksheet.write_number(row + 1, 11, cm, cell_format)
        worksheet.write_number(row + 1, 12, td, cell_format)
        worksheet.write_number(row + 1, 13, tp, cell_format)
        worksheet.write_number(row + 1, 14, tea, cell_format)
        worksheet.write_number(row + 1, 15, cm + td + tp, cell_format)

        worksheet.write(row + 2, 10, "Total HETD (hors TEA)", cell_format)
        worksheet.write_number(row + 2, 11, cm * 1.5, cell_format)
        worksheet.write_number(row + 2, 12, td, cell_format)
        if self.m_type_teacher == TeacherType.EC or self.m_type_teacher == TeacherType.PRAG or self.m_type_teacher == TeacherType.PRCE:
            worksheet.write_number(row + 2, 13, tp, cell_format)
            worksheet.write_number(row + 2, 16, cm * 1.5 + td + tp, cell_format)
        else:
            worksheet.write_number(row + 2, 13, tp * 2 / 3, cell_format)
            worksheet.write_number(row + 2, 16, cm * 1.5 + td + tp * 2 / 3, cell_format)

        if self.m_type_teacher == TeacherType.EC or self.m_type_teacher == TeacherType.PRAG or self.m_type_teacher == TeacherType.PRCE:
            worksheet.write_number(row + 1, 16, cm * 1.5 + td + tp, cell_format)
            # Heures supplementaires
            if (cm * 1.5 + td + tp) > self.m_nb_hours_perform:
                worksheet.write(row + 4, 10, u"Heures supplémentaires (hors TEA) - Comptage Université", cell_format)
                worksheet.write(row + 5, 10, u"Heures supplémentaires (hors TEA) - Comptage Ministère", cell_format)

                extra_hour = 0
                if self.m_type_teacher == TeacherType.EC:
                    cmtd = (cm * 1.5) + td

                    if cmtd >= self.m_nb_hours_perform:
                        extra_hour = abs(self.m_nb_hours_perform - cmtd)
                        extra_hour += tp * 2 / 3
                    else:
                        extra_hour = abs((self.m_nb_hours_perform - cmtd - tp) * 2 / 3)
                elif self.m_type_teacher == TeacherType.PRCE or self.m_type_teacher == TeacherType.PRAG:
                    extra_hour = abs(self.m_nb_hours_perform - (cm * 1.5) - td - tp)

                worksheet.write_number(row + 4, 11, extra_hour, cell_format)
                worksheet.write_number(row + 5, 11, total_extra_hour, cell_format)
        else:
            worksheet.write_number(row + 1, 16, cm * 1.5 + td + tp * 2 / 3, cell_format)

        workbook.close()
