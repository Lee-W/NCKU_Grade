import re
import getpass
from html.parser import HTMLParser
from collections import OrderedDict

import requests
import xlsxwriter

from HTML_Form_Parser.HTML_form_parser import HTMLFormParser


class NckuGradeCrawler:
    MAIN_URL = "http://140.116.165.71:8888/ncku/"
    LOGIN_URL = MAIN_URL+"qrys02.asp"
    LOGOUT_URL = MAIN_URL+"logouts.asp"
    INDEX_URL = MAIN_URL+"/qrys05.asp"
    ENCODING = "big5"
    HEADER = {'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
              'X-Requested-With': 'XMLHttpRequest'}

    def __init__(self):
        self.session = requests.session()
        self.data = dict()
        self.overall_summary = OrderedDict()
        self.all_semester_data = OrderedDict()
        self.semeseters = list()

    def login(self):
        self.session.post(NckuGradeCrawler.LOGIN_URL, data=self.data)

    def logout(self):
        self.session.post(NckuGradeCrawler.LOGOUT_URL)

    def set_stu_info(self, stu_id, passwd):
        self.data = {'ID': stu_id.upper(), 'PWD': passwd}

    def parse_all_semester_data(self):
        self.__parse_index_page()

        self.all_semester_data = OrderedDict()
        for sem in self.semeseters:
            s_name = sem[:4] + ("2" if "¤U" in sem else "1")
            self.all_semester_data[s_name] = self.__parse_semester_data(sem)
        self.__overall_summerize()

    def __parse_index_page(self):
        req = self.session.post(NckuGradeCrawler.INDEX_URL,
                                data=self.data,
                                cookies=self.session.cookies)
        self.__parse_available_semeseter(req.text)
        req.encoding = NckuGradeCrawler.ENCODING
        self.__parse_overall_summary(req.text)

    def __parse_available_semeseter(self, raw_html):
        parser = SemeseterNameParser()
        for line in raw_html.splitlines():
            parser.feed(line)
        self.semeseters = parser.get_semesters()

    def __parse_overall_summary(self, raw_html):
        parser = HTMLFormParser()
        for line in raw_html.splitlines():
            parser.feed(line)
        title = parser.get_tables()[-1][1][2:-2]
        content = parser.get_tables()[-1][-1][2:-2]
        self.overall_summary = OrderedDict(zip(title, content))

    def __parse_semester_data(self, semeseter_name):
        param = {'submit1': bytes(semeseter_name, 'cp1252')}
        req = self.session.post(NckuGradeCrawler.INDEX_URL,
                                params=param, data=self.data,
                                headers=NckuGradeCrawler.HEADER, cookies=self.session.cookies)
        req.encoding = NckuGradeCrawler.ENCODING

        parser = HTMLFormParser()
        for line in req.text.splitlines():
            parser.feed(line)
        data = parser.get_tables()[3]
        semester_data = {"courses": NckuGradeCrawler.__table_to_json(data[1:-2]),
                         "summary": NckuGradeCrawler.__split_summary(data[-1][0])}

        gpa = NckuGradeCrawler.__calculate_gpa(semester_data["courses"])
        semester_data["summary"]["GPA"] = gpa
        return semester_data

    @staticmethod
    def __table_to_json(table):
        table_json = list()
        for row in table[1:]:
            json_element = OrderedDict()
            for index, col in enumerate(row):
                title = table[0][index]
                json_element[title] = col
            table_json.append(json_element)
        return table_json

    @staticmethod
    def __split_summary(summary):
        expresion = r"(\D*):(\d*[.]?\d+)"
        matchs = re.findall(expresion, summary)

        summary_in_dict = OrderedDict()
        for match in matchs:
            summary_in_dict[match[0].strip()] = match[1].strip()
        return summary_in_dict

    @staticmethod
    def __calculate_gpa(courses):
        gpa, credits_sum = 0, 0

        course_credits, grades = [c["學分"] for c in courses], [c["分數"] for c in courses]
        for index, grade in enumerate(grades):
            if grade.isdecimal():
                credit = int(course_credits[index])
                credits_sum += credit

                grade = int(grade)
                if grade >= 80:
                    gpa += credit*4
                elif grade >= 70:
                    gpa += credit*3
                elif grade >= 60:
                    gpa += credit*2
                elif grade >= 50:
                    gpa += credit*1
        gpa = gpa/credits_sum
        return gpa

    def __overall_summerize(self):
        grade_sum, credits_sum, gpa_sum = 0, 0, 0
        general_course = dict()
        for sem_data in self.all_semester_data.values():
            summary = sem_data["summary"]
            grade_sum += int(summary["加權總分"])
            credits_sum += int(summary["總修學分"])
            gpa_sum += float(summary["GPA"]) * int(summary["總修學分"])

            courses = sem_data["courses"]
            for course in courses:
                if course[""]:
                    course_category = course[""]
                    if course_category not in general_course:
                        general_course[course_category] = list()
                    general_course[course_category].append(course["科目名稱"])

        extra_info = OrderedDict({"加權總分": grade_sum,
                                  "平均": grade_sum/credits_sum,
                                  "GPA": gpa_sum/credits_sum})
        self.overall_summary.update(extra_info)

        self.all_semester_data["Summary"] = self.overall_summary
        self.all_semester_data["Category"] = general_course

    def get_all_semester_data(self):
        return self.all_semester_data

    def export_as_xlsx(self, file_name="Grade Summary"):
        workbook = xlsxwriter.Workbook(file_name+".xlsx")
        for sheet_name, content in self.all_semester_data.items():
            worksheet = workbook.add_worksheet(sheet_name)
            if sheet_name not in ("Summary", "Category"):
                NckuGradeCrawler.__export_semestser_sheet(worksheet, content)
            elif sheet_name is "Summary":
                NckuGradeCrawler.__export_overall_summary_sheet(worksheet, content)
            elif sheet_name is "Category":
                NckuGradeCrawler.__export_category_sheet(worksheet, content)
        workbook.close()

    @staticmethod
    def __export_semestser_sheet(worksheet, content):
        table = NckuGradeCrawler.__json_to_table(content["courses"])
        for row_index, row in enumerate(table):
            for col_index, col in enumerate(row):
                worksheet.write(row_index, col_index, col)

        summary = content["summary"]
        course_num = len(table)
        for key, value in enumerate(list(summary.keys())):
            worksheet.write(course_num+1, key, value)
        for key, value in enumerate(list(summary.values())):
            worksheet.write(course_num+2, key, value)

    @staticmethod
    def __json_to_table(json_dict):
        table = list()

        table.append(list(json_dict[0].keys()))
        for data in json_dict:
            table.append(list(data.values()))
        return table

    @staticmethod
    def __export_overall_summary_sheet(worksheet, content):
        title = list(content.keys())
        summary = list(content.values())
        for key, value in enumerate(title):
            worksheet.write(0, key, value)
        for key, value in enumerate(summary):
            worksheet.write(1, key, value)

    @staticmethod
    def __export_category_sheet(worksheet, content):
        category = list(content.keys())

        for row_index, cate in enumerate(category):
            worksheet.write(row_index, 0, cate)
            worksheet.write(row_index, 1, len(content[cate]))
            for col_index, course in enumerate(content[cate]):
                worksheet.write(row_index, 2+col_index, course)


class SemeseterNameParser(HTMLParser):
    def __init__(self):
        HTMLParser.__init__(self)
        self.semeseters = list()

    def handle_starttag(self, tag, attrs):
        if tag == "input":
            self.semeseters.append(attrs[2][1])

    def get_semesters(self):
        return self.semeseters


if __name__ == '__main__':
    STU_ID = input("Please input student ID: ")
    PASSWD = getpass.getpass("Please input password: ")

    gradeCrawer = NckuGradeCrawler()
    gradeCrawer.set_stu_info(STU_ID, PASSWD)
    gradeCrawer.login()
    gradeCrawer.parse_all_semester_data()
    print("Export to xlsx")
    gradeCrawer.export_as_xlsx()
    gradeCrawer.logout()
