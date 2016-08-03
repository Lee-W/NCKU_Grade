import re
import json
import getpass
from collections import OrderedDict

import requests
import xlsxwriter
from bs4 import BeautifulSoup


class NckuGradeCrawler:
    MAIN_URL = "http://140.116.165.71:8888/ncku/"
    LOGIN_URL = MAIN_URL+"qrys02.asp"
    LOGOUT_URL = MAIN_URL+"logouts.asp"
    INDEX_URL = MAIN_URL+"qrys05.asp"
    ENCODING = "big5"
    HEADER = {'Content-Type': ('application/x-www-form-urlencoded;'
                               'charset=UTF-8'),
              'X-Requested-With': 'XMLHttpRequest'}

    def __init__(self):
        self._session = requests.session()
        self._stu_info = dict()
        self._rule_path = "rule/origin_rule.json"
        self.semesters = list()
        self.overall_summary = OrderedDict()
        self.all_semester_data = OrderedDict()

    @property
    def stu_info(self):
        return self._stu_info

    def set_stu_info(self, stu_id, passwd):
        self._stu_info = {'ID': stu_id.upper(), 'PWD': passwd}

    @property
    def rule_path(self):
        return self._rule_path

    @rule_path.setter
    def rule_path(self, rf):
        self._rule_path = rf

    def login(self):
        self._session.post(NckuGradeCrawler.LOGIN_URL, data=self._stu_info)

    def logout(self):
        self._session.post(NckuGradeCrawler.LOGOUT_URL)

    def parse_all_semester_data(self):
        self.__load_gpa_rule()
        self.__parse_index_page()

        self.all_semester_data = OrderedDict()
        for sem in self.semesters:
            s_name = sem[:4] + ("2" if "下" in sem else "1")
            self.all_semester_data[s_name] = self.__parse_semester_data(sem)
        self.__overall_summerize()

    def __load_gpa_rule(self):
        with open(self.rule_path) as rule_file:
            self.rule = json.load(rule_file, object_pairs_hook=OrderedDict)

    def __parse_index_page(self):
        req = self._session.post(NckuGradeCrawler.INDEX_URL,
                                 data=self._stu_info,
                                 cookies=self._session.cookies)
        req.encoding = NckuGradeCrawler.ENCODING
        soup = BeautifulSoup(req.text, "html5lib")
        self.semesters = [tag['value'] for tag in soup.find_all('input')]
        self.__parse_overall_summary(soup)

    def __parse_overall_summary(self, soup):
        table = soup.find_all('table')[-1]
        title = [td.text.strip()
                 for td in table.find_all('tr')[1].find_all('td')[2:-1]]
        content = [td.text.strip()
                   for td in table.find_all('tr')[-1].find_all('td')[2:-1]]
        self.overall_summary = OrderedDict(zip(title, content))

    def __parse_semester_data(self, semester_name):
        param = {'submit1': (semester_name[:-1] +
                             "¤" +
                             ('W' if semester_name[-1] == '上'
                              else 'U')).encode('cp1252')}
        req = self._session.post(NckuGradeCrawler.INDEX_URL,
                                 params=param, data=self._stu_info,
                                 headers=NckuGradeCrawler.HEADER,
                                 cookies=self._session.cookies)
        req.encoding = NckuGradeCrawler.ENCODING
        soup = BeautifulSoup(req.text, 'html5lib')
        table = soup.find_all('table')[3]
        data = list()
        for tr in table.find_all('tr'):
            row_data = list()
            for td in tr.find_all('td'):
                row_data.append(td.text.strip())
            data.append(row_data)

        semester_data = {"courses": NckuGradeCrawler.__table_to_json(data[1:-2]),
                         "summary": NckuGradeCrawler.__split_summary(data[-1][0])}
        gpa = self.__calculate_gpa(semester_data["courses"])
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

    def __calculate_gpa(self, courses):
        gpa, credits_sum = 0, 0
        course_credits, grades = [c["學分"] for c in courses], [c["分數"][:-2] for c in courses]
        for index, grade in enumerate(grades):
            if grade.isdecimal():
                credit = int(course_credits[index])
                credits_sum += credit

                grade = int(grade)
                for threshold, point in self.rule.items():
                    if grade >= int(threshold):
                        gpa += credit*float(point)
                        break
        gpa = gpa/credits_sum
        return gpa

    def __overall_summerize(self):
        grade_sum, credits_sum, gpa_sum = 0, 0, 0
        general_course = dict()
        for sem_data in self.all_semester_data.values():
            summary = sem_data["summary"]
            grade_sum += int(summary["加權總分"])
            current_credicts = 0

            courses = sem_data["courses"]
            for course in courses:
                if course[""]:
                    course_category = course[""]
                    if course_category not in general_course:
                        general_course[course_category] = list()
                    general_course[course_category].append(course["科目名稱"])
                # skip courses without grade
                if course['分數'][:-2].isdecimal():
                    current_credicts += int(course['學分'])

            credits_sum += current_credicts
            gpa_sum += float(summary["GPA"]) * current_credicts

        extra_info = OrderedDict({"加權總分": grade_sum,
                                  "平均": grade_sum/credits_sum,
                                  "GPA": gpa_sum/credits_sum})
        self.overall_summary.update(extra_info)

        self.all_semester_data["Summary"] = self.overall_summary
        self.all_semester_data["Category"] = general_course

    def export_as_xlsx(self, file_name="Grade Summary"):
        workbook = xlsxwriter.Workbook(file_name+".xlsx")
        for sheet_name, content in self.all_semester_data.items():
            worksheet = workbook.add_worksheet(sheet_name)
            if sheet_name not in ("Summary", "Category"):
                NckuGradeCrawler.__export_semestser_sheet(worksheet, content)
            elif sheet_name is "Summary":
                NckuGradeCrawler.__export_overall_summary_sheet(worksheet,
                                                                content)
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


if __name__ == '__main__':
    STU_ID = input("Please input student ID: ")
    PASSWD = getpass.getpass("Please input password: ")
    GPA_RULE = input(("Choose GPA rule\n"
                      "1. Origin Rule (Before 104)\n"
                      "2. New Rule (After 104)\n"
                      "Please Enter[1 or 2]: \n"
                      ))
    gradeCrawer = NckuGradeCrawler()
    gradeCrawer.set_stu_info(STU_ID, PASSWD)
    gradeCrawer.rule_path = 'rule/{}_rule.json'.format(
        "new" if GPA_RULE == '2' else "origin"
    )
    gradeCrawer.login()
    gradeCrawer.parse_all_semester_data()
    print("Export to xlsx")
    gradeCrawer.export_as_xlsx()
    gradeCrawer.logout()
