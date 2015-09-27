import re
import getpass
from html.parser import HTMLParser
from collections import OrderedDict

import requests
import xlsxwriter

from HTML_Form_Parser.HtmlFormParser import HtmlFormParser


class NckuGradeCrawler:
    MAIN_URL = "http://140.116.165.71:8888/ncku/"
    LOGIN_URL = MAIN_URL+"qrys02.asp"
    LOGOUT_URL = MAIN_URL+"logouts.asp"
    INDEX_URL = MAIN_URL+"/qrys05.asp"
    ENCODING = "big5"
    HEADER = {'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
              'X-Requested-With': 'XMLHttpRequest'}

    def __init__(self):
        self.s = requests.session()

    def login(self):
        self.s.post(NckuGradeCrawler.LOGIN_URL, data=self.data)

    def logout(self):
        self.s.post(NckuGradeCrawler.LOGOUT_URL)

    def set_stu_info(self, stu_id, passwd):
        self.data = {'ID': stu_id.upper(), 'PWD': passwd}

    def parse_all_semester_data(self):
        self.__parse_index_page()

        self.all_semester = OrderedDict()
        for s in self.semeseters:
            s_name = s[:4] + ("2" if "¤U" in s else "1")
            self.all_semester[s_name] = self.__parse_semester_data(s)
        self.__overall_summerize()

    def __parse_index_page(self):
        req = self.s.post(NckuGradeCrawler.INDEX_URL,
                          data=self.data,
                          cookies=self.s.cookies)
        fp, sp = HTMLFormParser(), SemeseterNameParser()
        for line in req.text.splitlines():
            fp.feed(line)
            sp.feed(line)
        self.semeseters = sp.get_semesters()
        column = ["必承", "選承", "暑承", "輔/雙", "抵承", "實得", "總修學分"]
        self.overall_summary = OrderedDict(zip(column, fp.get_tables()[-1][-1][2:-2]))

    def __parse_semester_data(self, semeseter_name):
        param = {'submit1': bytes(semeseter_name, 'cp1252')}
        req = self.s.post(NckuGradeCrawler.INDEX_URL,
                          params=param,
                          data=self.data,
                          headers=NckuGradeCrawler.HEADER,
                          cookies=self.s.cookies)
        req.encoding = NckuGradeCrawler.ENCODING

        p = HTMLFormParser()
        for line in req.text.splitlines():
            p.feed(line)
        data = p.get_tables()[3]
        semester_data = {"courses": self.__table_to_json(data[1:-2]),
                         "summary": self.__split_summary(data[-1][0])}

        gpa = self.__calculate_gpa(semester_data["courses"])
        semester_data["summary"]["GPA"] = gpa
        return semester_data

    def __table_to_json(self, table):
        table_json = list()
        for row in table[1:]:
            json_element = OrderedDict()
            for index, col in enumerate(row):
                title = table[0][index]
                json_element[title] = col
            table_json.append(json_element)
        return table_json

    def __split_summary(self, summary):
        expresion = "(\D*):(\d*[.]?\d+)"
        m = re.findall(expresion, summary)

        summary_in_dict = OrderedDict()
        for match in m:
            summary_in_dict[match[0].strip()] = match[1].strip()
        return summary_in_dict

    def __calculate_gpa(self, courses):
        gpa, credits_sum = 0, 0

        credits, grades = [c["學分"] for c in courses], [c["分數"] for c in courses]
        for index, grade in enumerate(grades):
            if grade.isdecimal():
                credit = int(credits[index])
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
        for key, value in self.all_semester.items():
            summary = value["summary"]
            credit = int(summary["總修學分"])
            grade_sum += int(summary["加權總分"])
            credits_sum += credit
            gpa_sum += float(summary["GPA"]) * credit

            courses = value["courses"]
            for c in courses:
                if c[""]:
                    course_category = c[""]
                    if course_category not in general_course:
                        general_course[course_category] = list()
                    general_course[course_category].append(c["科目名稱"])

        self.all_semester["Summary"] = self.overall_summary
        extra_info = OrderedDict({"加權總分": grade_sum,
                                  "平均": grade_sum/credits_sum,
                                  "GPA": gpa_sum/credits_sum})
        self.all_semester["Summary"].update(extra_info)

        self.all_semester["Category"] = general_course

    def get_all_semester_data(self):
        return self.all_semester

    def export_as_xlsx(self, file_name="Grade Summary"):
        workbook = xlsxwriter.Workbook(file_name+".xlsx")
        for sheet_name, content in self.all_semester.items():
            worksheet = workbook.add_worksheet(sheet_name)
            if sheet_name != "Summary" and sheet_name != "Category":
                table = self.__json_to_table(content["courses"])
                for row_index, row in enumerate(table):
                    for col_index, col in enumerate(row):
                        worksheet.write(row_index, col_index, col)

                summary = content["summary"]
                course_num = len(table)
                for key, value in enumerate(list(summary.keys())):
                    worksheet.write(course_num+1, key, value)
                for key, value in enumerate(list(summary.values())):
                    worksheet.write(course_num+2, key, value)
            elif sheet_name == "Summary":
                title = list(content.keys())
                summary = list(content.values())
                for key, value in enumerate(title):
                    worksheet.write(0, key, value)
                for key, value in enumerate(summary):
                    worksheet.write(1, key, value)
            elif sheet_name == "Category":
                category = list(content.keys())

                for index_r, cate in enumerate(category):
                    worksheet.write(index_r, 0, cate)
                    worksheet.write(index_r, 1, len(content[cate]))
                    for index_c, course in enumerate(content[cate]):
                        worksheet.write(index_r, 2+index_c, course)

        workbook.close()

    def __json_to_table(self, json_dict):
        table = list()

        table.append(list(json_dict[0].keys()))
        for data in json_dict:
            table.append(list(data.values()))
        return table


class SemeseterNameParser(HTMLParser):
    def __init__(self):
        HTMLParser.__init__(self)
        self.semeseters = list()

    def handle_starttag(self, tag, attrs):
        if tag == "input":
            self.semeseters.append(attrs[2][1])

    def get_semesters(self):
        return (self.semeseters)


if __name__ == '__main__':
    stu_id = input("Please input student ID: ")
    passwd = getpass.getpass("Please input password: ")

    g = NckuGradeCrawler()
    g.set_stu_info(stu_id, passwd)
    g.login()
    g.parse_all_semester_data()
    print("Export to xlsx")
    g.export_as_xlsx()
    g.logout()
