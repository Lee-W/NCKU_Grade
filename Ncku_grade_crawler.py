import re
import getpass
import json
from html.parser import HTMLParser
from collections import OrderedDict

import requests
from HTML_Form_Parser.HTML_form_parser import HTMLFormParser


class NckuGradeCrawler:
    MAIN_URL = "http://140.116.165.71:8888/ncku/"
    LOGIN_URL = MAIN_URL+"qrys02.asp"
    LOGOUT_URL = MAIN_URL+"logouts.asp"
    INDEX_URL = MAIN_URL+"/qrys05.asp"
    ENCODING = "big5"
    header = {'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
              'X-Requested-With': 'XMLHttpRequest'}

    def __init__(self):
        self.s = requests.session()

    def login(self):
        self.s.post(NckuGradeCrawler.LOGIN_URL, data=self.data)

    def logout(self):
        self.s.post(NckuGradeCrawler.LOGOUT_URL)

    def set_stu_info(self, stu_id, passwd):
        self.data = {'ID': stu_id.upper(), 'PWD': passwd}

    def get_all_semester_data(self):
        self.all_semester = OrderedDict()
        sems = self.get_available_semester_name()
        for s in sems:
            s_name = s[:4] + ("2" if "¤U" in s else "1")
            self.all_semester[s_name] = self.__parse_semester_data(s)
        self.__overall_summerize()
        return self.all_semester

    def get_available_semester_name(self):
        req = self.s.post(NckuGradeCrawler.INDEX_URL,
                          data=self.data,
                          cookies=self.s.cookies)

        p = SemeseterNameParser()
        for line in req.text.splitlines():
            p.feed(line)
        return p.get_semesters()

    def __parse_semester_data(self, semeseter_name):
        param = {'submit1': bytes(semeseter_name, 'cp1252')}
        req = self.s.post(NckuGradeCrawler.INDEX_URL,
                          params=param,
                          data=self.data,
                          headers=NckuGradeCrawler.header,
                          cookies=self.s.cookies)
        req.encoding = NckuGradeCrawler.ENCODING

        p = HTMLFormParser()
        for line in req.text.splitlines():
            p.feed(line)
        data = p.get_tables()[3]
        semester_data = {"grades": self.__table_to_json(data[1:-2]),
                         "summary": self.__split_summary(data[-1][0])}
        return semester_data

    def __split_summary(self, summary):
        expresion = "(\D*):(\d*[.]?\d+)"
        m = re.findall(expresion, summary)

        summary_in_dict = OrderedDict()
        for match in m:
            summary_in_dict[match[0].strip()] = match[1].strip()
        return summary_in_dict

    def __overall_summerize(self):
        # TODO: implement
        pass

    def __calculate_gpa(self):
        # TODO: implement
        pass

    def __table_to_json(self, table):
        table_json = list()
        for row in table[1:]:
            json_element = OrderedDict()
            for index, col in enumerate(row):
                title = table[0][index]
                json_element[title] = col
            table_json.append(json_element)
        return table_json


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
    data = g.get_all_semester_data()
    print(json.dumps(data, indent=4,  ensure_ascii=False))

    g.logout()
