import logging
from tkinter import Tk
from tkinter import Frame
from tkinter import Label
from tkinter import Entry
from tkinter import Button
from tkinter import Radiobutton
from tkinter import IntVar
from tkinter.filedialog import asksaveasfilename

from NckuGradeCrawler import NckuGradeCrawler


class GradeGUI(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master = master
        self.grid(columnspan=2000)

        self.rule_choosen = IntVar()
        self.rule_choosen.set(1)
        self.__create_widgets()
        self._grade_crawler = NckuGradeCrawler()

    def __create_widgets(self):
        self.stu_no_label = Label(self, text="學號 ： ")
        self.stu_no_field = Entry(self, width=10)
        self.passwd_label = Label(self, text="密碼 ： ")
        self.passwd_field = Entry(self, width=25)
        self.origin_rule_radio = Radiobutton(self,
                                             text="原GPA計算方式(104年以前)",
                                             variable=self.rule_choosen,
                                             value=1)
        self.new_rule_radio = Radiobutton(self,
                                          text="新GPA計算方式(104年以後)",
                                          variable=self.rule_choosen,
                                          value=2)
        self.export_btn = Button(self, text="輸出", width=20,
                                 command=self.__export_method)
        self.success_label = Label(self, text="輸出成功")

        self.passwd_field.config(show="*")

        self.stu_no_label.grid(row=0, column=0)
        self.stu_no_field.grid(row=0, column=1, columnspan=25, sticky="W")
        self.passwd_label.grid(row=1, column=0)
        self.passwd_field.grid(row=1, column=1, columnspan=25)
        self.origin_rule_radio.grid(row=2, column=0)
        self.new_rule_radio.grid(row=2, column=1)
        self.export_btn.grid(row=3, column=0)

    def __choose_export_path(self):
        default_file_name = "Grade Summary"
        filetypes = [("Microsoft Excel 2007/2010/2013 XML", "*.xlsx")]
        title = "Save the file as..."
        self.file_name = asksaveasfilename(parent=self.master,
                                           initialfile=default_file_name,
                                           filetypes=filetypes,
                                           title=title)

    def __export_method(self):
        stu_no = self.stu_no_field.get()
        passwd = self.passwd_field.get()

        self.__choose_export_path()
        if self.file_name:
            try:
                self._grade_crawler.set_stu_info(stu_no, passwd)
                self._grade_crawler.rule_path = (
                    'rule/{}_rule.json'.format(
                        'new' if self.rule_choosen.get() == 2 else 'origin'
                    )
                )
                logging.debug("This is choosen "+str(self.rule_choosen.get()))
                self._grade_crawler.login()
                self._grade_crawler.parse_all_semester_data()
                self._grade_crawler.export_as_xlsx(self.file_name)
                self._grade_crawler.logout()
            except Exception as e:
                self.success_label['text'] = "未知的錯誤"
                logging.debug(str(e))
            self.success_label.grid(row=3, column=1)


if __name__ == '__main__':
    logging.basicConfig(level=logging.DEBUG)

    root = Tk()
    root.title("NCKU Grade")
    app = GradeGUI(master=root)
    app.mainloop()
