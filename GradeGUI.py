from tkinter import Tk
from tkinter import Frame
from tkinter import Label
from tkinter import Entry
from tkinter import Button
from tkinter.filedialog import asksaveasfilename

from NckuGradeCrawler import NckuGradeCrawler


class GradeGUI(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master = master
        self.grid(columnspan=2000)

        self.__create_widgets()

    def __create_widgets(self):
        self.stu_no_label = Label(self, text="學號 ： ")
        self.stu_no_field = Entry(self, width=10)
        self.passwd_label = Label(self, text="密碼 ： ")
        self.passwd_field = Entry(self, width=25)
        self.export_btn = Button(self, text="輸出", command=self.__export_method)
        self.success_label = Label(self, text="輸出成功")

        self.passwd_field.config(show="*")

        self.stu_no_label.grid(row=0, column=0)
        self.stu_no_field.grid(row=0, column=1, columnspan=25, sticky="W")
        self.passwd_label.grid(row=1, column=0)
        self.passwd_field.grid(row=1, column=1, columnspan=25)
        self.export_btn.grid(row=2, column=0)

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
            g = NckuGradeCrawler()
            g.set_stu_info(stu_no, passwd)
            g.login()
            g.parse_all_semester_data()
            g.export_as_xlsx(self.file_name)
            g.logout()
            self.success_label.grid(row=2, column=1)


if __name__ == '__main__':
    g = NckuGradeCrawler()
    root = Tk()
    root.title("NCKU Grade")
    app = GradeGUI(master=root)
    app.mainloop()
