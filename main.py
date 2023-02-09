import tkinter as tk
from tkinter import filedialog as fd
from tkinter import ttk
import executor as parser
from openpyxl.utils.exceptions import InvalidFileException


class Gui:
    def __init__(self):
        self.win = tk.Tk()
        # icon = tk.PhotoImage(file='lepr.ico')
        # self.win.iconphoto(False, icon)
        # self.win.iconbitmap('./123')
        self.win.title('Почти год бессонных ночей учебы, и я могу это :D')
        self.win.geometry('700x200+100+200')
        self.win.config(bg='#EBD4B5')
        self.win.resizable(True, False)
        self.win.grid_columnconfigure(0, minsize=200)
        self.win.grid_columnconfigure(1, minsize=400)

        tk.Button(text='Click to Open File', command=self.callback).grid(row=0, column=0, sticky='w')
        self.widget_path = tk.Label(self.win)
        self.widget_path.grid(row=0, column=1, sticky='we')

        tk.Label(self.win, text="Начало диапазона в формате: H:M").grid(row=1, column=0, sticky='w')
        self.time_start = tk.Entry(self.win)
        self.time_start.grid(row=1, column=1)
        self.time_start.insert(0, '00:00')

        tk.Label(self.win, text="Конец диапазона в формате: H:M").grid(row=2, column=0, sticky='w')
        self.time_stop = tk.Entry(self.win)
        self.time_stop.grid(row=2, column=1)
        self.time_stop.insert(0, '00:00')

        tk.Label(self.win, text="Выбрать Т.Т.").grid(row=3, column=0, sticky='w')

        self.lst_combobox = []
        self.combobox = ttk.Combobox(values=self.lst_combobox)
        self.combobox.grid(row=3, column=1, sticky='we')

        tk.Button(self.win, text='Вычеслить сумму чеков', command=self.calculate).grid(row=4, column=0, sticky='w')
        self.widget_answer = tk.Entry(self.win)
        self.widget_answer.grid(row=4, column=1, sticky='we')
        tk.Button(self.win, text='Записать в файл', command=self.save).grid(row=5, column=0, sticky='w')

        self.obj_file = None

    def save(self):
        if self.__set_time():
            self.obj_file.create_report(market=self.combobox.get() or 'ALL')

    def callback(self):
        self.widget_path.config(text='')
        self.widget_answer.delete(0, "end")
        path = fd.askopenfilename()
        self.widget_path.config(text=path)
        self.__create_obj_file(path)
        self.__set_lst_combobox()

    def __create_obj_file(self, path):
        try:
            self.obj_file = None
            self.obj_file = parser.Parser_xls(path)
        except InvalidFileException:
            self.widget_answer.insert(0, 'ERROR! Допустимые форматы: .xlsx,.xlsm,.xltx,.xltm')

    def __set_lst_combobox(self):
        if self.obj_file:
            self.combobox.config(values=self.obj_file.get_all_market())
        else:
            self.combobox.config(values=[])

    def calculate(self):
        self.widget_answer.delete(0, "end")
        if self.__set_time():
            market = self.combobox.get() or "ALL"
            result = self.obj_file.start_parse(market)
            self.widget_answer.insert(0, string=f'Сумма: {result[0]}   Kоличество чеков: {result[1]}')

    def __set_time(self):
        try:
            if self.obj_file:
                self.obj_file.set_time(start_time=self.time_start.get(), stop_time=self.time_stop.get())
        except ValueError:
            self.widget_answer.insert(0, 'Неверный формат времени')
            return False
        else: 
            return True      

    def run(self):
        self.win.mainloop()


if __name__ == "__main__":
    gui = Gui()
    gui.run()






