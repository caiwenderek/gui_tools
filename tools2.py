#!/usr/bin/python3
# -*-coding:utf-8 -*-

# Reference:**********************************************
# @Time    : 2022/10/18 10:21
# @Author  : wenbin.cai
# @File    : tools.py
# @Software: PyCharm
# Reference:**********************************************
import operator
import traceback
from functools import partial
from tkinter import Tk, filedialog, HORIZONTAL, messagebox, END, VERTICAL, RIGHT, INSERT, StringVar
import re
import unicodedata

from tkinter.scrolledtext import ScrolledText
from tkinter.ttk import Label, Entry, Button, Progressbar, Scrollbar, Treeview, Frame
from tkinter import Label as TLabel, LEFT
import textwrap

import tkinterDnD
import webbrowser

FILE_TYPES = [
            ('Text Files', '*.txt'),
            ('doc ext Files', '*.docx'),
            ('doc files', '*.doc')]  # type of files to select

"""
2。doc 解析
3。更新frame 数据
4。拉伸窗口
5. remove 处理

"""


class WinGUI(tkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.__win()
        self.file_label = self.__file_label()
        self.file = self.file_path()
        self.upload_bt = self.upload_button()

        # self.process_bt = self.process_button()
        self.processing_label = self.__processing_label()
        self.progressbar = self.__progressbar()

        self.tableview = self.create_table()
        self.total, self.calculate = self.create_summation()

    def __win(self):
        self.title("造价统计小工具")
        # 设置窗口大小、居中
        width = 600
        height = 500
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()
        print(screenwidth, screenheight)
        geometry = '%dx%d+%d+%d' % (width+120, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        # self.geometry(geometry)
        self.geometry('768x700')
        # self.resizable(width=False, height=False)

        # 设置高度
        self.grid_rowconfigure(0, minsize=30)
        self.grid_rowconfigure(1, minsize=30)
        self.grid_rowconfigure(2, weight=1, minsize=400)
        self.grid_rowconfigure(3, minsize=120)
        self.grid_rowconfigure(4, minsize=30)

        # 设置宽度
        self.grid_columnconfigure(0, minsize=40)
        self.grid_columnconfigure(1, weight=1, minsize=150)
        self.grid_columnconfigure(2, minsize=40)
        self.grid_columnconfigure(3, minsize=40)

        # y scrollbar
        self.grid_columnconfigure(4, minsize=20)

    def create_table(self):
        # 表头字段 表头宽度
        columns = {"原文": (500, True), "造价": (90, False)}
        # columns = {"原文": 680, "造价": 90}
        # 初始化表格 表格是基于Treeview，tkinter本身没有表格。show="headings" 为隐藏首列。
        # tk_table = Treeview(self, show="headings", columns=list(columns))

        tk_table = Treeview(
                master=self,  # 父容器
                columns=list(columns),  # 列标识符列表
                # height=30,  # 表格显示的行数
                # show='tree headings',  # 隐藏首列
                show='headings',  # 隐藏首列
                style='Treeview',  # 样式
                # xscrollcommand=xscroll.set,  # x轴滚动条
                # yscrollcommand=yscroll.set  # y轴滚动条
        )
        tk_table.tag_configure('odd', background='#E8E8E8')
        tk_table.tag_configure('even', background='#DFDFDF')

        # style = Style(self)
        # # style.configure('Treeview', rowheight=30)
        # # set ttk theme to "clam" which support the fieldbackground option
        # # style.theme_use("clam")
        # # style.theme_use("default")
        # style.configure("Treeview",
        #                 # padding=1,
        #                 borderwidth=0,
        #                 highlightbackground="red"
        #                 # background="black",
        #                 # fieldbackground="black", foreground="white"
        #                 )

        # 设置滚动条
        # xscroll = Scrollbar(self, orient=HORIZONTAL, command=tk_table.xview)
        yscroll = Scrollbar(self, orient=VERTICAL, command=tk_table.yview)
        tk_table.configure(yscrollcommand=yscroll.set)
        yscroll.grid(row=2, column=4, sticky='nse')

        tk_table.configure(ondrop=partial(self.drop, self.file), ondragstart=self.drag_command,)

        for text, (width, stretch) in columns.items():  # 批量设置列属性
            tk_table.heading(text, text=text, anchor='center')
            tk_table.column(text, anchor='e', width=width, stretch=stretch)
            # stretch 不自动拉伸
        tk_table.grid(row=2, column=0, padx=(10, 0), columnspan=4, sticky="news")
        return tk_table

    def create_summation(self):
        # 添加计算公式
        # calculate_frame.grid(row=3, column=0, padx=(5, 0), columnspan=5, sticky="news",
        #                      pady=(5, 0))

        calculate_text = ScrolledText(self, height=5, width=110, wrap='word')

        calculate_text.grid(row=3, column=0, padx=(5, 0), columnspan=5, sticky="news",
                             pady=(5, 0))

        # 添加合计
        total_label = Label(self, text='合计：', anchor='e',  justify=RIGHT)
        total_label.grid(row=4, column=0)

        total_var = StringVar(self, '0.0')
        total_ent = Entry(self, textvariable=total_var)
        total_ent.grid(row=4, column=1, sticky='w')
        total_ent.configure(state='readonly')

        help_add = TLabel(self, text='校验合计网站', fg="blue", cursor="hand2")
        help_add.bind("<Button-1>", lambda e: webbrowser.open_new('https://www.wncx.cn/gongshi/'))
        # help_add =  Label(self, text='合计=================================================================：', anchor='w',  justify=LEFT)
        help_add.grid(row=4, column=2)

        Button(self, text='退出', command=self.quit).grid(row=4, column=3, columnspan=2)

        return total_var, calculate_text

    def __file_label(self):
        label = Label(self, text="word或文本文件：")
        label.grid(row=0, column=0, padx=5)

        # label.place(x=5, y=5, width=120, height=30)
        return label

    def file_path(self):
        def check():
            return text in set(map(operator.itemgetter(1), FILE_TYPES))

        text = Entry(self,
                     validate="focusout",
                     validatecommand=check,
                     width=400
                     )
        # text.place(x=140, y=5, width=500, height=30)
        text.grid(row=0, column=1, columnspan=2)
        text.configure(ondrop=partial(self.drop, text), ondragstart=self.drag_command,)

        return text

    def upload_button(self):
        def upload_file():
            filename = filedialog.askopenfilename(filetypes=FILE_TYPES)
            self.file.delete(0, END)
            self.file.insert(0, filename)
            self.processing()

        btn = Button(self, text="上传", command=upload_file)
        btn.grid(row=0, column=3, padx=5, pady=2)
        return btn

    # def process_button(self):
    #     btn = Button(self, text="处理", command=self.processing)
    #     btn.grid(row=0, column=3, padx=5)
    #     return btn

    def processing(self):

        t = Tools(self.file.get())
        try:
            data, calculate_str, sum_str = t.run()
            self.render_table(data)
            self.add_total_line(sum_str, calculate_str)
        except Exception as e:
            error_info = traceback.format_exc()
            print(error_info)
            messagebox.showinfo(error_info)
            self.file.delete(0, END)
            return None

    def clear_tables(self):
        self.listBox.delete(*self.listBox.get_children())

    def __processing_label(self):
        label = Label(self, text="处理进度")
        label.grid(row=1, column=0)
        return label

    def __progressbar(self):
        progressbar = Progressbar(self, orient=HORIZONTAL)
        # progressbar.grid(row=1, column=1, columnspan=3, sticky='news')
        # progressbar.place(x=80, y=40, width=710, height=30)
        return progressbar

    def render_table(self, data):
        def wrap(string, lenght=30):
            return '\n'.join(textwrap.wrap(string, lenght))

        self.tableview.delete(*self.tableview.get_children())  # 清空原先表格
        # 导入初始数据
        for i, (content, price) in enumerate(data):
            content = content[-55:]
            self.tableview.insert('', END, values=(content, price), tags=('odd',) if i % 2 == 0 else ('even',))

        return self.tableview

    def add_total_line(self, total, calculate_str):
        # self.total.delete(0, 'end')
        self.total.set(total)
        self.calculate.delete('1.0', 'end')
        self.calculate.insert(INSERT, calculate_str)


class Win(WinGUI, Frame):
    def __init__(self):
        super().__init__()
        self.__event_bind()

    def __event_bind(self):
        # self.configure(ondrop=partial(self.drop, self.file),
        #                ondragstart=self.drag_command,)

        self.register_drop_target("*")
        self.bind("<<Drop>>", partial(self.drop, self.file))

        self.register_drag_source("*")
        self.bind("<<DragInitCmd>>", self.drag_command)

    def drop(self, text, event):
        # This function is called, when stuff is dropped into a widget
        text.delete(0, 'end')
        text.insert(0, event.data)
        self.processing()
        # stringvar.set(event.data)

    def drag_command(self, event):
        # This function is called at the start of the drag,
        # it returns the drag type, the content type, and the actual content
        return (tkinterDnD.COPY, "DND_Text", "Some nice dropped text!")


class Tools:
    plus_regex = re.compile(r"核增造?价?(\d+(\.\d+)?万?)元?[.,;。，；]")
    minus_regex = re.compile(r"核减造?价?(\d+(\.\d+)?万?)元?[.,;。，；]")

    def __init__(self, file_path):
        self.file = file_path

    def search_num(self, txt):
        if not txt:
            return
        matcher = self.plus_regex.search(txt)
        if matcher:
            g = matcher.group(1)
            return self.process_num(g)
        matcher = self.minus_regex.search(txt)
        if matcher:
            g = matcher.group(1)
            return -self.process_num(g)

    @staticmethod
    def process_num(txt):
        append = 1
        if txt.endswith('万'):
            txt = txt[:-1]
            append = 10000
        return float(txt) * append

    @staticmethod
    def wide_chars(s):
        return sum(unicodedata.east_asian_width(x) == 'W' for x in s)

    def width(self, s):
        return len(s) + self.wide_chars(s)

    def run(self):
        prices = []
        data = []
        print('详细识别造价如下'.center(80, '='))
        for line in open(self.file, encoding='utf-8'):
            line = line.strip()
            if not line:
                continue

            price = self.search_num(line)
            price = round(price, 2) if price else 0

            # printf(line, single)
            data.append((line, price))
            if not price:
                continue
            prices.append(price)

        c = [f"+{str(a)}" if a > 0 else str(a) for a in prices]
        if c and c[0][0] == '+':
            c[0] = c[0][1:]

        calculate_str = f"计算公式：{''.join(c)}"
        sum_str = f"{sum(prices)}"

        print(calculate_str)
        print(sum_str)
        return data, calculate_str, sum_str


if __name__ == "__main__":
    win = Win()
    win.mainloop()
