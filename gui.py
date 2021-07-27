import tkinter
import tkinter as tk
from tkinter import *
import time
from netsuite_clean_all_case import *
from sub import *

LOG_LINE_NUM = 0


class Application(tk.Frame):

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        # self.pack()
        self.set_init_window()
        # self.create_widgets()

    def set_init_window(self):
        self.master.title("Paul's Tool")  # 窗口名
        # self.init_window_name.geometry('320x160+10+10')                         #290 160为窗口大小，+10 +10 定义窗口弹出时的默认展示位置
        self.master.geometry('1300x800+10+10')
        self.master["bg"] = "Gainsboro"  # 窗口背景色，其他背景色见：blog.csdn.net/chl0000/article/details/7657887
        # self.init_window_name.attributes("-alpha",0.9)                          #虚化，值越小虚化程度越高
        # 标签
        self.init_data_label = Label(self.master, text="Noise case key words")
        self.init_data_label.grid(row=0, column=0)
        self.result_data_label = Label(self.master, text="Script running status")
        self.result_data_label.grid(row=0, column=12)
        self.log_label = Label(self.master, text="Mission log")
        self.log_label.grid(row=12, column=0)
        # 文本框
        self.init_data_Text = Text(self.master, width=65, height=35)  # 原始数据录入框
        self.init_data_Text.grid(row=1, column=0, rowspan=10, columnspan=10)
        self.result_data_Text = Text(self.master, width=70, height=44)  # 处理结果展示
        self.result_data_Text.grid(row=1, column=12, rowspan=15, columnspan=10)
        self.log_data_Text = Text(self.master, width=65, height=9)  # 日志框
        self.log_data_Text.grid(row=13, column=0, columnspan=10)

        # 按钮
        self.clean_button = Button(self.master, text="Clean All Noise Tasks", bg="WhiteSmoke", width=20,
                                   command=self.do_clean)  # 调用内部方法  加()为直接调用
        self.clean_button.grid(row=2, column=11)
        self.take_button = Button(self.master, text="Take All PSA Tasks", bg="WhiteSmoke", width=20,
                                  command=self.do_take)  # 调用内部方法  加()为直接调用
        self.take_button.grid(row=1, column=11)
        self.send_button = Button(self.master, text="Send an email to Amy", bg="WhiteSmoke", width=20,
                                  command=self.do_send)  # 调用内部方法  加()为直接调用
        self.send_button.grid(row=4, column=11)
        self.resend_button = Button(self.master, text="Resend All Listed Tasks", bg="WhiteSmoke", width=20,
                                    command=self.do_resend)  # 调用内部方法  加()为直接调用
        self.resend_button.grid(row=3, column=11)
        self.quit = Button(self.master, text="Quit", fg="red", bg="WhiteSmoke", width=20,
                           command=self.master.destroy)
        self.quit.grid(row=16, column=20)

    def write_log_to_text(self, logmsg):
        global LOG_LINE_NUM
        current_time = self.get_current_time()
        logmsg_in = str(current_time) + " " + str(logmsg) + "\n"  # 换行
        if LOG_LINE_NUM <= 7:
            self.log_data_Text.insert(END, logmsg_in)
            LOG_LINE_NUM = LOG_LINE_NUM + 1
        else:
            self.log_data_Text.delete(1.0, 2.0)
            self.log_data_Text.insert(END, logmsg_in)

    def create_widgets(self):
        self.master.title("Paul's Tool")  # 窗口名
        self.master.geometry('320x160+10+10')  # 290 160为窗口大小，+10 +10 定义窗口弹出时的默认展示位置
        self.master.geometry('1068x681+10+10')
        self.master["bg"] = "pink"
        self.clean_all_task = Button(self)
        self.clean_all_task["text"] = "Clean All Noise Task"
        self.clean_all_task["command"] = self.do_clean
        self.clean_all_task.pack(side="top")
        self.take_all_task = Button(self)
        self.take_all_task["text"] = "Take All PSA Task"
        self.take_all_task["command"] = self.do_take
        self.take_all_task.pack(anchor="center")
        self.quit = Button(self, text="QUIT", fg="red",
                           command=self.master.destroy)
        self.quit.pack(side="bottom")

    def do_clean(self):
        robot = CleanAllCase()
        var = [
            "To Base Brands CC",
            "Amware Logistics Unknown To Unknown",
            "Almo Unknown To Unknown",
            "Home Depot Canada Unknown To Unknown",
            "To Nurse Assist, Inc.",
            "Medline Unknown To Unknown",
            "P2P - Cat5 Commerce Unknown To Unknown",
            "Tractor Supply Drop Ship Unknown To Unknown",
            "Unknown Unknown To Unknown",
            "Walmart Unknown To Unknown",
            "Kroger Unknown To Unknown",
            "TM File processing",
            # "iTrade Network Unknown To Phillips Foods, Inc",
            "Unknown Unknown To Total Quality Logistics 2",
            "Amazon Unknown To Unknown",
            "Amazon.ca Unknown To Unknown",
            "Chewy.com Unknown To Unknown",
            "Digi-Key Corporation Unknown To Unknown"
        ]
        for search_key in var:
            robot.change_criteria("contains", search_key)
            robot.clean_all_case()
        robot.change_criteria("is not empty", "Hello")
        win32api.MessageBox(0, "No more noise in queue. :)", "Cleaning Done", win32con.MB_OK)
        self.write_log_to_text("Cleaned all the message successfully !")

    def do_take(self):
        robot = CleanAllCase()
        robot.take_task()
        self.write_log_to_text("Took all the PSA tasks successfully !")

    def do_send(self):
        SendEmails.send_amy_log()
        self.write_log_to_text("Sent a log email to Amy successfully !")

    def do_resend(self):
        self.write_log_to_text("Resent all the transactions in list successfully !")

    def add_cloudftp(self):
        self.write_log_to_text("Setup a cloudftp for customer successfully !")

    def get_current_time(self):
        current_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        return current_time
