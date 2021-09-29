import threading
import tkinter
import tkinter as tk
from tkinter import *
from tkinter import scrolledtext
import time
from netsuite_clean_all_case import *
import netsuite_take_tasks as ntt
from outlook_send_emails import *
import _thread
import excel_test

LOG_LINE_NUM = 0
STATUS_LINE_NUM = 0


def get_current_time():
    current_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    return current_time


class Application(tk.Frame):

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        # self.pack()
        self.set_init_window()
        self.lock = threading.Lock
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
        self.result_data_Text = scrolledtext.ScrolledText(self.master, width=100, height=44, wrap=tk.WORD)  # 处理结果展示
        self.result_data_Text.grid(row=1, column=12, rowspan=15, columnspan=10)
        self.log_data_Text = scrolledtext.ScrolledText(self.master, width=65, height=9, wrap=tk.WORD)  # 日志框
        self.log_data_Text.grid(row=13, column=0, columnspan=10)

        # 按钮
        self.clean_button = Button(self.master, text="Clean All Noise Tasks", bg="WhiteSmoke", width=20,
                                   command=self.new_clean_thread)  # 调用内部方法  加()为直接调用
        self.clean_button.grid(row=2, column=11)
        self.take_button = Button(self.master, text="Take All PSA Tasks", bg="WhiteSmoke", width=20,
                                  command=self.new_take_thread)  # 调用内部方法  加()为直接调用
        self.take_button.grid(row=1, column=11)
        self.send_button = Button(self.master, text="Send an email to Amy", bg="WhiteSmoke", width=20,
                                  command=self.new_send_thread)  # 调用内部方法  加()为直接调用
        self.send_button.grid(row=4, column=11)
        self.resend_button = Button(self.master, text="Resend All Listed Tasks", bg="WhiteSmoke", width=20,
                                    command=self.do_resend)  # 调用内部方法  加()为直接调用
        self.resend_button.grid(row=3, column=11)
        self.update_button = Button(self.master, text="Update All Task Notes", bg="WhiteSmoke", width=20,
                                    command=self.new_update_thread)  # 调用内部方法  加()为直接调用
        self.update_button.grid(row=5, column=11)
        self.quit = Button(self.master, text="Quit", fg="red", bg="WhiteSmoke", width=20,
                           command=self.master.destroy)
        self.quit.grid(row=16, column=20)

    def write_log_to_text(self, logmsg):
        global LOG_LINE_NUM
        current_time = get_current_time()
        logmsg_in = str(current_time) + " " + str(logmsg) + "\n"  # 换行
        self.log_data_Text.insert(END, logmsg_in)

    def write_status_to_text(self, statusmsg):
        current_time = get_current_time()
        statusmsg_in = str(current_time) + " " + str(statusmsg) + "\n"  # 换行
        self.result_data_Text.insert(END, statusmsg_in)
    #
    # def create_widgets(self):
    #     self.master.title("Paul's Tool")  # 窗口名
    #     self.master.geometry('320x160+10+10')  # 290 160为窗口大小，+10 +10 定义窗口弹出时的默认展示位置
    #     self.master.geometry('1068x681+10+10')
    #     self.master["bg"] = "pink"
    #     self.clean_all_task = Button(self)
    #     self.clean_all_task["text"] = "Clean All Noise Task"
    #     self.clean_all_task["command"] = self.do_clean
    #     self.clean_all_task.pack(side="top")
    #     self.take_all_task = Button(self)
    #     self.take_all_task["text"] = "Take All PSA Task"
    #     self.take_all_task["command"] = self.do_take
    #     self.take_all_task.pack(anchor="center")
    #     self.quit = Button(self, text="QUIT", fg="red",
    #                        command=self.master.destroy)
    #     self.quit.pack(side="bottom")

    def do_update(self):
        self.write_status_to_text("Start updating all the PSA tasks:")
        self.write_status_to_text("--------------------------------------------------")
        robot = excel_test.TakeTasks()
        robot.update_all_notes()
        self.write_status_to_text("--------------------------------------------------")
        self.write_status_to_text("Complete updating all the PSA tasks!")
        self.write_log_to_text("Updated all the PSA tasks successfully !")

    def new_update_thread(self):
        update_thread = threading.Thread(target=self.do_update)
        update_thread.start()

    def new_clean_thread(self):
        take_thread = threading.Thread(target=self.do_clean)
        take_thread.start()

    def do_clean(self):

        self.write_status_to_text("Start cleaning all the noise cases:")
        self.write_status_to_text("--------------------------------------------------")
        robot = CleanAllCase()
        var = [
            'Bon Tool Company 997 To Amazon',
            "To Base Brands CC",
            "Bon Tool Company 997 To Amazon",
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
            "iTrade Network Unknown To Phillips Foods, Inc fka Phillips Seafood",
            "Amazon Unknown To Unknown",
            "Amazon.ca Unknown To Unknown",
            "Chewy.com Unknown To Unknown",
            "Digi-Key Corporation Unknown To Unknown",
            "Unknown Unknown To Bestseller",
            "Unknown Unknown To Abbyson Living Corporation",
            "CSN Unknown To Unknown",
            "Five Below 850 To Jem Accessories, Inc."
        ]
        for search_key in var:
            robot.change_criteria("contains", search_key)
            self.write_status_to_text("Closing "+str(robot.clean_all_case())+" cases with key word: "+search_key)
        robot.change_criteria("is not empty", "Hello")
        win32api.MessageBox(0, "No more noise in queue. :)", "Cleaning Done", win32con.MB_OK)
        self.write_status_to_text("--------------------------------------------------")
        self.write_status_to_text("Complete cleaning all the noise cases!")

        self.write_log_to_text("Cleaned all the noise cases successfully!")

    def do_take(self):
        self.write_status_to_text("--------------------------------------------------")
        self.write_status_to_text("Start taking all the PSA tasks:")
        robot = ntt.TakeTasks()
        robot.take_task()
        self.write_status_to_text("Complete taking all the PSA tasks!")
        self.write_status_to_text("--------------------------------------------------")
        self.write_log_to_text("Took all the PSA tasks successfully !")

    def new_take_thread(self):
        take_thread = threading.Thread(target=self.do_take)
        take_thread.start()

    def do_send(self):
        self.write_status_to_text("--------------------------------------------------")
        self.write_status_to_text("Start sending a log email to Amy:")
        robot = SendEmails()
        robot.send_amy_log()
        self.write_status_to_text("Complete sending a log email to Amy!")
        self.write_status_to_text("--------------------------------------------------")
        self.write_log_to_text("Sent a log email to Amy successfully !")

    def new_send_thread(self):
        take_thread = threading.Thread(target=self.do_send)
        take_thread.start()

    def do_resend(self):
        self.write_log_to_text("Resent all the transactions in list successfully !")

    def add_cloudftp(self):
        self.write_log_to_text("Setup a cloudftp for customer successfully !")



