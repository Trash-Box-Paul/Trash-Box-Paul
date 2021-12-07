import threading
import win32gui
from netsuite_take_tasks import TakeTasks
import win32api, win32con
import tkinter as tk
from tkinter import ttk
from tkinter import scrolledtext
from tkinter import Menu
from tkinter import messagebox as msg
from netsuite_clean_all_case import *
from tkinter import Spinbox
from outlook_send_emails import *
import time
from time import sleep
import GUI.ToolTip as tt
from tkinter import *
from threading import Thread
from queue import Queue
import excel_test
import GUI.Queues as bq
import os
from tkinter import filedialog as fd
from os import path, makedirs
from datetime import datetime

# Module level GLOBALS
GLOBAL_CONST = 42

# Create Backup Folder
fDir = path.dirname(path.dirname(__file__))
netDir = fDir + '\\Backup'
if not path.exists(netDir):
    makedirs(netDir, exist_ok=True)


def get_current_time():
    current_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    return current_time


# =====================================================


class OOP():
    def __init__(self):  # Initializer method
        # Create instance
        self.clean_robot = CleanAllCase()
        self.excel_robot = excel_test.TakeTasks()
        self.root_browser = self.clean_robot.driver
        self.win = tk.Tk()
        # Single thread
        self.run_thread = None
        # Add a title       
        self.win.title("TPS Automation Tool 0.0.1")
        # Create a Queue
        self.gui_queue = Queue()
        self.create_widgets()
        # Initial the address of File Entries
        self.defaultFileEntries()

    def defaultFileEntries(self):
        self.fileEntry.delete(0, tk.END)
        self.fileEntry.insert(0, fDir + '\\Latest_Spread_Sheet.xlsm')
        self.fileEntry.config(state='readonly')
        if len(fDir) > self.entryLen:
            self.fileEntry.config(width=self.entryLen)  # limit width to adjust GUI
        self.netwEntry.delete(0, tk.END)
        self.netwEntry.insert(0, netDir)
        self.netwEntry.config(state='readonly')
        if len(netDir) > self.entryLen:
            self.netwEntry.config(width=self.entryLen)  # limit width to adjust GUI

    def thread_go(self, thread):
        if self.run_thread is not None:
            if self.run_thread.is_alive():
                win32api.MessageBox(0, "Please wait until last mission complete !", "Please Wait",
                                    win32con.MB_OK)
            else:
                self.run_thread = thread
                thread.start()
        else:
            self.run_thread = thread
            thread.start()
    # def open_or_not(self,filename):
    #     if(win32gui.FindWindow(None,"excel")

    def open_spread_thread(self):
        thread_temp = Thread(os.startfile(self.fileEntry.get()))
        thread_temp.setDaemon(True)
        self.thread_go(thread_temp)

    def do_clean(self):
        self.progress_bar.start()
        self.write_status_to_text("Start cleaning all the noise cases:")
        self.clean_robot.open_new_window()
        var = [
            "Bon Tool Company 997 To Amazon",
            "Amware Logistics Unknown To Unknown",
            "Almo Unknown To Unknown",
            "Amazon Unknown To Unknown",
            "Amazon.ca Unknown To Unknown",
            "Chewy.com Unknown To Unknown",
            "Home Depot Canada Unknown To Unknown",
            "Medline Unknown To Unknown",
            "P2P - Cat5 Commerce Unknown To Unknown",
            "Tractor Supply Drop Ship Unknown To Unknown",
            "Unknown Unknown To Unknown",
            "Walmart Unknown To Unknown",
            "Kroger Unknown To Unknown",
            "To For Life Products",
            "TM File processing",
            "To Base Brands CC",
            "To Nurse Assist, Inc.",
            "3PL Central 997 To Cali Bamboo",
            "iTrade Network Unknown To Phillips Foods, Inc",
            "Unknown Unknown To Total Quality Logistics 2",
            "iTrade Network Unknown To Phillips Foods, Inc fka Phillips Seafood",
            "Digi-Key Corporation Unknown To Unknown",
            "Unknown Unknown To Bestseller",
            "Unknown Unknown To Abbyson Living Corporation",
            "CSN Unknown To Unknown",
            "Five Below 850 To Jem Accessories, Inc.",
            "No Subject",
            "Ace Bayou Corp 846 To Amazon"
        ]
        for search_key in var:
            self.clean_robot.change_criteria("contains", search_key)
            self.write_status_to_text(
                "Closing " + str(self.clean_robot.clean_all_case()) + " cases with key word: " + search_key)
        self.clean_robot.change_criteria("is not empty", "Hello")
        win32api.MessageBox(0, "No more noise in queue. :)", "Cleaning Done", win32con.MB_OK)
        self.progress_bar.stop()
        self.write_status_to_text("Complete cleaning all the noise cases!")
        self.write_status_to_text("----------------------------------------")

    def new_clean_thread(self):
        clean_thread = threading.Thread(target=self.do_clean)
        clean_thread.setDaemon(True)
        self.thread_go(clean_thread)

    def do_update(self):
        self.write_status_to_text("Start updating all the notes for PSA tasks in NetSuite...")
        self.clean_robot.open_new_window()
        self.progress_bar.start()
        self.excel_robot.update_all_notes()
        self.clean_robot.close_script_window()
        self.progress_bar.stop()
        self.write_status_to_text("Updated all the notes for PSA tasks in NetSuite successfully !")
        self.write_status_to_text("----------------------------------------")

    def update_spread_thread(self):
        thread_update = threading.Thread(target=self.do_update)
        thread_update.setDaemon(True)
        self.thread_go(thread_update)

    def do_grab(self):
        self.write_status_to_text("Start taking all the PSA tasks...")
        self.progress_bar.start()
        self.clean_robot.open_new_window()
        self.clean_robot.take_task()
        self.progress_bar.stop()
        self.clean_robot.close_script_window()
        self.write_status_to_text("Complete taking all the PSA tasks!")
        self.write_status_to_text("----------------------------------------")

    def grab_task_thread(self):
        thread_grab = threading.Thread(target=self.do_grab)
        thread_grab.setDaemon(True)
        self.thread_go(thread_grab)

    def write_status_to_text(self, statusmsg):
        current_time = get_current_time()
        statusmsg_in = str(current_time) + " " + str(statusmsg) + "\n"  # 换行
        self.log_data_Text.insert(END, statusmsg_in)

    # Create Queue instance  
    def use_queues(self, loops=5):
        # Now using a class member Queue        
        while True:
            print(self.gui_queue.get())

    # Button callback

    def click_me(self):
        self.action.configure(text='Hello ' + self.name.get())
        print(self)
        # self.create_thread()                # now called from imported module
        bq.write_to_scrol(self)

        # Spinbox callback

    def _spin(self):
        value = self.spin.get()
        self.scrol.insert(tk.INSERT, value + '\n')

    # GUI Callback  
    def checkCallback(self, *ignored_args):
        # only enable one checkbutton
        if self.chVarUn.get():
            self.check3.configure(state='disabled')
        else:
            self.check3.configure(state='normal')
        if self.chVarEn.get():
            self.check2.configure(state='disabled')
        else:
            self.check2.configure(state='normal')

        # Radiobutton Callback

    def radCall(self):
        radSel = self.radVar.get()
        if radSel == 0:
            self.mighty2.configure(text='Blue')
        elif radSel == 1:
            self.mighty2.configure(text='Gold')
        elif radSel == 2:
            self.mighty2.configure(text='Red')

        # update progressbar in callback loop

    def run_progressbar(self):
        self.progress_bar["maximum"] = 100
        for i in range(101):
            sleep(0.05)
            self.progress_bar["value"] = i  # increment progressbar
            self.progress_bar.update()  # have to call update() in loop
        self.progress_bar["value"] = 0  # reset/clear progressbar

    def start_progressbar(self):
        self.progress_bar.start()

    def stop_progressbar(self):
        self.progress_bar.stop()

    def do_collect(self):
        self.write_status_to_text("Start collecting emails for PSA tasks in NetSuite...")
        self.clean_robot.open_new_window()
        self.progress_bar.start()
        self.excel_robot.send_all_tps()
        self.clean_robot.close_script_window()
        self.progress_bar.stop()
        self.write_status_to_text("Complete collecting emails for PSA tasks in NetSuite")
        self.write_status_to_text("----------------------------------------")

    def new_collect_thread(self):
        thread_temp = Thread(target=self.do_collect)
        thread_temp.setDaemon(True)
        self.thread_go(thread_temp)

    def do_info(self):
        self.write_status_to_text("Start grabbing project and other information:")
        self.excel_robot.grab_task_name_ID()
        self.write_status_to_text("Complete grabbing project and other information:")

    def new_info_thread(self):
        info_thread = threading.Thread(target=self.do_info)
        info_thread.start()

    def new_open_thread(self):
        self.clean_robot.open_new_window()

    def progressbar_stop_after(self, wait_ms=1000):
        self.win.after(wait_ms, self.progress_bar.stop)

    def usingGlobal(self):
        global GLOBAL_CONST
        GLOBAL_CONST = 777

    # Exit GUI cleanly
    def _quit(self):
        self.win.quit()
        self.win.destroy()
        exit()

        #####################################################################################

    def create_widgets(self):
        tabControl = ttk.Notebook(self.win)  # Create Tab Control
        tab2 = ttk.Frame(tabControl)  # Add a second tab
        tabControl.add(tab2, text='General')  # Make second tab visible
        tab1 = ttk.Frame(tabControl)  # Create a tab
        tabControl.add(tab1, text='CleanNoise')  # Add the tab
        tab3 = ttk.Frame(tabControl)
        tabControl.add(tab3, text='NewTP')  # Add the tab
        tabControl.pack(expand=1, fill="both")  # Pack to make visible

        # LabelFrame using tab1 as the parent
        mighty = ttk.LabelFrame(tab1, text=' Noise List ')
        mighty.grid(column=0, row=0, padx=8, pady=4, sticky='WE')

        # Modify adding a Label using mighty as the parent instead of win
        a_label = ttk.Label(mighty, text="Enter a keyword:")
        a_label.grid(column=0, row=0, sticky='W')
        self.button_len = 25
        # Adding a Textbox Entry widget
        self.name = tk.StringVar()
        self.name_entered = ttk.Entry(mighty, width=24, textvariable=self.name)
        self.name_entered.grid(column=0, row=1, sticky='W', columnspan=2)
        self.name_entered.delete(0, tk.END)
        self.name_entered.insert(0, '< default name >')

        # Adding a Button
        self.action = ttk.Button(mighty, text="Add", command=self.click_me)
        self.action.grid(column=3, row=1)

        ttk.Label(mighty, text="Choose a Filter:").grid(column=2, row=0)
        number = tk.StringVar()
        self.number_chosen = ttk.Combobox(mighty, width=14, textvariable=number, state='readonly')
        self.number_chosen['values'] = (1, 2, 4, 42, 100)
        self.number_chosen.grid(column=2, row=1)
        self.number_chosen.current(0)
        ttk.Label(mighty, text="Expired Days: ").grid(column=0, row=2, sticky='W')
        # Adding a Spinbox widget
        self.spin = Spinbox(mighty, values=(1, 2, 4, 42, 100), width=5, bd=9, command=self._spin)  # using range
        self.spin.grid(column=1, row=2, sticky='W')  # align left

        # Using a scrolled Text control    
        scrol_w = 62;
        scrol_h = 15  # increase sizes
        self.scrol = scrolledtext.ScrolledText(mighty, width=scrol_w, height=scrol_h - 5, wrap=tk.WORD)
        self.scrol.grid(column=0, row=3, sticky='WE', columnspan=100)

        for child in mighty.winfo_children():  # add spacing to align widgets within tabs
            child.grid_configure(padx=4, pady=2)

            # =====================================================================================
        # Tab Control 2 ----------------------------------------------------------------------
        self.mighty2 = ttk.LabelFrame(tab2, text=' Mission Log ')
        self.mighty2.grid(column=0, row=2, padx=8, pady=4)

        # Add a textbox for log

        self.log_data_Text = scrolledtext.ScrolledText(self.mighty2, width=scrol_w, height=scrol_h, wrap=tk.WORD)  # 日志框
        self.log_data_Text.grid(row=3, column=0, columnspan=10)

        # Create a container to hold buttons
        self.buttons_frame = ttk.LabelFrame(tab2, text=' Automation Mission ')
        self.buttons_frame.grid(column=0, row=1, sticky='WE', padx=10, pady=5)

        self.mngFilesFrame = ttk.LabelFrame(tab2, text=' Manage Spread Sheet: ')
        self.mngFilesFrame.grid(column=0, row=0, sticky='WE', padx=10, pady=5)

        # Add a Progressbar to Tab 2
        self.progress_bar = ttk.Progressbar(self.buttons_frame, orient='horizontal', length=515, mode='determinate')
        self.progress_bar.grid(column=0, row=3, columnspan=3)

        # Add Buttons for Progressbar commands
        # ttk.Button(self.buttons_frame, text=" Run Progressbar   ", command=self.run_progressbar,
        #            width=self.button_len).grid(column=0, row=0, sticky='W')
        # ttk.Button(self.buttons_frame, text=" Stop immediately ", command=self.stop_progressbar,
        #            width=self.button_len).grid(column=1, row=0, sticky='W')
        # ttk.Button(self.buttons_frame, text=" Stop after second ", command=self.progressbar_stop_after,
        #            width=self.button_len).grid(column=2, row=0, sticky='W')
        # ttk.Button(self.buttons_frame, text=" Start Progressbar  ", command=self.start_progressbar,
        #            width=self.button_len).grid(column=0, row=1, sticky='W')
        ttk.Button(self.buttons_frame, text=" Grab all PSA tasks  ", command=self.grab_task_thread,
                   width=self.button_len).grid(column=1, row=1, sticky='W')
        ttk.Button(self.buttons_frame, text=" Update Netsuite Note ", command=self.update_spread_thread,
                   width=self.button_len).grid(column=0, row=0, sticky='W')
        ttk.Button(self.buttons_frame, text=" Clean Noise Cases ", command=self.new_clean_thread,
                   width=self.button_len).grid(column=1, row=0, sticky='W')
        ttk.Button(self.buttons_frame, text="New TP EDI Collection ", command=self.new_collect_thread,
                   width=self.button_len).grid(column=0, row=1, sticky='W')
        ttk.Button(self.buttons_frame, text="Display Script Browser Window ",
                   command=self.clean_robot.display_script_window,
                   width=self.button_len).grid(column=2, row=0, sticky='W')
        ttk.Button(self.buttons_frame, text="Grab project name and ID ", command=self.new_info_thread,
                   width=self.button_len).grid(column=0, row=2, sticky='W')
        ttk.Button(self.buttons_frame, text="Hide Script Browser Window ", command=self.clean_robot.hide_script_window,
                   width=self.button_len).grid(column=2, row=1, sticky='W')
        #
        # for child in self.buttons_frame.winfo_children():
        #     child.grid_configure(padx=2, pady=2)
        # Create Manage Files Frame ------------------------------------------------

        ttk.Button(self.mngFilesFrame, text=" Open Spreadsheet   ", command=self.open_spread_thread,
                   width=self.button_len).grid(column=0, row=3,
                                               sticky='W')
        ttk.Button(self.mngFilesFrame, text=" Synchronize Spreadsheet  ", command=self.start_progressbar,
                   width=self.button_len).grid(column=2,
                                               row=3,
                                               sticky='W')



        def new_send_thread():
            take_thread = threading.Thread(target=do_send)
            take_thread.start()

        def do_send():
            self.write_status_to_text("Start sending a log email to Amy:")
            robot = SendEmails()
            robot.send_amy_log()
            self.write_status_to_text("Complete sending a log email to Amy!")

        ttk.Button(self.mngFilesFrame, text=" Send a log backup ", command=new_send_thread,
                   width=self.button_len).grid(column=0, row=4, sticky='W')

        def backupFile():
            import shutil
            src = self.fileEntry.get()
            file = src.split('\\')[-1]
            dst = self.netwEntry.get() + '\\' + datetime.now().strftime("%b_%d_%Y") + file

            try:
                shutil.copy(src, dst)
                msg.showinfo('Copy File to Network', 'Succes: File copied.')
            except FileNotFoundError as err:
                msg.showerror('Copy File to Network', '*** Failed to copy file! ***\n\n' + str(err))
            except Exception as ex:
                msg.showerror('Copy File to Network', '*** Failed to copy file! ***\n\n' + str(ex))

        ttk.Button(self.mngFilesFrame, text=" Backup Spreadsheet ", command=backupFile,
                   width=self.button_len).grid(column=1, row=3,
                                               sticky='W')

        # Button Callback
        def getFileName():
            print('hello from getFileName')
            fDir = path.dirname(__file__)
            fName = fd.askopenfilename(parent=self.win, initialdir=fDir)
            if len(fName) != 0:
                self.fileEntry.config(state='enabled')
                self.fileEntry.delete(0, tk.END)
                self.fileEntry.insert(0, fName)

        # Add Widgets to Manage Files Frame
        lb = ttk.Button(self.mngFilesFrame, text="Choose Spreadsheet", command=getFileName, width=self.button_len)
        lb.grid(column=2, row=0, sticky=tk.W)

        # -----------------------------------------------------
        file = tk.StringVar()
        self.entryLen = 55
        self.fileEntry = ttk.Entry(self.mngFilesFrame, width=self.entryLen, textvariable=file)
        self.fileEntry.grid(column=0, row=0, sticky=tk.W, columnspan=2)

        def copyFile():
            import shutil
            fDir = path.dirname(path.dirname(__file__)) + '\\Backup'
            fName = fd.askdirectory(parent=self.win, initialdir=fDir)
            print(fName)
            self.netwEntry.config(state='enabled')
            self.netwEntry.delete(0, tk.END)
            self.netwEntry.insert(0, fName)
            if len(fName) != 0:
                self.netwEntry.config(state='enabled')
                self.netwEntry.delete(0, tk.END)
                self.netwEntry.insert(0, fName)

        cb = ttk.Button(self.mngFilesFrame, text="Choose Backup Folder", command=copyFile, width=self.button_len)
        cb.grid(column=2, row=1, sticky=tk.W)

        # -----------------------------------------------------
        logDir = tk.StringVar()
        self.netwEntry = ttk.Entry(self.mngFilesFrame, width=self.entryLen, textvariable=logDir)
        self.netwEntry.grid(column=0, row=1, sticky=tk.W, columnspan=2)

        # Add some space around each label
        for child in self.mngFilesFrame.winfo_children():
            child.grid_configure(padx=8, pady=4)

        for child in self.buttons_frame.winfo_children():
            child.grid_configure(padx=8, pady=4)

        for child in self.mighty2.winfo_children():
            child.grid_configure(padx=8, pady=4)
        # Creating a Menu Bar ==========================================================
        menu_bar = Menu(self.win)
        self.win.config(menu=menu_bar)

        # Add menu items
        file_menu = Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="New")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self._quit)
        menu_bar.add_cascade(label="File", menu=file_menu)

        # Display a Message Box
        def _msgBox():
            msg.showinfo('About this tool', 'A Python Automation tool created by Paul Wu.\nVersion: 1.0.0 ')

        # Add another Menu to the Menu Bar and an item
        help_menu = Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="About", command=_msgBox)  # display messagebox when clicked
        menu_bar.add_cascade(label="Help", menu=help_menu)

        # Change the main windows icon
        self.win.iconbitmap('pyc.ico')

        # It is not necessary to create a tk.StringVar()
        # strData = tk.StringVar()
        strData = self.spin.get()

        # call function
        self.usingGlobal()

        # self.name_entered.focus()
        # Set focus to Tab 2
        tabControl.select(0)

        # Add Tooltips -----------------------------------------------------
        # Add a Tooltip to the Spinbox
        tt.create_ToolTip(self.spin, 'This is a Spinbox control')

        # Add Tooltips to more widgets
        tt.create_ToolTip(self.name_entered, 'This is an Entry control')
        tt.create_ToolTip(self.action, 'This is a Button control')
        tt.create_ToolTip(self.scrol, 'This is a ScrolledText control')

# ======================
# Start GUI
# ======================
