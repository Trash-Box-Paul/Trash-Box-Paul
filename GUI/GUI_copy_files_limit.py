import win32api, win32con
import tkinter as tk
from tkinter import ttk
from tkinter import scrolledtext
from tkinter import Menu
from tkinter import messagebox as msg
from tkinter import Spinbox
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

# Module level GLOBALS
GLOBAL_CONST = 42
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
        self.win = tk.Tk()

        self.run_thread = None

        # Add a title       
        self.win.title("TPS Automation 1.0.0")

        # Create a Queue
        self.gui_queue = Queue()

        self.create_widgets()

        self.defaultFileEntries()

    def defaultFileEntries(self):
        self.fileEntry.delete(0, tk.END)
        self.fileEntry.insert(0, fDir + '\Paul_Spread_Sheet_Senior_Version.xlsm')
        if len(fDir) > self.entryLen:
            #             self.fileEntry.config(width=len(fDir) + 3)
            self.fileEntry.config(width=self.entryLen)  # limit width to adjust GUI
            self.fileEntry.config(state='readonly')

        self.netwEntry.delete(0, tk.END)
        self.netwEntry.insert(0, netDir)
        if len(netDir) > self.entryLen:
            #             self.netwEntry.config(width=len(netDir) + 3)
            self.netwEntry.config(width=self.entryLen)  # limit width to adjust GUI

    def thread_go(self, thread):
        if self.run_thread is not None:
            if self.run_thread.isAlive():
                win32api.MessageBox(0, "Please wait until last mission complete)", "Please Wait",
                                    win32con.MB_OK)
            else:
                self.run_thread = thread
                self.run_thread.start()
        else:
            self.run_thread = thread
            self.run_thread.start()

    def open_spread_thread(self):
        thread_temp = Thread(os.startfile(self.fileEntry.get()))
        thread_temp.setDaemon(True)
        self.thread_go(thread_temp)

    def update_spread_thread(self):
        thread_temp = Thread(self.do_update())
        thread_temp.setDaemon(True)
        self.thread_go(thread_temp)

    def write_status_to_text(self, statusmsg):
        current_time = get_current_time()
        statusmsg_in = str(current_time) + " " + str(statusmsg) + "\n"  # 换行
        self.log_data_Text.insert(END, statusmsg_in)

    # Create Queue instance  
    def use_queues(self, loops=5):
        # Now using a class member Queue        
        while True:
            print(self.gui_queue.get())

    def method_in_a_thread(self, num_of_loops=10):
        for idx in range(num_of_loops):
            sleep(1)
            self.scrol.insert(tk.INSERT, str(idx) + '\n')

            # Running methods in Threads

    def create_thread(self, num=1):
        self.run_thread = Thread(target=self.method_in_a_thread, args=[num])
        self.run_thread.setDaemon(True)
        self.run_thread.start()

        # start queue in its own thread
        write_thread = Thread(target=self.use_queues, args=[num], daemon=True)
        write_thread.start()

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

    def do_update(self):
        robot = excel_test.TakeTasks()
        robot.update_all_notes()
        self.write_status_to_text("Updated all the PSA tasks successfully !")

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
        scrol_w = 65;
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

        # Add a Progressbar to Tab 2
        self.progress_bar = ttk.Progressbar(self.buttons_frame, orient='horizontal', length=530, mode='determinate')
        self.progress_bar.grid(column=0, row=2, columnspan=100)

        # Add Buttons for Progressbar commands
        ttk.Button(self.buttons_frame, text=" Run Progressbar   ", command=self.run_progressbar,
                   width=self.button_len).grid(column=0, row=0, sticky='W')
        ttk.Button(self.buttons_frame, text=" Start Progressbar  ", command=self.start_progressbar,
                   width=self.button_len).grid(column=0, row=1, sticky='W')
        ttk.Button(self.buttons_frame, text=" Stop immediately ", command=self.stop_progressbar,
                   width=self.button_len).grid(column=1, row=0, sticky='W')
        ttk.Button(self.buttons_frame, text=" Stop after second ", command=self.progressbar_stop_after,
                   width=self.button_len).grid(column=2, row=0, sticky='W')
        #
        # for child in self.buttons_frame.winfo_children():
        #     child.grid_configure(padx=2, pady=2)

        for child in self.mighty2.winfo_children():
            child.grid_configure(padx=8, pady=2)

        # Create Manage Files Frame ------------------------------------------------
        self.mngFilesFrame = ttk.LabelFrame(tab2, text=' Manage Spread Sheet: ')

        self.mngFilesFrame.grid(column=0, row=0, sticky='WE', padx=10, pady=5)

        ttk.Button(self.mngFilesFrame, text=" Open Spreadsheet   ", command=self.open_spread_thread,
                   width=self.button_len).grid(column=0, row=3,
                                               sticky='W')
        ttk.Button(self.mngFilesFrame, text=" Synchronize Spreadsheet  ", command=self.start_progressbar,
                   width=self.button_len).grid(column=2,
                                               row=3,
                                               sticky='W')
        ttk.Button(self.mngFilesFrame, text=" Update Netsuite Note ", command=self.update_spread_thread,
                   width=self.button_len).grid(column=1, row=3,
                                               sticky='W')
        ttk.Button(self.mngFilesFrame, text=" Create New Spreadsheet ", command=self.update_spread_thread,
                   width=self.button_len).grid(column=0, row=4,
                                               sticky='W')

        # Button Callback
        def getFileName():
            print('hello from getFileName')
            fDir = path.dirname(__file__)
            fName = fd.askopenfilename(parent=self.win, initialdir=fDir)
            print(fName)
            if fName is not None:
                self.fileEntry.config(state='enabled')
                self.fileEntry.delete(0, tk.END)
                self.fileEntry.insert(0, fName)


        # Add Widgets to Manage Files Frame
        lb = ttk.Button(self.mngFilesFrame, text="Browse to Spreadsheet...", command=getFileName, width=self.button_len)
        lb.grid(column=0, row=0, sticky=tk.W)

        # -----------------------------------------------------
        file = tk.StringVar()
        self.entryLen = 60
        self.fileEntry = ttk.Entry(self.mngFilesFrame, width=self.entryLen, textvariable=file)
        self.fileEntry.grid(column=1, row=0, sticky=tk.W, columnspan=100)

        def copyFile():
            import shutil
            fDir = path.dirname(path.dirname(__file__))+'\\Backup'
            fName = fd.askopenfile(parent=self.win, initialdir=fDir)
            print(fName)
            self.fileEntry.config(state='enabled')
            self.fileEntry.delete(0, tk.END)
            self.fileEntry.insert(0, fName)
            if fName is not None:
                self.fileEntry.config(state='enabled')
                self.fileEntry.delete(0, tk.END)
                self.fileEntry.insert(0, fName)
            src = self.fileEntry.get()
            file = src.split('\\')[-1]
            dst = self.netwEntry.get() + '\\' + file

            try:
                shutil.copy(src, dst)
                msg.showinfo('Copy File to Network', 'Succes: File copied.')
            except FileNotFoundError as err:
                msg.showerror('Copy File to Network', '*** Failed to copy file! ***\n\n' + str(err))
            except Exception as ex:
                msg.showerror('Copy File to Network', '*** Failed to copy file! ***\n\n' + str(ex))

        cb = ttk.Button(self.mngFilesFrame, text="Copy Spreadsheet To :   ", command=copyFile, width=self.button_len)
        cb.grid(column=0, row=1, sticky=tk.W)

        # -----------------------------------------------------
        logDir = tk.StringVar()
        self.netwEntry = ttk.Entry(self.mngFilesFrame, width=self.entryLen, textvariable=logDir)
        self.netwEntry.grid(column=1, row=1, sticky=tk.W, columnspan=100)

        # Add some space around each label
        for child in self.mngFilesFrame.winfo_children():
            child.grid_configure(padx=6, pady=6)

        for child in self.buttons_frame.winfo_children():
            child.grid_configure(padx=6, pady=6)

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
