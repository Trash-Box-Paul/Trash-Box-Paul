import os
import sys

current_dir = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(current_dir)[0]
sys.path.append(rootPath)

import tkinter as tk
from gui import Application

root = tk.Tk()
root.title("TPS Team Tool")
app = Application(master=root)
app.mainloop()
