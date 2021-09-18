import os.path
import time
import os
import netsuite_clean_all_case
import win32com
from win32com.client import Dispatch, constants
from datetime import datetime
import pythoncom


def open_spread(path):
    os.startfile(path)
