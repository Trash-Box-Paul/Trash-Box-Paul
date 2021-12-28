import time
import sys
import webdrivermanager
import win32api, win32con
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from datetime import datetime
from debug_browser import DebugBrowser
import testraw
from pynput.keyboard import Key, Controller


class Worker(object):

    def __init__(self):
        self.keyboard = Controller()
        self.driver = DebugBrowser().driver
        #  Use selenium manager to check the version of chrome and selenium
        target_url = "https://907826.app.netsuite.com/app/center/card.nl?sc=-29&whence="
        self.driver.maximize_window()
        self.driver.get(target_url)
        if not ("https://907826.app.netsuite.com/app/center/" in self.driver.current_url):
            win32api.MessageBox(0, "Please login first and try again. :)", "Please Login",
                                win32con.MB_OK)
            sys.exit(0)
        self.root = self.driver.current_window_handle
