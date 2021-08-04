import pytest
import time
import sys
import json
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
from debug_browser import DebugBrowser


class TakeTasks:
    def __init__(self):
        # Step # | name | target | value
        chrome_driver = r'.\drivers\chromedriver.exe'
        # chrome_options = Options()
        # chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9000")
        # self.driver = webdriver.Chrome(executable_path=chrome_driver, options=chrome_options)
        self.driver = webdriver.Chrome(executable_path=chrome_driver, options=DebugBrowser().debug_chrome())
        # 1 | open | Chrome with debugger address |\
        # if not self.driver.toString().contains("null"):
        #     self.driver.quit()
        cur_handle = self.driver.current_window_handle  # get current handle
        all_handle = self.driver.window_handles  # get all handles
        target_url = "https://907826.app.netsuite.com/app/center/card.nl?sc=-29&whence=#"
        for h in all_handle:
            if h != cur_handle:
                self.driver.switch_to.window(h)  # Switch to the new pop-up window
                break
        # 2 | open | /app/center/card.nl?sc=-29&whence= |\
        time.sleep(2)
        self.driver.set_window_size(960, 1080)
        self.driver.set_window_position(0, 0)
        self.driver.get(target_url)
        if not ("https://907826.app.netsuite.com/app/center/" in self.driver.current_url):
            win32api.MessageBox(0, "Please login first and try again. :)", "Please Login",
                                win32con.MB_OK)
            sys.exit(0)

    def teardown_method(self):
        # Step # | name | target | value
        self.driver.close()
        # 1 | close | Chrome with debugger address |\

    def refresh_list_down(self):
        # Step # | name | target | value
        # 1 | scroll | Scroll to the top of window |\
        title_list = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[2]/div[1]/h2"
        refresh_icon = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[2]/div[1]/div/span[3]"
        element = self.driver.find_element(By.XPATH, refresh_icon)
        ele = self.driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[2]/div["
                                                 "1]/div")
        self.driver.execute_script("arguments[0].style.display='block';", ele)
        self.driver.execute_script("arguments[0].style.display='block';", element)
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, title_list))
            )
        finally:
            element = self.driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div["
                                                         "1]/div[2]")
            self.driver.execute_script("arguments[0].scrollIntoView(true)", element)
            element = self.driver.find_element(By.XPATH, title_list)
            actions = ActionChains(self.driver)
            actions.move_to_element(element).perform()
            # 2 | MouseMoveAt | Title: Paul's All Case View | hover element
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, refresh_icon))
            )
        finally:
            element = self.driver.find_element(By.XPATH, refresh_icon)
            # ele = self.driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[2]/div["
            #                                          "1]/div")
            # self.driver.execute_script("arguments[0].style.display='block';", ele)
            # self.driver.execute_script("arguments[0].style.display='block';", element)
            # time.sleep(3)
            actions = ActionChains(self.driver)
            actions.move_to_element(element).perform()
            element.click()
        # 3 | move mouse and click | Refresh Icon | hover element
        element = self.driver.find_element(By.XPATH, title_list)
        self.driver.execute_script("arguments[0].scrollIntoView(true)", element)
        # 3 | move mouse and click | Refresh Icon | hover element

    def take_task(self):
        # Step # | name | target | value
        tab_home = "/html/body/div[1]/div[1]/div[2]/ul[3]/li/a"
        # self.driver.find_element(By.XPATH, tab_home).click()
        # target_url = "https://907826.app.netsuite.com/app/center/card.nl?sc=-29&whence=#"
        # self.driver.get(target_url)
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, tab_home))
            )
        finally:
            self.driver.find_element(By.XPATH, tab_home).click()
            self.refresh_list_down()
        # 1 | click | case tab |
        number_sum = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[2]/div[2]/div/div/form/div[" \
                     "2]/table/tbody/tr/td/table/tbody/tr/td/a"
        # first_table = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[1]/div[2]/div/div/div/div/table"
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, number_sum))
            )
        finally:
            ele = self.driver.find_element(By.XPATH, number_sum)
            # self.driver.execute_script("arguments[0].scrollIntoView(true)", ele)
        html = ele.get_attribute('innerHTML')
        case_sum = int(html)
        # 2 | read | case number |
        while case_sum > 0:
            if case_sum <= 0:
                # win32api.MessageBox(0, "No more case in queue. :)", "Cleaning Done", win32con.MB_OK)
                # sys.exit(0)
                break
            else:
                # table_content = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[1]/div[
                # 2]/div/div/div/div/table"
                table_content = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[2]/div[2]/div/div/div"
                first_row_inner_xpath = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[2]/div[" \
                                        "2]/div/div/div/div/table/tbody/tr[1]/td"
                ele = self.driver.find_element(By.XPATH, first_row_inner_xpath)
                # self.driver.execute_script("arguments[0].scrollIntoView(true)", ele)
                text = ele.get_attribute('innerHTML')
                if text == "No Search Results Match Your Criteria.":
                    break
                    # win32api.MessageBox(0, "No more case in queue. :)", "Cleaning Done", win32con.MB_OK)
                    # sys.exit(0)
                first_pencil = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[2]/div[" \
                               "2]/div/div/div/div/table/tbody/tr[1]/td[2]/a[1]"
                input_name = "/html/body/div[1]/div[2]/div[3]/form/table/tbody/tr[2]/td/table/tbody/tr[" \
                             "1]/td/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/div/span[2]/span/div[1]/input"
                my_name = "/html/body/div[8]/div/div/table/tbody/tr/td"
                save_icon = "/html/body/div[1]/div[2]/div[3]/form/table/tbody/tr[" \
                            "1]/td/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/input "
                element = self.driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div["
                                                             "2]/div[1]/h2")
                self.driver.execute_script("arguments[0].scrollIntoView(true)", element)
                # print(last_row_xpath)
                # 3 | click | first row |
                try:
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, first_pencil))
                    )
                finally:
                    time.sleep(3)
                    element = self.driver.find_element(By.XPATH, first_pencil)
                    element.click()
                try:
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, input_name))
                    )
                finally:
                    self.driver.find_element(By.XPATH, input_name).send_keys(Keys.CONTROL + "a")
                    self.driver.find_element(By.XPATH, input_name).send_keys("Paul Wu")
                    time.sleep(3)
                    self.driver.find_element(By.XPATH, input_name).send_keys(Keys.ENTER)
                # try:
                #     WebDriverWait(self.driver, 10).until(
                #         EC.presence_of_element_located((By.XPATH, my_name))
                #     )
                # finally:
                #     time.sleep(2)
                #     self.driver.find_element(By.XPATH, my_name).click()
                # 4 | shift + last line
                self.driver.find_element(By.XPATH, save_icon).click()
                time.sleep(1)
                self.refresh_list_down()
                self.driver.find_element(By.XPATH, tab_home).click()
                try:
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, number_sum))
                    )
                finally:
                    ele = self.driver.find_element(By.XPATH, number_sum)
                    self.driver.execute_script("arguments[0].scrollIntoView(true)", ele)
                html = ele.get_attribute('innerHTML')
                case_sum = int(html)

            # 7 |  update the case number


