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
        target_url = "https://907826.app.netsuite.com/app/center/card.nl?sc=-29&whence="
        self.driver.get(target_url)
        for h in all_handle:
            if h != cur_handle:
                self.driver.switch_to.window(h)  # Switch to the new pop-up window
                break
        # 2 | open | /app/center/card.nl?sc=-29&whence= |\
        time.sleep(2)
        self.driver.set_window_size(1920,1080)
        self.driver.set_window_position(-2000, -2000)
        if self.driver.current_url != target_url:
            win32api.MessageBox(0, "Please login first and try again. :)", "Please Login",
                                win32con.MB_OK)
            sys.exit(0)

    def teardown_method(self):
        # Step # | name | target | value
        self.driver.quit()
        # 1 | close | Chrome with debugger address |\

    def refresh_list(self):
        # Step # | name | target | value
        js_top = "var q=document.documentElement.scrollTop=0"
        self.driver.execute_script(js_top)
        # 1 | scroll | Scroll to the top of window |\
        element = self.driver.find_element(By.XPATH, "//div[2]/div/div/h2")
        actions = ActionChains(self.driver)
        actions.move_to_element(element).perform()
        # 2 | MouseMoveAt | Title: Paul's All Case View | hover element
        time.sleep(2)
        refresh_icon = "//div[2]/div/div/div/span[3]"
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, refresh_icon))
            )
        finally:
            element = self.driver.find_element(By.XPATH, refresh_icon)
            actions = ActionChains(self.driver)
            actions.move_to_element(element).perform()
            element.click()
        # 3 | move mouse and click | Refresh Icon | hover element

    def take_task(self):
        # Step # | name | target | value
        tab_case = "/html/body/div[1]/div[1]/div[2]/ul[4]/li[2]/a/span"
        self.driver.find_element(By.XPATH, tab_case).click()
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, tab_case))
            )
        finally:
            self.driver.find_element(By.XPATH, tab_case).click()
        # 1 | click | case tab |
        number_sum = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[1]/div[2]/div/div/form/div[2]/table[" \
                     "2]/tbody/tr/td/table/tbody/tr/td/a "
        # first_table = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[1]/div[2]/div/div/div/div/table"
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, number_sum))
            )
        finally:
            ele = self.driver.find_element(By.XPATH, number_sum)
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
                table_content = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[1]/div[" \
                                "2]/div/div/div/div/table/tbody "
                first_row_inner_xpath = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[1]/div[" \
                                        "2]/div/div/div/div/table/tbody/tr[1]/td "
                ele = self.driver.find_element(By.XPATH, first_row_inner_xpath)
                text = ele.get_attribute('innerHTML')
                if text == "No Search Results Match Your Criteria.":
                    break
                    # win32api.MessageBox(0, "No more case in queue. :)", "Cleaning Done", win32con.MB_OK)
                    # sys.exit(0)
                ele = self.driver.find_element(By.XPATH, table_content)
                html = ele.get_attribute('innerHTML')
                soup = BeautifulSoup(html, 'html5lib')
                target = int(len(soup.find_all('span')) / 4)  # number of tr
                # print(target)
                # 2 | count | case number in one page |
            first_row_xpath = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[1]/div[" \
                              "2]/div/div/div/div/table/tbody/tr[1]/td[8]/span "
            last_row_xpath = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[1]/div[" \
                             "2]/div/div/div/div/table/tbody/tr[" + str(target) + "]/td[8]/span"
            # print(last_row_xpath)
            input_box = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[1]/div[" \
                        "2]/div/div/div/div/table/tbody/tr[1]/td[8]/span/div/span/div[1]/input "
            select_close = "/html/body/div[7]/div/div/div[15]"
            # 3 | click | first row |
            self.driver.find_element(By.XPATH, first_row_xpath).click()
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, input_box))
                )
            finally:
                self.driver.find_element(By.XPATH, input_box).send_keys("Closed")
                time.sleep(2)
                self.driver.find_element(By.XPATH, input_box).click()
            # 4 | shift + last line

            ele = self.driver.find_element(By.XPATH, last_row_xpath)
            action_chains = ActionChains(self.driver)
            action_chains.key_down(Keys.SHIFT).click(ele).key_up(Keys.SHIFT).perform()
            # 5 | click | id=uir_totalcount |
            js_top = "var q=document.documentElement.scrollTop=0"
            self.driver.execute_script(js_top)
            self.driver.find_element(By.ID, "uir_totalcount").click()
            # 6 |  refresh the list
            time.sleep(2)
            self.refresh_list()
            time.sleep(2)
            self.refresh_list()
            time.sleep(2)
            self.refresh_list()
            self.driver.find_element(By.XPATH, tab_case).click()
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, number_sum))
                )
            finally:
                ele = self.driver.find_element(By.XPATH, number_sum)
            html = ele.get_attribute('innerHTML')
            case_sum = int(html)
            # 7 |  update the case number



