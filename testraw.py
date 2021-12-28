import pytest
import time
import json
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
import webdrivermanager

class TestPsd:
    def __init__(self):
        firefox_driver = r'.\drivers\geckodriver.exe'
        firefox_options = Options()
        firefox_options.add_experimental_option("debuggerAddress", "127.0.0.1:9000")
        firefox_driver = webdrivermanager.GeckoDriverManager()
        try:
            self.driver = webdriver.Chrome(executable_path=firefox_driver.get_driver_filename(),
                                           options=firefox_options)
        except:
            firefox_driver.download_and_install(firefox_driver.get_latest_version())
            time.sleep(3)
            print("!!!!!!!!!!!!!!!!")
            time.sleep(3)
            self.driver = webdriver.Chrome(executable_path=firefox_driver.get_driver_filename(),
                                           options=firefox_options)
        cur_handle = self.driver.current_window_handle  # get current handle
        all_handle = self.driver.window_handles  # get all handles
        target_url = "http://psdtool.tc.net/psdTool/"
        self.driver.get(target_url)
        self.driver.maximize_window()
        self.driver.set_window_position(960, 0)
        for h in all_handle:
            if h != cur_handle:
                self.driver.switch_to.window(h)  # Switch to the new pop-up window
                break
        time.sleep(2)
        input_username = "/html/body/form/div[3]/table[1]/tbody/tr[2]/td/table/tbody " \
                         "/tr/td/table/tbody/tr[3]/td[2]/input"
        input_password = "/html/body/form/div[3]/table[1]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[4]/td[" \
                         "2]/input "
        icon_login = "/html/body/form/div[3]/table[1]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[6]/td/input"
        if self.driver.current_url != target_url:
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, input_username))
                )
            finally:
                self.driver.find_element(By.XPATH, input_username).send_keys(Keys.CONTROL + "a")
                self.driver.find_element(By.XPATH, input_username).send_keys("Paul.Wu")
                self.driver.find_element(By.XPATH, input_password).send_keys(Keys.CONTROL + "a")
                self.driver.find_element(By.XPATH, input_password).send_keys("Smokingman527!")
                self.driver.execute_script('arguments[0].click()', self.driver.find_element(By.XPATH, icon_login))
                time.sleep(2)

        # 1 | open | Chrome with debugger address |\
        # if not self.driver.toString().contains("null"):
        #     self.driver.quit()
        # cur_handle = self.driver.current_window_handle  # get current handle
        # all_handle = self.driver.window_handles  # get all handles
        # target_url = "http://psdtool.tc.net/psdTool/"
        # self.driver.get(target_url)

    def teardown_method(self):
        self.driver.quit()

    def psd_resend(self, log_groups):
        tm_tab = "/html/body/form/table/tbody/tr[2]/td[1]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/a"
        tm_tab_raw = "/html/body/form/table/tbody/tr[2]/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[" \
                     "2]/td/table/tbody/tr/td/a "
        tm_tab_all = "/html/body/form/table/tbody/tr[2]/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[" \
                     "1]/td/table/tbody/tr/td/a "
        input_all = "/html/body/form/table/tbody/tr[3]/td/div/table/tbody/tr/td/div/div/div[1]/input[1]"
        for log_id in log_groups:
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, tm_tab))
                )
            finally:
                element = self.driver.find_element(By.XPATH, tm_tab)
                actions = ActionChains(self.driver)
                actions.move_to_element(element).perform()
                # actions.click(element).perform()
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, tm_tab_all))
                )
            finally:
                element = self.driver.find_element(By.XPATH, tm_tab_all)
                actions = ActionChains(self.driver)
                actions.move_to_element(element)
                actions.click(element).perform()
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, input_all))
                )
            finally:
                time.sleep(2)
                self.driver.execute_script('arguments[0].click()', self.driver.find_element(By.XPATH, input_all))
                self.driver.find_element(By.XPATH, input_all).send_keys(Keys.CONTROL+"a")
                self.driver.find_element(By.XPATH, input_all).send_keys(log_id)
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "/html/body/form/table/tbody/tr["
                                                   "3]/td/div/table/tbody/tr/td/div/div/div[1]/input[2]"))
                )
            finally:
                self.driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr["
                                                   "3]/td/div/table/tbody/tr/td/div/div/div[1]/input[2]").click()

            psd_all_table = "/html/body/form/table/tbody/tr[3]/td/div/table/tbody/tr/td/div/div"
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "/html/body/form/table/tbody/tr[3]/td/div"))
                )
            finally:
                time.sleep(3)
                ele = self.driver.find_element(By.XPATH, psd_all_table)
                html = ele.get_attribute("innerHTML")
                soup = BeautifulSoup(html, 'html5lib')
                tables = soup.findAll('table')
                tab = tables[0]
                table_body = tab.tbody
                number_tr = int(len(table_body.find_all('tr'))) - 4
                print(number_tr)
            if number_tr == 1:
                element = self.driver.find_element(By.XPATH, tm_tab)
                actions = ActionChains(self.driver)
                actions.move_to_element(element).perform()
                try:
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, tm_tab_raw))
                    )
                finally:
                    element = self.driver.find_element(By.XPATH, tm_tab_raw)
                    actions = ActionChains(self.driver)
                    actions.move_to_element(element)
                    actions.click(element).perform()
                input_raw = "/html/body/form/table/tbody/tr[3]/td/div/table/tbody/tr/td/div/table[1]/tbody/tr[1]/td[" \
                            "2]/input[1] "
                icon_get = "/html/body/form/table/tbody/tr[3]/td/div/table/tbody/tr/td/div/table[1]/tbody/tr[1]/td[" \
                           "2]/input[2] "
                try:
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, input_raw))
                    )
                finally:
                    time.sleep(2)
                    self.driver.find_element(By.XPATH, input_raw).send_keys(Keys.CONTROL+"a")
                    self.driver.find_element(By.XPATH, input_raw).send_keys(log_id)
                    self.driver.find_element(By.XPATH, icon_get).click()
                    self.driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr["
                                                       "3]/td/div/table/tbody/tr/td/div/table[3]/tbody/tr[2]/td["
                                                       "1]/input") .send_keys(Keys.CONTROL+"a")
                    self.driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr["
                                                       "3]/td/div/table/tbody/tr/td/div/table[3]/tbody/tr[2]/td["
                                                       "1]/input") .send_keys("Paul.Wu")
                    self.driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr["
                                                       "3]/td/div/table/tbody/tr/td/div/table[3]/tbody/tr[2]/td["
                                                       "2]/input").send_keys(Keys.CONTROL+"a")
                    self.driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr["
                                                       "3]/td/div/table/tbody/tr/td/div/table[3]/tbody/tr[2]/td["
                                                       "2]/input") .send_keys("Smokingman527")
                self.driver.find_element(By.XPATH,
                                         "/html/body/form/table/tbody/tr[3]/td/div/table/tbody/tr/td/div/table["
                                         "3]/tbody/tr[3]/td/input[1]").click()

                self.driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr["
                                                   "3]/td/div/table/tbody/tr/td/div/table[ "
                                                   "3]/tbody/tr[4]/td/input").click()
                self.driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr["
                                                   "3]/td/div/table/tbody/tr/td/div/table[3]/tbody/tr["
                                                   "5]/td/input").click()




        #
        # self.driver.find_element(By.LINK_TEXT, "Raw Edi").click()
        # time.sleep(2)
        # element = self.driver.find_element(By.CSS_SELECTOR, "body")
        # actions = ActionChains(self.driver)
        # actions.move_to_element(element).perform()
        # self.driver.find_element(By.ID, "ctl00_BodyContent_ctl00_txtLogId").click()
        # self.driver.find_element(By.ID, "ctl00_BodyContent_ctl00_txtLogId").send_keys(log_id)
        # self.driver.find_element(By.ID, "ctl00_BodyContent_ctl00_btnGetData").click()
        # time.sleep(2)
        # # 4 | click | id=ctl00_BodyContent_ctl00_rbTNYes |
        # # 5 | click | id=ctl00_BodyContent_ctl00_chkWaiveCharges |
        # self.driver.find_element(By.ID, "ctl00_BodyContent_ctl00_chkWaiveCharges").click()
        # # 6 | click | id=ctl00_BodyContent_ctl00_btnSubmit |
        # # self.driver.find_element(By.ID, "ctl00_BodyContent_ctl00_btnSubmit").click()
        #
