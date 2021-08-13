import openpyxl
import time
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import win32api, win32con
import sys
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys

class TakeTasks:

    def __init__(self):
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9000")
        chrome_driver = r'.\drivers\chromedriver.exe'
        self.driver = webdriver.Chrome(executable_path=chrome_driver, options=chrome_options)
        self.workbook = openpyxl.load_workbook('Paul_Spread_Sheet_Senior_Version.xlsm', data_only=True)
        self.testbook = openpyxl.load_workbook('pytest.xlsx')
        self.worksheet = self.workbook.get_sheet_by_name('task')
        self.testsheet = self.testbook.get_sheet_by_name('Sheet1')

    def get_task(self):
        count = 0
        for item in list(self.worksheet.columns)[0]:
            if item.value == 'PSA Task Name':
                continue

            if item.value is None:
                break
            print(str(item.value))
            self.testsheet['A' + str(count)] = str(item.value)
            count += 1

        self.testsheet.save('pytest.xlsx')

    def edit_note(self, task_num, new_note):
        edit_icon = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[1]/div[2]/div/div/div/div/table/tbody/tr["
        edit_icon_ex = "]/td[2]/a[1]"
        table_content = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[1]/div[" \
                        "2]/div/div/div/div"
        ele = self.driver.find_element(By.XPATH, edit_icon + str(task_num) + edit_icon_ex)
        ele.click()
        text_box = "/html/body/div[1]/div[2]/div[3]/table[1]/tbody/tr[3]/td/div[1]/div/div/form/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/div/span[2]/span/textarea"
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, text_box))
            )
        finally:
            # time.sleep(1)
            ele = self.driver.find_element(By.XPATH, text_box)
            if ele.text is not None and ele.text == new_note:
                self.driver.get("https://907826.app.netsuite.com/app/center/card.nl?sc=-29&whence=")
            else:
                ele.click()
                ele.send_keys(Keys.CONTROL+'a')
                ele.send_keys(new_note)
                # print(new_note)
                # time.sleep(10)
                ele = self.driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div[3]/table[2]/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/input")
                ele.click()
                time.sleep(1)
                self.driver.get("https://907826.app.netsuite.com/app/center/card.nl?sc=-29&whence=")
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, table_content))
            )
        finally:
            # time.sleep(1)
            ele = self.driver.find_element(By.XPATH, table_content)

    def update_all_notes(self):
        cur_handle = self.driver.current_window_handle  # get current handle
        all_handle = self.driver.window_handles  # get all handles
        target_url = "https://907826.app.netsuite.com/app/center/card.nl?sc=-29&whence="
        for h in all_handle:
            if h != cur_handle:
                self.driver.switch_to.window(h)  # Switch to the new pop-up window
                break
        # 2 | open | /app/center/card.nl?sc=-29&whence= |\
        # time.sleep(2)
        self.driver.get(target_url)
        if not ("https://907826.app.netsuite.com/app/center/" in self.driver.current_url):
            win32api.MessageBox(0, "Please login first and try again. :)", "Please Login",
                                win32con.MB_OK)
            return 'Mission Failed'
        table_content = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[1]/div[" \
                        "2]/div/div/div/div"

        number_sum = "/html/body/div[1]/div[2]/div/div/div/div[5]/div[2]/div[1]/div[2]/div/div/form/div[" \
                     "2]/table/tbody/tr/td/table/tbody/tr/td/a "
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, number_sum))
            )
        finally:
            # time.sleep(1)
            ele = self.driver.find_element(By.XPATH, number_sum)
        html = ele.get_attribute('innerHTML')
        case_sum = int(html)
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, table_content))
            )
        finally:
            # time.sleep(1)
            ele = self.driver.find_element(By.XPATH, table_content)
        html = ele.get_attribute('innerHTML')
        soup = BeautifulSoup(html, 'html5lib')
        tables = soup.findAll('table')
        tab = tables[0]
        table_body = tab.tbody
        tr_group = table_body.find_all('tr')
        target = int(len(tr_group) - 1)  # number of tr
        task_name = []
        print(target)
        for tr in tr_group:
            if not ("text" in tr['class']):
                td_group = tr.find_all('td')
                task_name.append(td_group[2].text)
        print(task_name)
        for num in range(0, case_sum):
            for excel_num in range(0, case_sum):
                if str.strip(self.worksheet['A' + str(excel_num + 2)].value) == str.strip(task_name[num]):
                    if self.worksheet['D' + str(excel_num+2)].value is not None:
                        self.edit_note(num + 1, self.worksheet['D' + str(excel_num+2)].value)
                    break

