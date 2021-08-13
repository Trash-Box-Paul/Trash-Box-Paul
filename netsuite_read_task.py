import time
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options


class TestLogin:
    def setup_method(self):
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9000")
        chrome_driver = r'.\drivers\chromedriver.exe'
        self.driver = webdriver.Chrome(executable_path=chrome_driver, options=chrome_options)
    # chrome.exe - -remote - debugging - port = 9000 - -user - data -
    # dir = "C:\Users\paul.wu\PycharmProjects\practice\seleinumChrome\AutomationProfile"

    def teardown_method(self):
        self.driver.quit()

    def refresh_list(self):
        js_top = "var q=document.documentElement.scrollTop=0"
        self.driver.execute_script(js_top)
        element = self.driver.find_element(By.XPATH, "//div[2]/div/div/h2")
        actions = ActionChains(self.driver)
        actions.move_to_element(element).perform()
        # 7 | mouseMoveAt | xpath=//div[2]/div/div/div/span[3] |
        time.sleep(2)
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[2]/div/div/div/span[3]"))
            )
        finally:
            element = self.driver.find_element(By.XPATH, "//div[2]/div/div/div/span[3]")
            actions = ActionChains(self.driver)
            actions.move_to_element(element).perform()
            element.click()

    def read_task(self):
        self.driver.get("https://907826.app.netsuite.com")
        self.driver.set_window_size(550, 691)

        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "i0116"))
            )
        finally:
            self.driver.find_element(By.ID, "i0116").click()
        time.sleep(2)
        self.driver.find_element(By.ID, "i0116").send_keys('Paul.Wu@truecommerce.com')
        time.sleep(2)
        self.driver.find_element(By.ID, "idSIButton9").click()

        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "i0118"))
            )
        finally:
            self.driver.find_element(By.ID, "i0118").send_keys("Smokingman527")
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "idSIButton9"))
            )
        finally:
            self.driver.find_element(By.ID, "idSIButton9").submit()
        self.driver.find_element(By.ID, "idBtn_Back").click()
