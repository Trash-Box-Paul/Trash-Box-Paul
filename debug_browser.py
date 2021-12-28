import os
from selenium.webdriver.chrome.options import Options
import socket
import webdrivermanager
from selenium import webdriver
import time
from datetime import datetime


# Open the browser on port 9222, and save the information to C:/testfile
class Singleton(object):
    _INSTANCE = {}

    def __init__(self, cls):
        self.cls = cls

    def __call__(self, *args, **kwargs):
        instance = self._INSTANCE.get(self.cls, None)
        if not instance:
            instance = self.cls(*args, **kwargs)
            self._INSTANCE[self.cls] = instance
        return instance

    def __getattr__(self, key):
        return getattr(self.cls, key, None)


@Singleton
class DebugBrowser:
    def __init__(self):
        self.ip = '127.0.0.1'
        self.port = 9000
        self.user_file = 'C:/test'
        self.chrome_option = Options()
        self.chrome_address = 'C:\Program Files\Google\Chrome\Application'
        self.driver = self.driver_setup()
        self.time = datetime.now().strftime("%b_%d_%Y")

    def debug_chrome(self):
        """
        :return: chrome_option
        """
        # If the debugging port has been activated, directly add the monitoring option
        if self.check_port():
            self.chrome_option.add_experimental_option('debuggerAddress', '{}:{}'.format(self.ip, self.port))
        # Restart the browser if it is not started to listening to the debug port
        else:
            os.popen('cd {}'.format(self.chrome_address) +
                     ' && chrome.exe --remote-debugging-port={} --user-data-dir="{}"'.format(self.port, self.user_file))
            self.chrome_option.add_experimental_option('debuggerAddress', '{}:{}'.format(self.ip, self.port))
        return self.chrome_option

    def check_port(self):
        """
        Determine whether the debug port is listening
        :return:check
        """
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        result = sock.connect_ex((self.ip, self.port))
        if result == 0:
            check = True
        else:
            check = False
        sock.close()
        return check

    def driver_setup(self):
        chrome_driver = webdrivermanager.ChromeDriverManager()
        try:
            temp_driver = webdriver.Chrome(executable_path=chrome_driver.get_driver_filename(),
                                           options=self.debug_chrome())
            return temp_driver

        except:
            chrome_driver.download_and_install(chrome_driver.get_latest_version())
            time.sleep(3)
            print("!!!!!!!!!!!!!!!!")
            time.sleep(3)
            self.driver = webdriver.Chrome(executable_path=chrome_driver.get_driver_filename(),
                                           options=self.debug_chrome())
            return temp_driver
