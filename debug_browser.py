import os
from selenium.webdriver.chrome.options import Options
import socket


# Open the browser on port 9222, and save the information to C:/testfile
class DebugBrowser:
    def __init__(self):
        self.ip = '127.0.0.1'
        self.port = 9000
        self.user_file = 'C:/test'
        self.chrome_option = Options()
        self.chrome_address = 'C:\Program Files\Google\Chrome\Application'

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
