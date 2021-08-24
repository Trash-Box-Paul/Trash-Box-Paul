import netsuite_clean_all_case
import win32api, win32con
import excel_test

test1 = excel_test.TakeTasks()
# test1.cloud_ftp("Fudgeamentals - Test Profile")
# test1.resend_all_case()
# email_list=[]
# email_list.append("paul.wu@truecommerce.com")
# test1.send_initial_emails(email_list, "Volkswagen.De")
test1.send_all_tps()
# import unittest
#
#
# class MyTestCase(unittest.TestCase):
#     def test_something(self):
#         self.assertEqual(True, False)
#
#
# if __name__ == '__main__':
#     unittest.main()
