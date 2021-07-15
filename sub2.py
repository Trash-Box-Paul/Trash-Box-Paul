import netsuite_clean_all_case
import win32api, win32con

test1 = netsuite_clean_all_case.CleanAllCase()
# test1.take_task()
test1.resend_all_case()
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
