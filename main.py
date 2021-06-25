from selenium import webdriver
import time
import testraw
import netsuite_clean_all_case
import win32api, win32con

test1 = netsuite_clean_all_case.CleanAllCase()
test1.__init__()
var = ["To Base Brands CC",
       "Amware Logistics Unknown To Unknown",
       "Almo Unknown To Unknown",
       "Home Depot Canada Unknown To Unknown"
       ]
for search_key in var:
    test1.change_criteria("contains", search_key)
    test1.clean_all_case()

test1.change_criteria("is not empty", "Hello")
win32api.MessageBox(0, "No more noise in queue. :)", "Cleaning Done", win32con.MB_OK)
# log = "1050935919"
# test1.test_psdautologin(log)
# url = 'http://psdtool.tc.net/psdTool'


