import netsuite_clean_all_case
import win32api, win32con
import excel_test


test1 = netsuite_clean_all_case.CleanAllCase()
test1.cloud_ftp("Fudgeamentals - Test Profile")