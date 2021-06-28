import netsuite_read_task

test1 = netsuite_read_task.TestLogin()
log = "1050935919"
test1.setup_method()
test1.test_login()
# test1.teardown_method()