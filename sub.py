import netsuite_clean_all_case

test1 = netsuite_clean_all_case.CleanAllCase()
test1.__init__()
test1.change_criteria("is not empty", "Hello")
