import unittest2
import keas_methods as methods

class keasPositiveTestCases(unittest2.TestCase):

    @staticmethod # signal to Unittest framework that this is a function inside the class (vs. @classmethod)
    def test_create_new_user(): # test_ in the name is mandatory
        methods.setup()
        methods.log_in(methods.keas_username,methods.keas_password)
        methods.add_new_user("C:/Users/kando/Desktop/YZTek/data_keas.xlsx", "Sheet1")
        methods.log_out()
        methods.teardown()