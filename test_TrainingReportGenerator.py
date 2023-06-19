import unittest
from TrainingReportGenerator import *


class test_tuple1(unittest.TestCase):

    def test_init_general_use(self):
        test_tup = tuple1("Test", 2)
        self.assertTrue(test_tup.get_text() == "Test" and test_tup.get_hours() == 2)

    def test_init_max_text_length(self):
        test_tup = tuple1("This Test tests for the max length of text in tuple1", 2)
        self.assertEqual(test_tup.get_text(), "This Test tests for the max length of text in tupl")

    def test_init_min_hours(self):
        test_tup = tuple1("Test", 0)
        self.assertGreater(test_tup.get_hours(), 0)

    def test_init_max_hours(self):
        test_tup = tuple1("Test", 41)
        self.assertLessEqual(test_tup.get_hours(), 40)


if __name__ == '__main__':
    unittest.main()
