import unittest
from TrainingReportGenerator import *


class test_tuple1(unittest.TestCase):

    def test_general_use(self):
        test_tup = tuple1("Test", 2)
        self.assertTrue(test_tup.get_text() == "Test" and test_tup.get_hours() == 2)
        test_tup.set_hours(8)
        self.assertEqual(test_tup.get_hours(), 8)

    def test_max_text_length(self):
        test_tup = tuple1("This Test tests for the max length of text in tuple1", 2)
        self.assertEqual(test_tup.get_text(), "This Test tests for the max length of text in tupl")

    def test_min_hours(self):
        test_tup = tuple1("Test", 0)
        self.assertGreater(test_tup.get_hours(), 0)
        test_tup.set_hours(0)
        self.assertGreater(test_tup.get_hours(), 0)

    def test_max_hours(self):
        test_tup = tuple1("Test", 41)
        self.assertLessEqual(test_tup.get_hours(), 40)
        test_tup.set_hours(41)
        self.assertLessEqual(test_tup.get_hours(), 40)


if __name__ == '__main__':
    unittest.main()
