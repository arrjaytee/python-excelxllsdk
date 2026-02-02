from unittest import TestCase

from ExcelXLLSDK.com.application import find_Application

class ComTestCase(TestCase):
    def test_find_Application(self):
        with self.assertRaises(RuntimeError):
            find_Application(0)

