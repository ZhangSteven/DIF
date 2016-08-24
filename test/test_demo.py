# coding=utf-8
"""
Test demo.py

"""

import unittest2
import datetime
from xlrd import open_workbook
from DIF.demo import open_dif
from DIF.utility import get_current_path



class TestDemo(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestDemo, self).__init__(*args, **kwargs)

    def setUp(self):
        """
            Run before a test function
        """
        pass

    def tearDown(self):
        """
            Run after a test finishes
        """
        pass



    def test_compile(self):
        """
        Call the open_dif() method, make sure there is no syntax errors.
        """
        filename = get_current_path() + '\\samples\\sample_DIF_20151210.xls'
        port_values = {}
        open_dif(filename, port_values)

        self.assertEqual(1,1)   # trivial