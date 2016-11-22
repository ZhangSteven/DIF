# coding=utf-8
"""
Test methods open_expense.py

"""

import unittest2
import datetime
from xlrd import open_workbook
from DIF.utility import get_current_path
from DIF.open_dif import InconsistentExpenseDate, InvalidTickerFormat, \
                            convert_to_BLP_ticker, open_dif

                            

class TestDif(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestDif, self).__init__(*args, **kwargs)

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



    def test_validate_expense_date(self):
        """
        
        """
        filename = get_current_path() + '\\samples\\expense_error1.xls'
        port_values = {}

        with self.assertRaises(InconsistentExpenseDate):
            open_dif(filename, port_values)



    def test_convert_to_BLP_ticker(self):
        """
        
        """
        ticker = 'H88888'   # too long
        with self.assertRaises(InvalidTickerFormat):
            convert_to_BLP_ticker(ticker)

        ticker = 'N123'     # too short
        with self.assertRaises(InvalidTickerFormat):
            convert_to_BLP_ticker(ticker)

        ticker = 'A1234'    # wrong start character
        with self.assertRaises(InvalidTickerFormat):
            convert_to_BLP_ticker(ticker)

        ticker = 'N0000'    # all zeros code
        with self.assertRaises(InvalidTickerFormat):
            convert_to_BLP_ticker(ticker)

        self.assertEqual(convert_to_BLP_ticker('H0939'), '939 HK')
        self.assertEqual(convert_to_BLP_ticker('N0011'), '11 HK')
        self.assertEqual(convert_to_BLP_ticker('N2388'), '2388 HK')
        self.assertEqual(convert_to_BLP_ticker('H1186'), '1186 HK')