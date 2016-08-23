"""
Test methods open_expense.py

"""

import unittest2
import datetime
from xlrd import open_workbook
from DIF.utility import get_current_path
from DIF.open_summary import read_date

class TestExpense(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestExpense, self).__init__(*args, **kwargs)

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



    def test_read_date(self):
        """
        Test the read_date() function.
        """
        filename = get_current_path() + '\\samples\\expense_sample.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Expense Report')

        d = read_date(ws, 5, 1)
        self.assertEqual(d, datetime.datetime(2015,12,10))



    def test_expense_fields(self):

        filename = get_current_path() + '\\samples\\expense_sample.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Expense Report')

        