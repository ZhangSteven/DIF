"""
Test the read_holding() method from open_holding.py

"""

import unittest2
import datetime
from xlrd import open_workbook
from DIF.utility import get_current_path
from DIF.open_holding import read_holding, read_bond_fields

class TestHolding(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestHolding, self).__init__(*args, **kwargs)

    def setUp(self):
        """
            Run before a test function
        """
        self.port_values = {}



    def tearDown(self):
        """
            Run after a test finishes
        """
        pass



    def test_read_bond_fields(self):
        """
        Test the read_bond_fields() method.
        """

        filename = get_current_path() + '\\samples\\holdings_sample2.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 62    # where the bond section starts

        fields, n = read_bond_fields(ws, row)
        self.assertEqual(n, 4)
        self.assertEqual(len(fields), 15)

        f = ''
        for s in fields:
            f = f + s + ', '

        # check the fields are read correctly
        self.assertEqual(f, 
            'par_amount, is_listed, listed_location, fx_trade_date, coupon_rate, coupon_start_date, maturity_date, average_cost, amortized_cost, book_cost, interest_bought, amortized_value, accrued_interest, amortized_gain_loss, fx_gain_loss, ')