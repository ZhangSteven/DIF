"""
Test the read_holding() method from open_holding.py

"""

import unittest2
import datetime
from xlrd import open_workbook
from DIF.utility import get_current_path
from DIF.open_holding import read_holding

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



    def test_normal(self):
        """
        Test the four values read from the summary sheet.
        """

        # here we do not use try/except claues becuase we provide
        # samples the exists, we will have another set of tests 
        # for error conditions.
        filename = get_current_path() + '\\samples\\holdings_sample.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')

        read_holding(ws, self.port_values)

        p = self.port_values
        self.assertAlmostEqual(1,1)