"""
Test the read_cash() method from open_cash.py

Note that the class will be instantiated again for each test method run,
and the setup() and teardown() methods are called each time a test method run.
"""

import unittest2
import datetime
from xlrd import open_workbook
from DIF.utility import get_current_path
from DIF.open_summary import read_portfolio_summary, read_date, find_cell_string

class TestCash(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestCash, self).__init__(*args, **kwargs)

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


    def test_find_cell_string(self):
        filename = get_current_path() + '\\samples\\summary_sample.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Sum.')

        n = find_cell_string(ws, 0, 1, 'Valuation Period :')
        self.assertEqual(n, 5)



    def test_read_date(self):
        """
        Test the read_date() function.
        """
        filename = get_current_path() + '\\samples\\summary_sample.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Sum.')
        port_values = {}
        d = read_date(ws, 5, 1)
        self.assertEqual(d, datetime.datetime(2015,12,10))



    def test_normal(self):
        """
        Test the four values read from the summary sheet.
        """

        # here we do not use try/except claues becuase we provide
        # samples the exists, we will have another set of tests 
        # for error conditions.
        filename = get_current_path() + '\\samples\\summary_sample.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Sum.')

        read_portfolio_summary(ws, self.port_values)

        p = self.port_values
        self.assertAlmostEqual(p['unit_price'], 10.6789)
        self.assertAlmostEqual(p['nav'], 459813450.8976)
        self.assertAlmostEqual(p['number_of_units'], 42663321.3938)
        self.assertEqual(p['date'], datetime.datetime(2015,12,10))



    def test_error2(self):
        """
        Test the four values read from the summary sheet.
        """

        # here we do not use try/except claues becuase we provide
        # samples the exists, we will have another set of tests 
        # for error conditions.
        filename = get_current_path() + '\\samples\\summary_error2.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Sum.')

        with self.assertRaises(TypeError):
            read_date(ws, 5, 1)

        # try:
        #     read_portfolio_summary(ws, self.port_values)
        # except ValueError as e:
        #     # expected to have this error with the message 'date'
        #     self.assertEqual(str(e), 'date')
        # except Exception:
        #     # if error other ValueError occurs, then something must
        #     # be wrong
        #     self.assertEqual('something is worng', 0)
        # else:
        #     # if error does not occur, then something must be wrong
        #     # because the cell has a wrong value. (B6)
        #     self.assertEqual('something is worng', 1)



    def test_error3(self):
        """
        Test the four values read from the summary sheet.
        """

        # here we do not use try/except claues becuase we provide
        # samples the exists, we will have another set of tests 
        # for error conditions.
        filename = get_current_path() + '\\samples\\summary_error3.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Sum.')

        try:
            read_portfolio_summary(ws, self.port_values)
        except ValueError as e:
            # expected to have this error with the message 'date'
            self.assertEqual(str(e), 'number_of_units')
        except Exception:
            # if error other ValueError occurs, then something must
            # be wrong
            self.assertEqual('something is worng', 0)
        else:
            # if error does not occur, then something must be wrong
            # because the cell has a wrong value. (C47)
            self.assertEqual('something is worng', 1)



    def test_error4(self):
        """
        Test the four values read from the summary sheet.
        """

        # here we do not use try/except claues becuase we provide
        # samples the exists, we will have another set of tests 
        # for error conditions.
        filename = get_current_path() + '\\samples\\summary_error4.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Sum.')

        try:
            read_portfolio_summary(ws, self.port_values)
        except ValueError as e:
            # expected to have this error with the message 'date'
            self.assertEqual(str(e), 'nav')
        except Exception:
            # if error other ValueError occurs, then something must
            # be wrong
            self.assertEqual('something is worng', 0)
        else:
            # if error does not occur, then something must be wrong
            # because the nav is 0. (J53)
            self.assertEqual('something is worng', 1)