"""
Test the read_cash() method from open_cash.py

Note that the class will be instantiated again for each test method run,
and the setup() and teardown() methods are called each time a test method run.
"""

import unittest2
import datetime
from xlrd import open_workbook
from DIF.utility import get_current_path
from DIF.open_cash import read_cash

class TestCash(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
         # return the path to test excel file
        self.filename = get_current_path() + '\\samples\\cash_sample1.xls'
        super(TestCash, self).__init__(*args, **kwargs)

    def setUp(self):
        """
            Run before a test function
        """
        wb = open_workbook(filename=self.filename)

        self.port_values = {}

        # find sheets that contain cash
        sheet_names = wb.sheet_names()
        for sn in sheet_names:
            if len(sn) > 4 and sn[-4:] == '-BOC':
                # print('read from sheet {0}'.format(sn))
                ws = wb.sheet_by_name(sn)
                read_cash(ws, self.port_values)



    def tearDown(self):
        """
            Run after a test finishes
        """
        pass



    def test_hkd(self):
        """
        Test the HKD cash account.
        """
        self.assertTrue('cash_accounts' in self.port_values)

        cash_accounts = self.port_values['cash_accounts']
        self.assertEqual(len(cash_accounts), 4) # read in 4 sheets

        cash_account = cash_accounts[1]     # the first sheet is "HKD"
      
        self.assertEqual(cash_account['currency'], 'HKD')
        self.assertEqual(cash_account['account_num'], '012-875-0-053124-1')
        self.assertEqual(cash_account['account_type'], 'Current Account')
        self.assertEqual(cash_account['bank'], 'Bank of China (Hong Kong) Ltd')
        self.assertEqual(cash_account['date'], datetime.datetime(2015,12,10))
        self.assertAlmostEqual(cash_account['balance'], 6536572.95)
        self.assertEqual(cash_account['fx_rate'], 1.0)
        self.assertAlmostEqual(cash_account['hkd_equivalent'], 6536572.95)

        

    def test_usd(self):
        """
        Test the USD cash account.
        """
        cash_accounts = self.port_values['cash_accounts']
        cash_account = cash_accounts[2]     # the first sheet is "USD"
      
        self.assertEqual(cash_account['currency'], 'USD')
        self.assertEqual(cash_account['account_num'], '012-875-0-804911-9')
        self.assertEqual(cash_account['account_type'], 'Current Account')
        self.assertEqual(cash_account['bank'], 'Bank of China (Hong Kong) Ltd')
        self.assertEqual(cash_account['date'], datetime.datetime(2015,12,10))
        self.assertAlmostEqual(cash_account['balance'], 8298021.81)
        self.assertAlmostEqual(cash_account['fx_rate'], 7.7502)
        self.assertAlmostEqual(cash_account['hkd_equivalent'], 64311328.63)



    def test_cny(self):

        """
        Test the CNY cash account.
        """
        cash_accounts = self.port_values['cash_accounts']
        cash_account = cash_accounts[3]     # the first sheet is "CNY"
      
        self.assertEqual(cash_account['currency'], 'CNY')
        self.assertEqual(cash_account['account_num'], '012-875-0-603962-0')
        self.assertEqual(cash_account['account_type'], 'Current Account')
        self.assertEqual(cash_account['bank'], 'Bank of China (Hong Kong) Ltd')
        self.assertEqual(cash_account['date'], datetime.datetime(2015,12,10))
        self.assertAlmostEqual(cash_account['balance'], 386920)
        self.assertAlmostEqual(cash_account['fx_rate'], 1.2037)
        self.assertAlmostEqual(cash_account['hkd_equivalent'], 465735.604)



    def test_sgd(self):

        """
        Test the SGD cash account.
        """
        cash_accounts = self.port_values['cash_accounts']
        cash_account = cash_accounts[4]     # the first sheet is "SGD"
      
        self.assertEqual(cash_account['currency'], 'SGD')
        self.assertEqual(cash_account['account_num'], '012-875-0-604032-3')
        self.assertEqual(cash_account['account_type'], 'Current Account')
        self.assertEqual(cash_account['bank'], 'Bank of China (Hong Kong) Ltd')
        self.assertEqual(cash_account['date'], datetime.datetime(2015,12,10))
        self.assertAlmostEqual(cash_account['balance'], 0)
        self.assertAlmostEqual(cash_account['fx_rate'], 5.5201)
        self.assertAlmostEqual(cash_account['hkd_equivalent'], 0)
