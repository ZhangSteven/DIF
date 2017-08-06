"""
Test the read_cash() method from open_cash.py

Note that the class will be instantiated again for each test method run,
and the setup() and teardown() methods are called each time a test method run.
"""

import unittest2
import datetime, os
from xlrd import open_workbook
from DIF.utility import get_current_path
from DIF.open_cash import read_cash
from DIF.open_bal import consolidate_cash



class TestCashBAL(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestCashBAL, self).__init__(*args, **kwargs)



    def test_consolidate_cash(self):
        filename = os.path.join(get_current_path(), 'samples', 'CLM BAL 2017-07-27.xls')
        wb = open_workbook(filename=filename)
        port_values = {}
        sheet_names = wb.sheet_names()
        for sn in sheet_names:
            if sn.startswith('CA') or sn.startswith('SA'):
                ws = wb.sheet_by_name(sn)
                read_cash(ws, port_values)

        self.assertEqual(len(port_values['cash_accounts']), 9) # read in 9 sheets
        consolidate_cash(port_values)
        self.assertEqual(len(port_values['cash_accounts']), 7) # combined into 7 accounts
                                                               # by bank/currency
        cash_accounts = port_values['cash_accounts']
        self.verify_citi_hkd(extract_accounts(cash_accounts, 'HKD', 'Citibank N.A.'))
        self.verify_sh_hkd(extract_accounts(cash_accounts, 'HKD', 'Industrial and Commercial Bank of China (Macau) Ltd'))
        self.verify_sh_mop(extract_accounts(cash_accounts, 'MOP', 'Industrial and Commercial Bank of China (Macau) Ltd'))
        self.verify_sh_usd(extract_accounts(cash_accounts, 'USD', 'Industrial and Commercial Bank of China (Macau) Ltd'))
        self.verify_sh_cny(extract_accounts(cash_accounts, 'CNY', 'Industrial and Commercial Bank of China (Macau) Ltd'))
        self.verify_gf_hkd(extract_accounts(cash_accounts, 'HKD', 'China Guangfa Bank Co., Ltd Macau Branch'))
        self.verify_luso_hkd(extract_accounts(cash_accounts, 'HKD', 'Luso International Banking Ltd.'))



    def verify_citi_hkd(self, cash_account):
        """
        Verify the HKD balance from CitiBank.
        """
        self.assertEqual(cash_account['account_num'], '006-391-17836395')
        self.assertEqual(cash_account['account_type'], 'Saving & Checking Account')
        self.assertEqual(cash_account['date'], datetime.datetime(2017,7,27))
        self.assertAlmostEqual(cash_account['balance'], 7278.97)
        self.assertAlmostEqual(cash_account['fx_rate'], 1.03004645509512)
        self.assertAlmostEqual(cash_account['local_currency_equivalent'], 7497.68)



    def verify_sh_hkd(self, cash_account):
        """
        Verify the HKD balance from Industrial and Commercial Bank of China (Macau) Ltd
        """
        self.assertEqual(cash_account['date'], datetime.datetime(2017,7,27))
        self.assertAlmostEqual(cash_account['balance'], 59953.17)
        self.assertAlmostEqual(cash_account['fx_rate'], 1.03004645509512)
        self.assertAlmostEqual(cash_account['local_currency_equivalent'], 61754.55)



    def verify_sh_mop(self, cash_account):
        """
        Verify the MOP balance from Industrial and Commercial Bank of China (Macau) Ltd
        """
        self.assertEqual(cash_account['date'], datetime.datetime(2017,7,27))
        self.assertAlmostEqual(cash_account['balance'], 2678374.76)
        self.assertAlmostEqual(cash_account['fx_rate'], 1.0)
        self.assertAlmostEqual(cash_account['local_currency_equivalent'], 2678374.76)



    def verify_sh_usd(self, cash_account):
        """
        Verify the USD balance from Industrial and Commercial Bank of China (Macau) Ltd
        """
        self.assertEqual(cash_account['date'], datetime.datetime(2017,7,27))
        self.assertAlmostEqual(cash_account['balance'], 0)
        self.assertAlmostEqual(cash_account['fx_rate'], 8.04373577248334)
        self.assertAlmostEqual(cash_account['local_currency_equivalent'], 0)



    def verify_sh_cny(self, cash_account):
        """
        Verify the USD balance from Industrial and Commercial Bank of China (Macau) Ltd
        """
        self.assertEqual(cash_account['date'], datetime.datetime(2017,7,27))
        self.assertAlmostEqual(cash_account['balance'], 3801.94)
        self.assertAlmostEqual(cash_account['fx_rate'], 1.19289679964566)
        self.assertAlmostEqual(cash_account['local_currency_equivalent'], 4535.32)



    def verify_gf_hkd(self, cash_account):
        """
        Verify the HKD balance from GuangFa bank.
        """
        self.assertEqual(cash_account['date'], datetime.datetime(2017,7,27))
        self.assertAlmostEqual(cash_account['balance'], 72.88)
        self.assertAlmostEqual(cash_account['fx_rate'], 1.03004645509512)
        self.assertAlmostEqual(cash_account['local_currency_equivalent'], 75.07)



    def verify_luso_hkd(self, cash_account):
        """
        Verify the HKD balance from GuangFa bank.
        """
        self.assertEqual(cash_account['date'], datetime.datetime(2017,7,27))
        self.assertAlmostEqual(cash_account['balance'], 635.75)
        self.assertAlmostEqual(cash_account['fx_rate'], 1.03004645509512)
        self.assertAlmostEqual(cash_account['local_currency_equivalent'], 654.85)



def extract_accounts(cash_accounts, currency, bank):
    """
    Extract cash accounts by currency and bank.
    """
    for cash_account in cash_accounts:
        if cash_account['currency'] == currency and cash_account['bank'] == bank:
            return cash_account