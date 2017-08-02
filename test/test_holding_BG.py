"""
Test the read_holding() method from open_holding.py

"""

import unittest2
import os
from datetime import datetime
from xlrd import open_workbook
from DIF.utility import get_current_path
from DIF.open_holding import read_holding, read_bond_fields, read_currency, \
                            get_datemode, read_equity_fields, \
                            read_sub_section, read_section



class TestHoldingBG(unittest2.TestCase):
    """
    Test files for the Macau fund that uses almost the same format as DIF: 
    1. Guaranteed open fund (GNT)
    2. Balanced open fund (BAL)
    """
    def __init__(self, *args, **kwargs):
        super(TestHoldingBG, self).__init__(*args, **kwargs)

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



    def test_read_bond_fields_HTM(self):
        """
        Test the read_bond_fields() method using holdings_sample2.xls,
        containing only one bond section with HTM bonds.
        """
        filename = os.path.join(get_current_path(), 'samples', 'CLM BAL 2017-07-27.xls')
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 55    # the bond section starts at A56

        fields, n = read_bond_fields(ws, row)
        self.assertEqual(n, 4)
        self.assertEqual(len(fields), 15)

        f = ''
        for s in fields:
            f = f + s + ', '

        # check the fields are read correctly
        self.assertEqual(f, 
            'par_amount, is_listed, listed_location, fx_on_trade_day, coupon_rate, coupon_start_date, maturity_date, average_cost, amortized_cost, book_cost, interest_bought, amortized_value, accrued_interest, amortized_gain_loss, fx_gain_loss, ')



    def test_read_equity_fields_listed(self):
        """
        Test the read_equity_fields() method with listed eqiuty.
        """

        filename = os.path.join(get_current_path(), 'samples', 'CLM BAL 2017-07-27.xls')
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 129    # the equity section starts at A130

        fields, n = read_equity_fields(ws, row)
        self.assertEqual(n, 3)
        self.assertEqual(len(fields), 15)

        f = ''
        for s in fields:
            f = f + s + ', '

        # check the fields are read correctly
        self.assertEqual(f, 
            'number_of_shares, is_listed, listed_location, fx_on_trade_day, empty_field, last_trade_date, empty_field, average_cost, price, book_cost, empty_field, market_value, empty_field, market_gain_loss, fx_gain_loss, ')



    def test_read_currency(self):
        msg = 'V. Debt Securities (Held-to-Maturity) - US$  (持到期債務票據- 美元)'
        self.assertEqual(read_currency(msg), 'USD')

        msg = 'VI. Debt Securities (Avaliable for sales) - USD  (可供出售債務票據- 美元)'
        self.assertEqual(read_currency(msg), 'USD')
        
        msg = 'VIII.  Equities -USD  (股票-美元)'
        self.assertEqual(read_currency(msg), 'USD')



    def test_read_section_bond(self):

        filename = os.path.join(get_current_path(), 'samples', 'CLM BAL 2017-07-27.xls')
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 55    # USD HTM bond section starts at A56

        cell_value = 'V. Debt Securities (Held-to-Maturity) - US$  (持到期債務票據- 美元)'
        currency = read_currency(cell_value)
        fields, n = read_bond_fields(ws, row)   # read the bond
        row = row + n                           # field names
        port_values = {}

        n = read_section(ws, row, fields, 'bond', currency, port_values)
        self.validate_bond_HTM(port_values['bond'])
        


    def test_read_section_bond2(self):

        filename = os.path.join(get_current_path(), 'samples', 'CLM BAL 2017-07-27.xls')
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 111    # USD AFS bond section starts at A112

        cell_value = 'VI. Debt Securities (Avaliable for sales) - USD  (可供出售債務票據- 美元)'
        currency = read_currency(cell_value)
        fields, n = read_bond_fields(ws, row)   # read the bond
        row = row + n                           # field names
        port_values = {}

        n = read_section(ws, row, fields, 'bond', currency, port_values)
        self.validate_bond_trading(port_values['bond'])



    def test_read_section_equity(self):

        filename = os.path.join(get_current_path(), 'samples', 'CLM BAL 2017-07-27.xls')
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 129    # HKD Equity section starts at A130

        cell_value = 'VII. Equities - HKD  (股票-港幣)'
        currency = read_currency(cell_value)
        fields, n = read_equity_fields(ws, row)
        row = row + n
        port_values = {}

        n = read_section(ws, row, fields, 'equity', currency, port_values)
        self.validate_equity(port_values['equity'])



    def test_read_section_equity2(self):

        filename = os.path.join(get_current_path(), 'samples', 'CLM BAL 2017-07-27.xls')
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 154    # USD Equity section starts at A155

        cell_value = 'VIII.  Equities -USD  (股票-美元)'
        currency = read_currency(cell_value)
        fields, n = read_equity_fields(ws, row)
        row = row + n
        port_values = {}

        n = read_section(ws, row, fields, 'equity', currency, port_values)
        self.validate_equity2(port_values['equity'])




    # def test_read_holding(self):
    #     """
    #     Read the CLM BAL 2017-07-27.xls, we get below holdings:

    #     Bond        
    #                     total   zero holding
    #     USD (HTM)       3       1
    #     HKD (HTM)       4       4
    #     USD (Trading)   1       0
    #     HKD (Trading)   1       1
    #     total           9       6
                
    #     Equity      
    #                     total   zero holding
    #     HKD             13      1
    #     USD             13      11       
    #     total           26      12
    #     """
    #     filename = os.path.join(get_current_path(), 'samples', 'CLM BAL 2017-07-27.xls')
    #     wb = open_workbook(filename=filename)
    #     ws = wb.sheet_by_name('Portfolio Val.')

    #     port_values = {}
    #     read_holding(ws, port_values)

    #     bond_holding = port_values['bond']
    #     self.assertEqual(len(bond_holding), 9)
    #     self.assertEqual(self.count_zero_holding_bond(bond_holding), 6)

    #     equity_holding = port_values['equity']
    #     self.assertEqual(len(equity_holding), 26)
    #     self.assertEqual(self.count_zero_holding_equity(equity_holding), 12)

        # bond_holding_HTM_USD = self.extract_bond_holding(bond_holding, 'USD', 'HTM')
        # self.validate_bond_HTM(bond_holding_HTM_USD)

        # bond_holding_Trading_USD = self.extract_bond_holding(bond_holding, 'USD', 'Trading')
        # self.validate_bond_trading(bond_holding_Trading_USD)

        # bond_holding_HTM_SGD = self.extract_bond_holding(bond_holding, 'SGD', 'HTM')
        # self.assertEqual(len(bond_holding_HTM_SGD), 10)
        # self.assertEqual(self.count_zero_holding_bond(bond_holding_HTM_SGD), 10)

        # bond_holding_Trading_SGD = self.extract_bond_holding(bond_holding, 'SGD', 'Trading')
        # self.assertEqual(len(bond_holding_Trading_SGD), 1)
        # self.assertEqual(self.count_zero_holding_bond(bond_holding_Trading_SGD), 1)

        # bond_holding_HTM_CNY = self.extract_bond_holding(bond_holding, 'CNY', 'HTM')
        # self.assertEqual(len(bond_holding_HTM_CNY), 2)
        # self.assertEqual(self.count_zero_holding_bond(bond_holding_HTM_CNY), 0)

        # bond_holding_Trading_CNY = self.extract_bond_holding(bond_holding, 'CNY', 'Trading')
        # self.assertEqual(len(bond_holding_Trading_CNY), 15)
        # self.assertEqual(self.count_zero_holding_bond(bond_holding_Trading_CNY), 13)

        # equity_holding_listed = self.extract_equity_holding(equity_holding, 'HKD')
        # self.validate_listed_equity(equity_holding_listed)

        # equity_holding_preferred_shares = self.extract_equity_holding(equity_holding, 'USD')
        # self.validate_preferred_shares(equity_holding_preferred_shares)



    def extract_bond_holding(self, bond_holding, currency, accounting_treatment):
        """
        Extract bond holding given its currency and accounting treatment.
        """
        holding = []
        for bond in bond_holding:
            if bond['currency'] == currency and \
                bond['accounting_treatment'] == accounting_treatment:
                holding.append(bond)

        return holding



    def extract_equity_holding(self, equity_holding, currency):
        """
        Extract equity holding given its currency.
        """
        holding = []
        for equity in equity_holding:
            if equity['currency'] == currency:
                holding.append(equity)

        return holding



    def count_zero_holding_bond(self, bond_holding):
        """
        Count how many bonds in the holding has zero amount (par_amount).
        """
        empty_bond = 0
        for bond in bond_holding:
            if bond['par_amount'] == 0:
                empty_bond = empty_bond + 1

        return empty_bond



    def count_zero_holding_equity(self, equity_holding):
        """
        Count how many equity in the holding has zero amount (number_of_shares).
        """
        empty_equity = 0
        for equity in equity_holding:
            if equity['number_of_shares'] == 0:
                empty_equity = empty_equity + 1

        return empty_equity



    def validate_bond_trading(self, bond_holding):
        """
        Validate the USD bond AFS section from CLM BAL 2017-07-27.xls
        """
        self.assertEqual(len(bond_holding), 1) # should have 1 position
        self.assertEqual(self.count_zero_holding_bond(bond_holding), 0)

        i = 0
        for bond in bond_holding:
            i = i + 1

            if (i == 1):    # the first bond
                self.assertEqual(bond['isin'], 'XS1389124774')
                self.assertEqual(bond['name'], '(XS1389124774) DEMETER (SWISS RE LTD) 6.05%')
                self.assertEqual(bond['currency'], 'USD')
                self.assertEqual(bond['accounting_treatment'], 'Trading')
                self.assertAlmostEqual(bond['par_amount'], 3700000)
                self.assertEqual(bond['is_listed'], 'TBC')
                self.assertEqual(bond['listed_location'], 'TBC')
                self.assertAlmostEqual(bond['fx_on_trade_day'], 7.98890035806361)
                self.assertAlmostEqual(bond['coupon_rate'], 6.05/100)
                self.assertEqual(bond['coupon_start_date'], datetime(2017,2,15))
                self.assertEqual(bond['maturity_date'], datetime(2056,2,15))
                self.assertAlmostEqual(bond['average_cost'], 100)
                self.assertAlmostEqual(bond['price'], 107.951)
                self.assertAlmostEqual(bond['book_cost'], 3700000)
                self.assertAlmostEqual(bond['interest_bought'], 0)
                self.assertAlmostEqual(bond['market_value'], 3994187)
                self.assertAlmostEqual(bond['accrued_interest'], 101354.31)
                self.assertAlmostEqual(bond['market_gain_loss'], 294187)
                self.assertAlmostEqual(bond['fx_gain_loss'], 202891.03)



    def validate_bond_HTM(self, bond_holding):
        """
        Validate the USD bond HTM section from CLM BAL 2017-07-27.xls
        """
        self.assertEqual(len(bond_holding), 3) # should have 3 positions
        self.assertEqual(self.count_zero_holding_bond(bond_holding), 1)

        i = 0
        for bond in bond_holding:
            i = i + 1

            if (i == 1):    # the first bond
                self.assertEqual(bond['isin'], 'XS0508012092')
                self.assertEqual(bond['name'], '(XS0508012092) China Overseas Finance 5.5%')
                self.assertEqual(bond['currency'], 'USD')
                self.assertEqual(bond['accounting_treatment'], 'HTM')
                self.assertAlmostEqual(bond['par_amount'], 400000)
                self.assertEqual(bond['is_listed'], 'TBC')
                self.assertEqual(bond['listed_location'], 'TBC')
                self.assertAlmostEqual(bond['fx_on_trade_day'], 8.0024)
                self.assertAlmostEqual(bond['coupon_rate'], 5.5/100)
                self.assertEqual(bond['coupon_start_date'], datetime(2017,5,10))
                self.assertEqual(bond['maturity_date'], datetime(2020,11,10))
                self.assertAlmostEqual(bond['average_cost'], 95.75)
                self.assertAlmostEqual(bond['amortized_cost'], 98.25302)
                self.assertAlmostEqual(bond['book_cost'], 383000)
                self.assertAlmostEqual(bond['interest_bought'], 0)
                self.assertAlmostEqual(bond['amortized_value'], 393012.08)
                self.assertAlmostEqual(bond['accrued_interest'], 4766.67)
                self.assertAlmostEqual(bond['amortized_gain_loss'], 10012.08)
                self.assertAlmostEqual(bond['fx_gain_loss'], 15831.6)

            if (i == 2):    # the second bond
                self.assertEqual(bond['isin'], 'USG59606AA46')
                self.assertEqual(bond['name'], '(USG59606AA46) Mega Advance Investment 5%')
                self.assertEqual(bond['currency'], 'USD')
                self.assertEqual(bond['accounting_treatment'], 'HTM')
                self.assertAlmostEqual(bond['par_amount'], 400000)
                self.assertEqual(bond['is_listed'], 'TBC')
                self.assertEqual(bond['listed_location'], 'TBC')
                self.assertAlmostEqual(bond['fx_on_trade_day'], 8.0024)
                self.assertAlmostEqual(bond['coupon_rate'], 5.0/100)
                self.assertEqual(bond['coupon_start_date'], datetime(2017,5,12))
                self.assertEqual(bond['maturity_date'], datetime(2021,5,12))
                self.assertAlmostEqual(bond['average_cost'], 96.25)
                self.assertAlmostEqual(bond['amortized_cost'], 98.2888)
                self.assertAlmostEqual(bond['book_cost'], 385000)
                self.assertAlmostEqual(bond['interest_bought'], 0)
                self.assertAlmostEqual(bond['amortized_value'], 393155.2)
                self.assertAlmostEqual(bond['accrued_interest'], 4222.22)
                self.assertAlmostEqual(bond['amortized_gain_loss'], 8155.2)
                self.assertAlmostEqual(bond['fx_gain_loss'], 15914.28)

            if (i == 3):    # the last bond, an empty position
                self.assertEqual(bond['isin'], 'USG2108YAA31')
                self.assertEqual(bond['name'], '(USG2108YAA31) China Resource Land Ltd 4.625%')
                self.assertEqual(bond['currency'], 'USD')
                self.assertEqual(bond['accounting_treatment'], 'HTM')
                self.assertAlmostEqual(bond['par_amount'], 0)
                self.assertEqual(len(bond), 5)  # should have no more fields



    def validate_equity(self, equity_holding):
        """
        Validate the HKD equity section from CLM BAL 2017-07-27.xls
        """
        self.assertEqual(len(equity_holding), 13) # should have 13 positions
        self.assertEqual(self.count_zero_holding_equity(equity_holding), 1)

        i = 0
        for equity in equity_holding:
            i = i + 1

            if (i == 1):    # the first equity
                self.assertEqual(equity['ticker'], 'N0522')
                self.assertEqual(equity['name'], '(N0522) ASM Pacific Technology Ltd.')
                self.assertEqual(equity['currency'], 'HKD')
                self.assertEqual(equity['accounting_treatment'], 'Trading')
                self.assertAlmostEqual(equity['number_of_shares'], 4100)
                self.assertEqual(equity['is_listed'], 'Y')
                self.assertEqual(equity['listed_location'], 'Hong Kong')
                self.assertAlmostEqual(equity['fx_on_trade_day'], 1.03006767544628)
                self.assertEqual(equity['last_trade_date'], datetime(2017,7,27))
                self.assertAlmostEqual(equity['average_cost'], 121.822512195122)
                self.assertAlmostEqual(equity['price'], 101.3)
                self.assertAlmostEqual(equity['book_cost'], 499472.3)
                self.assertAlmostEqual(equity['market_value'], 415330)
                self.assertAlmostEqual(equity['market_gain_loss'], -84142.3)
                self.assertAlmostEqual(equity['fx_gain_loss'], -10.6)
                self.assertEqual(len(equity), 15)

            if (i == 12):    # the last non-zero holding
                self.assertEqual(equity['ticker'], 'H6881')
                self.assertEqual(equity['name'], '(H6881) China Galaxy Securities Co., Ltd.')
                self.assertEqual(equity['currency'], 'HKD')
                self.assertEqual(equity['accounting_treatment'], 'Trading')
                self.assertAlmostEqual(equity['number_of_shares'], 166000)
                self.assertEqual(equity['is_listed'], 'Y')
                self.assertEqual(equity['listed_location'], 'Hong Kong')
                self.assertAlmostEqual(equity['fx_on_trade_day'], 1.03006767544628)
                self.assertEqual(equity['last_trade_date'], datetime(2017,7,21))
                self.assertAlmostEqual(equity['average_cost'], 7.02195114457831)
                self.assertAlmostEqual(equity['price'], 6.93)
                self.assertAlmostEqual(equity['book_cost'], 1165643.89)
                self.assertAlmostEqual(equity['market_value'], 1150380)
                self.assertAlmostEqual(equity['market_gain_loss'], -15263.89)
                self.assertAlmostEqual(equity['fx_gain_loss'], -24.73)
                self.assertEqual(len(equity), 15)


            if (i == 13):    # this should have holding amount = 0
                self.assertEqual(equity['ticker'], 'N2800')
                self.assertEqual(equity['name'], '(N2800) Tracker Fund of HK')
                self.assertEqual(equity['currency'], 'HKD')
                self.assertEqual(equity['accounting_treatment'], 'Trading')
                self.assertAlmostEqual(equity['number_of_shares'], 0)
                self.assertEqual(len(equity), 5)  # should have no more fields



    def validate_equity2(self, equity_holding):
        """
        Validate the USD equity section from CLM BAL 2017-07-27.xls
        """
        self.assertEqual(len(equity_holding), 13) # should have 13 positions
        self.assertEqual(self.count_zero_holding_equity(equity_holding), 11)

        i = 0
        for equity in equity_holding:
            i = i + 1

            if (i == 1):    # the first equity
                self.assertEqual(equity['ticker'], 'US30303M1027')
                self.assertEqual(equity['name'], '(US30303M1027) Facebook Inc.')
                self.assertEqual(equity['currency'], 'USD')
                self.assertEqual(equity['accounting_treatment'], 'Trading')
                self.assertAlmostEqual(equity['number_of_shares'], 850)
                self.assertEqual(equity['is_listed'], 'Y')
                self.assertEqual(equity['listed_location'], 'US')
                self.assertAlmostEqual(equity['fx_on_trade_day'], 8.04421050463015)
                self.assertEqual(equity['last_trade_date'], datetime(2017,7,21))
                self.assertAlmostEqual(equity['average_cost'], 164.473305882353)
                self.assertAlmostEqual(equity['price'], 170.44)
                self.assertAlmostEqual(equity['book_cost'], 139802.31)
                self.assertAlmostEqual(equity['market_value'], 144874)
                self.assertAlmostEqual(equity['market_gain_loss'], 5071.69)
                self.assertAlmostEqual(equity['fx_gain_loss'], -66.36)
                self.assertEqual(len(equity), 15)

            if (i == 2):    # the last non-zero holding
                self.assertEqual(equity['ticker'], 'US01609W1027')
                self.assertEqual(equity['name'], '(US01609W1027) Alibaba Group Holding')
                self.assertEqual(equity['currency'], 'USD')
                self.assertEqual(equity['accounting_treatment'], 'Trading')
                self.assertAlmostEqual(equity['number_of_shares'], 880)
                self.assertEqual(equity['is_listed'], 'Y')
                self.assertEqual(equity['listed_location'], 'US')
                self.assertAlmostEqual(equity['fx_on_trade_day'], 8.04421050463015)
                self.assertEqual(equity['last_trade_date'], datetime(2017,7,21))
                self.assertAlmostEqual(equity['average_cost'], 151.428)
                self.assertAlmostEqual(equity['price'], 154.15)
                self.assertAlmostEqual(equity['book_cost'], 133256.64)
                self.assertAlmostEqual(equity['market_value'], 135652)
                self.assertAlmostEqual(equity['market_gain_loss'], 2395.36)
                self.assertAlmostEqual(equity['fx_gain_loss'], -63.25)
                self.assertEqual(len(equity), 15)

            if (i == 3):    # this should have holding amount = 0
                self.assertEqual(equity['ticker'], 'H0386')
                self.assertEqual(equity['name'], '(H0386) China Petroleum & Chemical Corporation')
                self.assertEqual(equity['currency'], 'USD') # based on program logic,
                                                            # it uses the currency of the quity
                                                            # section, therefore USD
                self.assertEqual(equity['accounting_treatment'], 'Trading')
                self.assertAlmostEqual(equity['number_of_shares'], 0)
                self.assertEqual(len(equity), 5)  # should have no more fields

            if (i == 7):    # this should have holding amount = 0
                self.assertEqual(equity['ticker'], 'N1555')
                self.assertEqual(equity['name'], '(N1555) MIE Holdings Corporation')
                self.assertEqual(equity['currency'], 'USD') # based on program logic,
                                                            # it uses the currency of the quity
                                                            # section, therefore USD
                self.assertEqual(equity['accounting_treatment'], 'Trading')
                self.assertAlmostEqual(equity['number_of_shares'], 0)
                self.assertEqual(len(equity), 5)  # should have no more fields