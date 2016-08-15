"""
Test the read_holding() method from open_holding.py

"""

import unittest2
from datetime import datetime
from xlrd import open_workbook
from DIF.utility import get_current_path
from DIF.open_holding import read_holding, read_bond_fields, read_currency, \
                            get_datemode, read_equity_fields, \
                            read_sub_section, read_section

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



    def test_read_bond_fields_HTM(self):
        """
        Test the read_bond_fields() method using holdings_sample2.xls,
        containing only one bond section with HTM bonds.
        """
        filename = get_current_path() + '\\samples\\holdings_sample2.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 62    # the bond section starts at A63

        fields, n = read_bond_fields(ws, row)
        self.assertEqual(n, 4)
        self.assertEqual(len(fields), 15)

        f = ''
        for s in fields:
            f = f + s + ', '

        # check the fields are read correctly
        self.assertEqual(f, 
            'par_amount, is_listed, listed_location, fx_on_trade_day, coupon_rate, coupon_start_date, maturity_date, average_cost, amortized_cost, book_cost, interest_bought, amortized_value, accrued_interest, amortized_gain_loss, fx_gain_loss, ')



    def test_read_bond_fields_trading(self):
        """
        Test the read_bond_fields() method using holdings_sample3.xls,
        containing only one bond section with trading bonds.
        """

        filename = get_current_path() + '\\samples\\holdings_sample3.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 114    # the bond section starts at A115

        fields, n = read_bond_fields(ws, row)
        self.assertEqual(n, 4)
        self.assertEqual(len(fields), 15)

        f = ''
        for s in fields:
            f = f + s + ', '

        # check the fields are read correctly
        self.assertEqual(f, 
            'par_amount, is_listed, listed_location, fx_on_trade_day, coupon_rate, coupon_start_date, maturity_date, average_cost, price, book_cost, interest_bought, market_value, accrued_interest, market_gain_loss, fx_gain_loss, ')



    def test_read_equity_fields_listed(self):
        """
        Test the read_equity_fields() method with listed eqiuty.
        """

        filename = get_current_path() + '\\samples\\holdings_sample_equity1.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 275    # the bond section starts at A276

        fields, n = read_equity_fields(ws, row)
        self.assertEqual(n, 5)
        self.assertEqual(len(fields), 15)

        f = ''
        for s in fields:
            f = f + s + ', '

        # check the fields are read correctly
        self.assertEqual(f, 
            'number_of_shares, currency, listed_location, fx_on_trade_day, empty_field, last_trade_date, empty_field, average_cost, price, book_cost, empty_field, market_value, empty_field, market_gain_loss, fx_gain_loss, ')


    def test_read_equity_fields_preferred_shares(self):
        """
        Test the read_equity_fields() method with listed eqiuty.
        """

        filename = get_current_path() + '\\samples\\holdings_sample_equity2.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 301    # the bond section starts at A302

        fields, n = read_equity_fields(ws, row)
        self.assertEqual(n, 3)
        self.assertEqual(len(fields), 15)

        f = ''
        for s in fields:
            f = f + s + ', '

        # check the fields are read correctly
        self.assertEqual(f, 
            'number_of_shares, currency, empty_field, fx_on_trade_day, empty_field, last_trade_date, empty_field, average_cost, price, book_cost, empty_field, market_value, empty_field, market_gain_loss, fx_gain_loss, ')



    def test_read_currency(self):
        msg = 'V. Debt Securities - US$  (債務票據- 美元)'
        self.assertEqual(read_currency(msg), 'USD')

        msg = 'V. Debt Securities - SGD  (債務票據- 星加坡元)'
        self.assertEqual(read_currency(msg), 'SGD')
        
        msg = 'X. Equities - USD (股票-美元)'
        self.assertEqual(read_currency(msg), 'USD')



    def test_read_sub_section_bond_HTM(self):
        """
        Test the HTM bonds
        """
        filename = get_current_path() + '\\samples\\holdings_sample2.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 68    # the bond sub section starts at A69
        accounting_treatment = 'HTM'
        fields = ['par_amount', 'is_listed', 'listed_location', 
                    'fx_on_trade_day', 'coupon_rate', 'coupon_start_date', 
                    'maturity_date', 'average_cost', 'amortized_cost', 
                    'book_cost', 'interest_bought', 'amortized_value', 
                    'accrued_interest', 'amortized_gain_loss', 'fx_gain_loss']
        asset_class = 'bond'
        currency = 'USD'
        bond_holding = []

        read_sub_section(ws, row, accounting_treatment, fields, asset_class, currency, bond_holding)

        self.validate_bond_HTM(bond_holding)



    def test_read_sub_section_bond_trading(self):
        """
        Test the trading bonds
        """
        filename = get_current_path() + '\\samples\\holdings_sample3.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 120    # the bond sub section starts at A121

        accounting_treatment = 'Trading'
        
        # fields for trading bonds
        fields = ['par_amount', 'is_listed', 'listed_location', 
                    'fx_on_trade_day', 'coupon_rate', 'coupon_start_date', 
                    'maturity_date', 'average_cost', 'price', 
                    'book_cost', 'interest_bought', 'market_value', 
                    'accrued_interest', 'market_gain_loss', 'fx_gain_loss']
        asset_class = 'bond'
        currency = 'USD'
        bond_holding = []

        read_sub_section(ws, row, accounting_treatment, fields, asset_class, currency, bond_holding)

        self.validate_bond_trading(bond_holding)


    def test_read_section_bond(self):

        filename = get_current_path() + '\\samples\\holdings_sample3.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 114    # the bond section starts at A115

        # fields for trading bonds
        fields = ['par_amount', 'is_listed', 'listed_location', 
                    'fx_on_trade_day', 'coupon_rate', 'coupon_start_date', 
                    'maturity_date', 'average_cost', 'price', 
                    'book_cost', 'interest_bought', 'market_value', 
                    'accrued_interest', 'market_gain_loss', 'fx_gain_loss']
        asset_class = 'bond'
        currency = 'USD'
        bond_holding = []

        read_section(ws, row, fields, asset_class, currency, bond_holding)
        self.validate_bond_trading(bond_holding)
        


    def test_read_section_listed_equity(self):

        filename = get_current_path() + '\\samples\\holdings_sample_equity1.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 275    # the equity section starts at A276

        # fields for trading bonds
        fields = ['number_of_shares', 'currency', 'listed_location', 
                    'fx_on_trade_day', 'empty_field', 'last_trade_date', 
                    'empty_field', 'average_cost', 'price', 'book_cost', 
                    'empty_field', 'market_value', 'empty_field', 
                    'market_gain_loss', 'fx_gain_loss']
        asset_class = 'equity'
        currency = 'HKD'
        equity_holding = []

        read_section(ws, row, fields, asset_class, currency, equity_holding)

        self.validate_listed_equity(equity_holding)
        

    def test_read_section_preferred_shares(self):

        filename = get_current_path() + '\\samples\\holdings_sample_equity2.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 301    # the equity section starts at A302

        # fields for trading bonds
        fields = ['number_of_shares', 'currency', 'empty_field', 
                    'fx_on_trade_day', 'empty_field', 'last_trade_date', 
                    'empty_field', 'average_cost', 'price', 'book_cost', 
                    'empty_field', 'market_value', 'empty_field', 
                    'market_gain_loss', 'fx_gain_loss']
        asset_class = 'equity'
        currency = 'USD'
        equity_holding = []

        read_section(ws, row, fields, asset_class, currency, equity_holding)

        self.validate_preferred_shares(equity_holding)



    def test_open_holding_all(self):
        """
        Read the holding file, the statistics are:

        Bond        
                        total   zero holding
        USD (HTM)       25      2
        USD (Trading)   61      39
        SGD (HTM)       10      10
        SGD (Trading)   1       1
        CNY (HTM)       2       0
        CNY (Trading)   15      13
        total           114     65
                
        Equity      
                        total   zero holding
        HKD             10      5
        USD             4       2       
        total           14      7
        """
        filename = get_current_path() + '\\samples\\holdings_sample.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')

        port_values = {}
        read_holding(ws, port_values)

        bond_holding = port_values['bond']
        self.assertEqual(len(bond_holding), 114)
        self.assertEqual(self.count_zero_holding_bond(bond_holding), 65)

        equity_holding = port_values['equity']
        self.assertEqual(len(equity_holding), 14)
        self.assertEqual(self.count_zero_holding_equity(equity_holding), 7)

        bond_holding_HTM_USD = self.extract_bond_holding(bond_holding, 'USD', 'HTM')
        self.validate_bond_HTM(bond_holding_HTM_USD)

        bond_holding_Trading_USD = self.extract_bond_holding(bond_holding, 'USD', 'Trading')
        self.validate_bond_trading(bond_holding_Trading_USD)

        bond_holding_HTM_SGD = self.extract_bond_holding(bond_holding, 'SGD', 'HTM')
        self.assertEqual(len(bond_holding_HTM_SGD), 10)
        self.assertEqual(self.count_zero_holding_bond(bond_holding_HTM_SGD), 10)

        bond_holding_Trading_SGD = self.extract_bond_holding(bond_holding, 'SGD', 'Trading')
        self.assertEqual(len(bond_holding_Trading_SGD), 1)
        self.assertEqual(self.count_zero_holding_bond(bond_holding_Trading_SGD), 1)

        bond_holding_HTM_CNY = self.extract_bond_holding(bond_holding, 'CNY', 'HTM')
        self.assertEqual(len(bond_holding_HTM_CNY), 2)
        self.assertEqual(self.count_zero_holding_bond(bond_holding_HTM_CNY), 0)

        bond_holding_Trading_CNY = self.extract_bond_holding(bond_holding, 'CNY', 'Trading')
        self.assertEqual(len(bond_holding_Trading_CNY), 15)
        self.assertEqual(self.count_zero_holding_bond(bond_holding_Trading_CNY), 13)

        equity_holding_listed = self.extract_equity_holding(equity_holding, 'HKD')
        self.validate_listed_equity(equity_holding_listed)

        equity_holding_preferred_shares = self.extract_equity_holding(equity_holding, 'USD')
        self.validate_preferred_shares(equity_holding_preferred_shares)



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
        Validate bond positions from holdings_sample3.xls, where there is
        only one section for bond for trading (USD bond for trading).
        """
        self.assertEqual(len(bond_holding), 61) # should have 61 positions
        self.assertEqual(self.count_zero_holding_bond(bond_holding), 39)

        i = 0
        for bond in bond_holding:
            i = i + 1

            if (i == 1):    # the first bond
                self.assertEqual(bond['isin'], 'US404280AS86')
                self.assertEqual(bond['name'], '(US404280AS86) HSBC Holding Plc 6.375%')
                self.assertEqual(bond['currency'], 'USD')
                self.assertEqual(bond['accounting_treatment'], 'Trading')
                self.assertAlmostEqual(bond['par_amount'], 400000)
                self.assertEqual(bond['is_listed'], 'Y')
                self.assertEqual(bond['listed_location'], 'TBC')
                self.assertAlmostEqual(bond['fx_on_trade_day'], 7.7553)
                self.assertAlmostEqual(bond['coupon_rate'], 6.375/100)
                self.assertEqual(bond['coupon_start_date'], datetime(2015,9,17))
                self.assertEqual(bond['maturity_date'], datetime(2049,12,29))
                self.assertAlmostEqual(bond['average_cost'], 101.125)
                self.assertAlmostEqual(bond['price'], 98.719)
                self.assertAlmostEqual(bond['book_cost'], 404500)
                self.assertAlmostEqual(bond['interest_bought'], 0)
                self.assertAlmostEqual(bond['market_value'], 394876)
                self.assertAlmostEqual(bond['accrued_interest'], 5950)
                self.assertAlmostEqual(bond['market_gain_loss'], -9624)
                self.assertAlmostEqual(bond['fx_gain_loss'], -2062.95000000018)

            if (i == 20):    # this should have holding amount = 0
                self.assertEqual(bond['isin'], 'USY68856AQ98')
                self.assertEqual(bond['name'], '(USY68856AQ98) Petronas Capital Ltd 4.5%')
                self.assertEqual(bond['currency'], 'USD')
                self.assertEqual(bond['accounting_treatment'], 'Trading')
                self.assertAlmostEqual(bond['par_amount'], 0)
                self.assertEqual(len(bond), 5)  # should have no more fields

            if (i == 24):    # this should have holding amount = 0
                self.assertEqual(bond['isin'], 'XS1219829949')
                self.assertEqual(bond['name'], '(XS1219829949) HAITONG INTL FIN 2015 3.5%')
                self.assertEqual(bond['currency'], 'USD')
                self.assertEqual(bond['accounting_treatment'], 'Trading')
                self.assertAlmostEqual(bond['par_amount'], 0)
                self.assertEqual(len(bond), 5)  # should have no more fields

            if (i == 61):    # the last bond
                self.assertEqual(bond['isin'], 'XS1329465667')
                self.assertEqual(bond['name'], '(XS1329465667) TOP LUXURY INV LTD 4.99%')
                self.assertEqual(bond['currency'], 'USD')
                self.assertEqual(bond['accounting_treatment'], 'Trading')
                self.assertAlmostEqual(bond['par_amount'], 5000000)
                self.assertEqual(bond['is_listed'], 'Y')
                self.assertEqual(bond['listed_location'], 'TBC')
                self.assertAlmostEqual(bond['fx_on_trade_day'], 7.7522)
                self.assertAlmostEqual(bond['coupon_rate'], 4.99/100)
                self.assertEqual(bond['coupon_start_date'], datetime(2015,12,17))
                self.assertEqual(bond['maturity_date'], datetime(2040,12,17))
                self.assertAlmostEqual(bond['average_cost'], 97.353)
                self.assertAlmostEqual(bond['price'], 97.353)
                self.assertAlmostEqual(bond['book_cost'], 4867650)
                self.assertAlmostEqual(bond['interest_bought'], 0)
                self.assertAlmostEqual(bond['market_value'], 4867650)
                self.assertAlmostEqual(bond['accrued_interest'], 0)
                self.assertAlmostEqual(bond['market_gain_loss'], 0)
                self.assertAlmostEqual(bond['fx_gain_loss'], -9735.29999999701)



    def validate_bond_HTM(self, bond_holding):
        """
        Validate bond positions from holdings_sample2.xls, where there is
        only one section for bond for HTM (USD bond for HTM).
        """
        self.assertEqual(len(bond_holding), 25) # should have 25 positions
        self.assertEqual(self.count_zero_holding_bond(bond_holding), 2)

        i = 0
        for bond in bond_holding:
            i = i + 1

            if (i == 1):    # the first bond
                self.assertEqual(bond['isin'], 'XS1021617698')
                self.assertEqual(bond['name'], '(XS1021617698) Yuzhou Properties Co. Ltd 8.625%')
                self.assertEqual(bond['currency'], 'USD')
                self.assertEqual(bond['accounting_treatment'], 'HTM')
                self.assertAlmostEqual(bond['par_amount'], 2100000)
                self.assertEqual(bond['is_listed'], 'Y')
                self.assertEqual(bond['listed_location'], 'TBC')
                self.assertAlmostEqual(bond['fx_on_trade_day'], 7.75290040843624)
                self.assertAlmostEqual(bond['coupon_rate'], 8.625/100)
                self.assertEqual(bond['coupon_start_date'], datetime(2015,7,24))
                self.assertEqual(bond['maturity_date'], datetime(2019,1,24))
                self.assertAlmostEqual(bond['average_cost'], 99.6833333333333)
                self.assertAlmostEqual(bond['amortized_cost'], 99.8946738095238)
                self.assertAlmostEqual(bond['book_cost'], 2093350)
                self.assertAlmostEqual(bond['interest_bought'], 11020.83)
                self.assertAlmostEqual(bond['amortized_value'], 2097788.15)
                self.assertAlmostEqual(bond['accrued_interest'], 68928.13)
                self.assertAlmostEqual(bond['amortized_gain_loss'], 4438.1499999999)
                self.assertAlmostEqual(bond['fx_gain_loss'], -5652.90000000037)

            if (i == 5):    # this should have holding amount = 0
                self.assertEqual(bond['isin'], 'USG52132AF72')
                self.assertEqual(bond['name'], '(USG52132AF72) Kaisa Group Holdings Ltd 8.875%')
                self.assertEqual(bond['currency'], 'USD')
                self.assertEqual(bond['accounting_treatment'], 'HTM')
                self.assertAlmostEqual(bond['par_amount'], 0)
                self.assertEqual(len(bond), 5)  # should have no more fields

            if (i == 12):    # this should have holding amount = 0
                self.assertEqual(bond['isin'], 'XS0782027857')
                self.assertEqual(bond['name'], '(XS0782027857) Sound Global Ltd 11.875%')
                self.assertEqual(bond['currency'], 'USD')
                self.assertEqual(bond['accounting_treatment'], 'HTM')
                self.assertAlmostEqual(bond['par_amount'], 0)
                self.assertEqual(len(bond), 5)  # should have no more fields

            if (i == 25):    # the last bond
                self.assertEqual(bond['isin'], 'XS1164776020')
                self.assertEqual(bond['name'], '(XS1164776020) COUNTRY GARDEN HLDG CO 7.5%')
                self.assertEqual(bond['currency'], 'USD')
                self.assertEqual(bond['accounting_treatment'], 'HTM')
                self.assertAlmostEqual(bond['par_amount'], 500000)
                self.assertEqual(bond['is_listed'], 'Y')
                self.assertEqual(bond['listed_location'], 'TBC')
                self.assertAlmostEqual(bond['fx_on_trade_day'], 7.7506)
                self.assertAlmostEqual(bond['coupon_rate'], 7.5/100)
                self.assertEqual(bond['coupon_start_date'], datetime(2015,9,9))
                self.assertEqual(bond['maturity_date'], datetime(2020,3,9))
                self.assertAlmostEqual(bond['average_cost'], 105.2)
                self.assertAlmostEqual(bond['amortized_cost'], 105.134344)
                self.assertAlmostEqual(bond['book_cost'], 526000)
                self.assertAlmostEqual(bond['interest_bought'], 7291.67)
                self.assertAlmostEqual(bond['amortized_value'], 525671.72)
                self.assertAlmostEqual(bond['accrued_interest'], 9583.33)
                self.assertAlmostEqual(bond['amortized_gain_loss'], -328.280000000027)
                self.assertAlmostEqual(bond['fx_gain_loss'], -210.399999999906)



    def validate_listed_equity(self, equity_holding):
        """
        Validate equity positions from holdings_sample_equity1.xls, where there is
        only one section for listed equity.
        """
        self.assertEqual(len(equity_holding), 10) # should have 10 positions
        self.assertEqual(self.count_zero_holding_equity(equity_holding), 5)

        i = 0
        for equity in equity_holding:
            i = i + 1

            if (i == 1):    # the first equity
                self.assertEqual(equity['ticker'], 'H0939')
                self.assertEqual(equity['name'], '(H0939) China Construction Bank Corporation')
                self.assertEqual(equity['currency'], 'HKD')
                self.assertEqual(equity['accounting_treatment'], 'Trading')
                self.assertAlmostEqual(equity['number_of_shares'], 100000)
                self.assertEqual(equity['listed_location'], 'Hong Kong')
                self.assertAlmostEqual(equity['fx_on_trade_day'], 1.0)
                self.assertEqual(equity['last_trade_date'], datetime(2015,11,13))
                self.assertAlmostEqual(equity['average_cost'], 5.4512989)
                self.assertAlmostEqual(equity['price'], 5.2)
                self.assertAlmostEqual(equity['book_cost'], 545129.89)
                self.assertAlmostEqual(equity['market_value'], 520000)
                self.assertAlmostEqual(equity['market_gain_loss'], -25129.89)
                self.assertAlmostEqual(equity['fx_gain_loss'], 0)
                self.assertEqual(len(equity), 14)  # should have no more fields

            if (i == 2):    # this should have holding amount = 0
                self.assertEqual(equity['ticker'], 'H1508')
                self.assertEqual(equity['name'], '(H1508) China Reinsurance (Group) Corporation - H Shares')
                self.assertEqual(equity['currency'], 'HKD')
                self.assertEqual(equity['accounting_treatment'], 'Trading')
                self.assertAlmostEqual(equity['number_of_shares'], 0)
                self.assertEqual(len(equity), 5)  # should have no more fields

            if (i == 10):    # the last equity
                self.assertEqual(equity['ticker'], 'N2388')
                self.assertEqual(equity['name'], '(N2388) BOC Hong Kong Holdings Ltd')
                self.assertEqual(equity['currency'], 'HKD')
                self.assertEqual(equity['accounting_treatment'], 'Trading')
                self.assertAlmostEqual(equity['number_of_shares'], 60000)
                self.assertEqual(equity['listed_location'], 'Hong Kong')
                self.assertAlmostEqual(equity['fx_on_trade_day'], 1.0)
                self.assertEqual(equity['last_trade_date'], datetime(2015,11,27))
                self.assertAlmostEqual(equity['average_cost'], 24.2753238333333)
                self.assertAlmostEqual(equity['price'], 23.5)
                self.assertAlmostEqual(equity['book_cost'], 1456519.43)
                self.assertAlmostEqual(equity['market_value'], 1410000)
                self.assertAlmostEqual(equity['market_gain_loss'], -46519.4300000001)
                self.assertAlmostEqual(equity['fx_gain_loss'], 0)
                self.assertEqual(len(equity), 14)  # should have no more fields



    def validate_preferred_shares(self, equity_holding):
        """
        Validate equity positions from holdings_sample_equity2.xls, where there is
        only one section for preferred shares (treated as equity).
        """
        self.assertEqual(len(equity_holding), 4) # should have 4 positions
        self.assertEqual(self.count_zero_holding_equity(equity_holding), 2)

        i = 0
        for equity in equity_holding:
            i = i + 1

            if (i == 1):    # this should have holding amount = 0
                self.assertEqual(equity['isin'], 'XS1122780106')
                self.assertEqual(equity['name'], '(XS1122780106) Bank of China 6.75%')
                self.assertEqual(equity['currency'], 'USD')
                self.assertEqual(equity['accounting_treatment'], 'Trading')
                self.assertAlmostEqual(equity['number_of_shares'], 0)
                self.assertEqual(len(equity), 5)  # should have no more fields

            if (i == 3):    # the first equity
                self.assertEqual(equity['isin'], 'USY39656AA40')
                self.assertEqual(equity['name'], '(USY39656AA40) ICBCAS 6%')
                self.assertEqual(equity['currency'], 'USD')
                self.assertEqual(equity['accounting_treatment'], 'Trading')
                self.assertAlmostEqual(equity['number_of_shares'], 2200000)
                self.assertAlmostEqual(equity['fx_on_trade_day'], 7.75165270664211)
                self.assertEqual(equity['last_trade_date'], datetime(2015,10,13))
                self.assertAlmostEqual(equity['average_cost'], 104.960605909091)
                self.assertAlmostEqual(equity['price'], 106.068)
                self.assertAlmostEqual(equity['book_cost'], 2309133.33)
                self.assertAlmostEqual(equity['market_value'], 2333496)
                self.assertAlmostEqual(equity['market_gain_loss'], 24362.6699999999)
                self.assertAlmostEqual(equity['fx_gain_loss'], -3354.5)
                self.assertEqual(len(equity), 13)  # should have no more fields

            if (i == 4):    # the last equity
                self.assertEqual(equity['isin'], 'XS1328130197')
                self.assertEqual(equity['name'], '(XS1328130197) CHINA CONSTRUCTION BANK 4.65%')
                self.assertEqual(equity['currency'], 'USD')
                self.assertEqual(equity['accounting_treatment'], 'Trading')
                self.assertAlmostEqual(equity['number_of_shares'], 2000000)
                self.assertAlmostEqual(equity['fx_on_trade_day'], 7.7502)
                self.assertEqual(equity['last_trade_date'], datetime(2015,12,10))
                self.assertAlmostEqual(equity['average_cost'], 100)
                self.assertAlmostEqual(equity['price'], 100)
                self.assertAlmostEqual(equity['book_cost'], 2000000)
                self.assertAlmostEqual(equity['market_value'], 2000000)
                self.assertAlmostEqual(equity['market_gain_loss'], 0)
                self.assertAlmostEqual(equity['fx_gain_loss'], 0)
                self.assertEqual(len(equity), 13)  # should have no more fields