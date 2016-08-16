"""
Test error conditions from the open_holding.py module.

"""

import unittest2
from datetime import datetime
from xlrd import open_workbook
from DIF.utility import get_current_path
from DIF.open_holding import read_holding, read_bond_fields, read_currency, \
                            get_datemode, read_equity_fields, \
                            read_sub_section, read_section

class TestHoldingError(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestHoldingError, self).__init__(*args, **kwargs)

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



    def foo(self, x):
        raise ValueError('test')



    def bar(self):
        pass



    def test_exp(self):
        """
        Test whether a specific exception happens, with a specific string
        representation. See: https://docs.python.org/2/library/unittest.html
        """

        # verify that ValueError happens
        self.assertRaises(ValueError, self.foo, 5)

        # Or we can use the context manager
        with self.assertRaises(ValueError) as cm:
            self.foo(5)

        # verify that the string representation is 'test'
        the_exception = cm.exception
        self.assertEqual(str(the_exception), 'test')

        # Or similarly, verfiy that ValueError happens whose string
        # representation matches a regular expression. Note this is
        # not exactly the same as the above.
        self.assertRaisesRegexp(ValueError, 'es', self.foo, 5)

        # same as the above
        with self.assertRaisesRegexp(ValueError, 'es'):
            self.foo(5)

        # also true
        with self.assertRaisesRegexp(ValueError, 'test'):
            self.foo(5)

        # fails, becuase no exception is generated.
        # self.assertRaises(Exception, self.bar)



    def test_read_sub_section_error1(self):
        """
        Test the HTM bonds
        """
        filename = get_current_path() + '\\samples\\holdings_error1.xls'
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

        with self.assertRaisesRegexp(ValueError, 'bad field type: not a string'):
            read_sub_section(ws, row, accounting_treatment, fields, asset_class, currency, bond_holding)

        self.assertEqual(len(bond_holding), 2)  # the first 2 bond positions are OK


    def test_read_sub_section_error2(self):
        """
        Test the HTM bonds
        """
        filename = get_current_path() + '\\samples\\holdings_error2.xls'
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

        with self.assertRaisesRegexp(ValueError, 'bad field type: not a float'):
            read_sub_section(ws, row, accounting_treatment, fields, asset_class, currency, bond_holding)

        self.assertEqual(len(bond_holding), 6)  # the first 2 bond positions are OK



    def test_read_sub_section_error3(self):
        """
        Test the HTM bonds
        """
        filename = get_current_path() + '\\samples\\holdings_error3.xls'
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

        with self.assertRaisesRegexp(ValueError, 'bad field type: not a float or empty string'):
            read_sub_section(ws, row, accounting_treatment, fields, asset_class, currency, bond_holding)

        self.assertEqual(len(bond_holding), 6)  # the first 2 bond positions are OK



    def test_read_sub_section_error4(self):
        """
        Test the HTM bonds
        """
        filename = get_current_path() + '\\samples\\holdings_error4.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 282    # the bond sub section starts at A283
        accounting_treatment = 'Trading'
        fields = ['number_of_shares', 'currency', 'listed_location', 
                    'fx_on_trade_day', 'empty_field', 'last_trade_date', 
                    'empty_field', 'average_cost', 'price', 'book_cost', 
                    'empty_field', 'market_value', 'empty_field', 
                    'market_gain_loss', 'fx_gain_loss']
        asset_class = 'equity'
        currency = 'HKD'
        equity_holding = []

        with self.assertRaisesRegexp(ValueError, 'inconsistent currency value'):
            read_sub_section(ws, row, accounting_treatment, fields, asset_class, currency, equity_holding)

        self.assertEqual(len(equity_holding), 2)  # the first 2 equity positions are OK
        										  # note: A285 is not treated as a position
        										  # A286 has amount zero, therefore currency
        										  # is not read at all.


    # def test_read_sub_section_bond_trading(self):
    #     """
    #     Test the trading bonds
    #     """
    #     filename = get_current_path() + '\\samples\\holdings_sample3.xls'
    #     wb = open_workbook(filename=filename)
    #     ws = wb.sheet_by_name('Portfolio Val.')
    #     row = 120    # the bond sub section starts at A121

    #     accounting_treatment = 'Trading'
        
    #     # fields for trading bonds
    #     fields = ['par_amount', 'is_listed', 'listed_location', 
    #                 'fx_on_trade_day', 'coupon_rate', 'coupon_start_date', 
    #                 'maturity_date', 'average_cost', 'price', 
    #                 'book_cost', 'interest_bought', 'market_value', 
    #                 'accrued_interest', 'market_gain_loss', 'fx_gain_loss']
    #     asset_class = 'bond'
    #     currency = 'USD'
    #     bond_holding = []

    #     read_sub_section(ws, row, accounting_treatment, fields, asset_class, currency, bond_holding)

    #     self.validate_bond_trading(bond_holding)


    # def test_read_section_bond(self):

    #     filename = get_current_path() + '\\samples\\holdings_sample3.xls'
    #     wb = open_workbook(filename=filename)
    #     ws = wb.sheet_by_name('Portfolio Val.')
    #     row = 114    # the bond section starts at A115

    #     # fields for trading bonds
    #     fields = ['par_amount', 'is_listed', 'listed_location', 
    #                 'fx_on_trade_day', 'coupon_rate', 'coupon_start_date', 
    #                 'maturity_date', 'average_cost', 'price', 
    #                 'book_cost', 'interest_bought', 'market_value', 
    #                 'accrued_interest', 'market_gain_loss', 'fx_gain_loss']
    #     asset_class = 'bond'
    #     currency = 'USD'
    #     bond_holding = []

    #     read_section(ws, row, fields, asset_class, currency, bond_holding)
    #     self.validate_bond_trading(bond_holding)
        


    # def test_read_section_listed_equity(self):

    #     filename = get_current_path() + '\\samples\\holdings_sample_equity1.xls'
    #     wb = open_workbook(filename=filename)
    #     ws = wb.sheet_by_name('Portfolio Val.')
    #     row = 275    # the equity section starts at A276

    #     # fields for trading bonds
    #     fields = ['number_of_shares', 'currency', 'listed_location', 
    #                 'fx_on_trade_day', 'empty_field', 'last_trade_date', 
    #                 'empty_field', 'average_cost', 'price', 'book_cost', 
    #                 'empty_field', 'market_value', 'empty_field', 
    #                 'market_gain_loss', 'fx_gain_loss']
    #     asset_class = 'equity'
    #     currency = 'HKD'
    #     equity_holding = []

    #     read_section(ws, row, fields, asset_class, currency, equity_holding)

    #     self.validate_listed_equity(equity_holding)
        

    # def test_read_section_preferred_shares(self):

    #     filename = get_current_path() + '\\samples\\holdings_sample_equity2.xls'
    #     wb = open_workbook(filename=filename)
    #     ws = wb.sheet_by_name('Portfolio Val.')
    #     row = 301    # the equity section starts at A302

    #     # fields for trading bonds
    #     fields = ['number_of_shares', 'currency', 'empty_field', 
    #                 'fx_on_trade_day', 'empty_field', 'last_trade_date', 
    #                 'empty_field', 'average_cost', 'price', 'book_cost', 
    #                 'empty_field', 'market_value', 'empty_field', 
    #                 'market_gain_loss', 'fx_gain_loss']
    #     asset_class = 'equity'
    #     currency = 'USD'
    #     equity_holding = []

    #     read_section(ws, row, fields, asset_class, currency, equity_holding)

    #     self.validate_preferred_shares(equity_holding)



    # def test_open_holding_all(self):
    #     """
    #     Read the holding file, the statistics are:

    #     Bond        
    #                     total   zero holding
    #     USD (HTM)       25      2
    #     USD (Trading)   61      39
    #     SGD (HTM)       10      10
    #     SGD (Trading)   1       1
    #     CNY (HTM)       2       0
    #     CNY (Trading)   15      13
    #     total           114     65
                
    #     Equity      
    #                     total   zero holding
    #     HKD             10      5
    #     USD             4       2       
    #     total           14      7
    #     """
    #     filename = get_current_path() + '\\samples\\holdings_sample.xls'
    #     wb = open_workbook(filename=filename)
    #     ws = wb.sheet_by_name('Portfolio Val.')

    #     port_values = {}
    #     read_holding(ws, port_values)

    #     bond_holding = port_values['bond']
    #     self.assertEqual(len(bond_holding), 114)
    #     self.assertEqual(self.count_zero_holding_bond(bond_holding), 65)

    #     equity_holding = port_values['equity']
    #     self.assertEqual(len(equity_holding), 14)
    #     self.assertEqual(self.count_zero_holding_equity(equity_holding), 7)

    #     bond_holding_HTM_USD = self.extract_bond_holding(bond_holding, 'USD', 'HTM')
    #     self.validate_bond_HTM(bond_holding_HTM_USD)

    #     bond_holding_Trading_USD = self.extract_bond_holding(bond_holding, 'USD', 'Trading')
    #     self.validate_bond_trading(bond_holding_Trading_USD)

    #     bond_holding_HTM_SGD = self.extract_bond_holding(bond_holding, 'SGD', 'HTM')
    #     self.assertEqual(len(bond_holding_HTM_SGD), 10)
    #     self.assertEqual(self.count_zero_holding_bond(bond_holding_HTM_SGD), 10)

    #     bond_holding_Trading_SGD = self.extract_bond_holding(bond_holding, 'SGD', 'Trading')
    #     self.assertEqual(len(bond_holding_Trading_SGD), 1)
    #     self.assertEqual(self.count_zero_holding_bond(bond_holding_Trading_SGD), 1)

    #     bond_holding_HTM_CNY = self.extract_bond_holding(bond_holding, 'CNY', 'HTM')
    #     self.assertEqual(len(bond_holding_HTM_CNY), 2)
    #     self.assertEqual(self.count_zero_holding_bond(bond_holding_HTM_CNY), 0)

    #     bond_holding_Trading_CNY = self.extract_bond_holding(bond_holding, 'CNY', 'Trading')
    #     self.assertEqual(len(bond_holding_Trading_CNY), 15)
    #     self.assertEqual(self.count_zero_holding_bond(bond_holding_Trading_CNY), 13)

    #     equity_holding_listed = self.extract_equity_holding(equity_holding, 'HKD')
    #     self.validate_listed_equity(equity_holding_listed)

    #     equity_holding_preferred_shares = self.extract_equity_holding(equity_holding, 'USD')
    #     self.validate_preferred_shares(equity_holding_preferred_shares)



    # def extract_bond_holding(self, bond_holding, currency, accounting_treatment):
    #     """
    #     Extract bond holding given its currency and accounting treatment.
    #     """
    #     holding = []
    #     for bond in bond_holding:
    #         if bond['currency'] == currency and \
    #             bond['accounting_treatment'] == accounting_treatment:
    #             holding.append(bond)

    #     return holding



    # def extract_equity_holding(self, equity_holding, currency):
    #     """
    #     Extract equity holding given its currency.
    #     """
    #     holding = []
    #     for equity in equity_holding:
    #         if equity['currency'] == currency:
    #             holding.append(equity)

    #     return holding



    # def count_zero_holding_bond(self, bond_holding):
    #     """
    #     Count how many bonds in the holding has zero amount (par_amount).
    #     """
    #     empty_bond = 0
    #     for bond in bond_holding:
    #         if bond['par_amount'] == 0:
    #             empty_bond = empty_bond + 1

    #     return empty_bond



    # def count_zero_holding_equity(self, equity_holding):
    #     """
    #     Count how many equity in the holding has zero amount (number_of_shares).
    #     """
    #     empty_equity = 0
    #     for equity in equity_holding:
    #         if equity['number_of_shares'] == 0:
    #             empty_equity = empty_equity + 1

    #     return empty_equity