"""
Test error conditions from the open_holding.py module.

"""

import unittest2
from datetime import datetime
from xlrd import open_workbook
from DIF.utility import get_current_path
from DIF.open_holding import read_holding, read_bond_fields, read_currency, \
                            get_datemode, read_equity_fields, \
                            read_sub_section, read_section, BadFieldName, \
                            BadAssetClass

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
        Test the equity with inconsistent currency value
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



    def test_read_equity_field_error1(self):
        """
        Test bad field name type
        """
        filename = get_current_path() + '\\samples\\holdings_error5.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 275    # the equity section starts at A276

        with self.assertRaisesRegexp(TypeError, 'bad field name type'):
            read_equity_fields(ws, row)



    def test_read_equity_field_error2(self):
        """
        Test bad field name (equity)
        """
        filename = get_current_path() + '\\samples\\holdings_error6.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 275    # the equity section starts at A276

        # the field name is not handled
        with self.assertRaises(BadFieldName):
            read_equity_fields(ws, row)



    def test_read_bond_field_error1(self):
        """
        Test bad field name type
        """
        filename = get_current_path() + '\\samples\\holdings_error7.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 62    # the bond section starts at A63

        with self.assertRaisesRegexp(TypeError, 'bad field name type'):
            read_bond_fields(ws, row)



    def test_read_bond_field_error2(self):
        """
        Test bad field name (bond)
        """
        filename = get_current_path() + '\\samples\\holdings_error8.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 62    # the bond section starts at A63

        with self.assertRaises(BadFieldName):
            read_bond_fields(ws, row)



    def test_read_section_error1(self):
        """
        Test bad accounting treatment (bond)
        """
        filename = get_current_path() + '\\samples\\holdings_error9.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 62    # the bond section starts at A63

        fields = ['par_amount', 'is_listed', 'listed_location', 
                    'fx_on_trade_day', 'coupon_rate', 'coupon_start_date', 
                    'maturity_date', 'average_cost', 'amortized_cost', 
                    'book_cost', 'interest_bought', 'amortized_value', 
                    'accrued_interest', 'amortized_gain_loss', 'fx_gain_loss']
        asset_class = 'bond'
        currency = 'USD'
        # bond_holding = []
        port_values = {}

        with self.assertRaisesRegexp(ValueError, 'bad accounting treatment'):
            read_section(ws, row, fields, asset_class, currency, port_values)



    def test_read_section_error2(self):
        """
        Test bad asset class (bond)
        """
        filename = get_current_path() + '\\samples\\holdings_error10.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 62    # the bond section starts at A63

        fields = ['par_amount', 'is_listed', 'listed_location', 
                    'fx_on_trade_day', 'coupon_rate', 'coupon_start_date', 
                    'maturity_date', 'average_cost', 'amortized_cost', 
                    'book_cost', 'interest_bought', 'amortized_value', 
                    'accrued_interest', 'amortized_gain_loss', 'fx_gain_loss']
        
        asset_class = 'Bond'	# it should be 'bond', this creates an error
        
        currency = 'USD'
        # bond_holding = []
        port_values = {}

        with self.assertRaises(BadAssetClass):
            read_section(ws, row, fields, asset_class, currency, port_values)