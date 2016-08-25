# coding=utf-8
# 
# Test overall functionality.
#

import unittest2
import datetime
from xlrd import open_workbook
from DIF.open_cash import read_cash
from DIF.open_summary import read_portfolio_summary
from DIF.open_holding import read_holding
from DIF.open_expense import read_expense
from DIF.open_dif import validate_cash_and_holding, InconsistentValue
from DIF.utility import get_current_path



class TestAll(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestAll, self).__init__(*args, **kwargs)

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



    def test_read_all(self):
        """
        Read the cash and holdings and then validate the numbers are
        read correctly.
        """
        filename = get_current_path() + '\\samples\\sample_DIF_20151210.xls'
        port_values = {}
        wb = open_workbook(filename=filename)
		
        ws = wb.sheet_by_name('Portfolio Sum.')
        read_portfolio_summary(ws, port_values)

		# find sheets that contain cash
        sheet_names = wb.sheet_names()
        for sn in sheet_names:
            if len(sn) > 4 and sn[-4:] == '-BOC':
                ws = wb.sheet_by_name(sn)
                read_cash(ws, port_values)

        ws = wb.sheet_by_name('Portfolio Val.')
        read_holding(ws, port_values)

        # make sure the holding and cash are read correctly
        try:
	        validate_cash_and_holding(port_values)
        except:
        	self.fail('validation failed')

	    # manually adjust the cash total and expect to see failure
        port_values['cash_total'] = port_values['cash_total'] - 0.01
        with self.assertRaises(InconsistentValue):
            validate_cash_and_holding(port_values)



    def test_read_all2(self):
        """
        Read the cash and holdings and then validate the numbers are
        read correctly. This time with a different excel file.
        """
        filename = get_current_path() + '\\samples\\sample_DIF_20151231.xls'
        port_values = {}
        wb = open_workbook(filename=filename)
		
        ws = wb.sheet_by_name('Portfolio Sum.')
        read_portfolio_summary(ws, port_values)

		# find sheets that contain cash
        sheet_names = wb.sheet_names()
        for sn in sheet_names:
            if len(sn) > 4 and sn[-4:] == '-BOC':
                ws = wb.sheet_by_name(sn)
                read_cash(ws, port_values)

        ws = wb.sheet_by_name('Portfolio Val.')
        read_holding(ws, port_values)

        # make sure the holding and cash are read correctly
        try:
	        validate_cash_and_holding(port_values)
        except:
        	self.fail('validation failed')

	    # manually adjust the cash total and expect to see failure
        port_values['equity_total'] = port_values['equity_total'] - 0.01
        with self.assertRaises(InconsistentValue):
            validate_cash_and_holding(port_values)