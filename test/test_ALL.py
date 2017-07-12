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
from DIF.open_dif import open_dif, validate_cash_and_holding, InconsistentValue
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
        try:
            open_dif(filename, port_values, get_current_path() + '\\samples')
        except:
            self.fail('something goes wrong.')



    def test_read_all2(self):
        """
        Read the cash and holdings and then validate the numbers are
        read correctly. This time with a different excel file.
        """
        filename = get_current_path() + '\\samples\\sample_DIF_20151231.xls'
        port_values = {}
        try:
            open_dif(filename, port_values, get_current_path() + '\\samples')
        except:
            self.fail('something goes wrong.')



    def test_read_all3(self):
        """
        With the futures cash account.
        """
        filename = get_current_path() + '\\samples\\CL Franklin DIF 2017-07-10.xls'
        port_values = {}
        
        try:
            open_dif(filename, port_values, get_current_path() + '\\samples')
        except:
            self.fail('something goes wrong.')
