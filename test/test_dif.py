# coding=utf-8
# 
# Test overall functionality.
#

import unittest2
from xlrd import open_workbook
from DIF.open_dif import open_dif
from DIF.utility import get_current_path



class TestDIF(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestDIF, self).__init__(*args, **kwargs)

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



    def test_read_dif(self):
        """
        Read the cash and holdings and then validate the numbers are
        read correctly.
        """
        filename = get_current_path() + '\\samples\\sample_DIF_20151210.xls'
        port_values = {}
        open_dif(filename, port_values)
        