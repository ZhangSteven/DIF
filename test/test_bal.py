"""
Test the open_bal() method to open the trustee Macau Balanced Fund.
"""

import unittest2
import os, datetime
from xlrd import open_workbook
from DIF.utility import get_current_path
from DIF.open_bal import open_bal, read_portfolio_summary



class TestBAL(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestBAL, self).__init__(*args, **kwargs)



    def test_summary(self):
        filename = os.path.join(get_current_path(), 'samples', 'CLM BAL 2017-07-27.xls')
        port_values = {}
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Sum.')
        read_portfolio_summary(ws, port_values)
        self.assertEqual(port_values['date'], datetime.datetime(2017,7,27))
        self.assertEqual(port_values['portfolio_id'], '30004')
        self.assertAlmostEqual(port_values['number_of_units'], 4837037.6736096)
        self.assertAlmostEqual(port_values['unit_price'], 11.892255)
        self.assertAlmostEqual(port_values['nav'], 57523285.86)



    def test_summary2(self):
        filename = os.path.join(get_current_path(), 'samples', 'CLM GNT 2017-07-27 changed.xls')
        port_values = {}
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Sum.')
        read_portfolio_summary(ws, port_values)
        self.assertEqual(port_values['date'], datetime.datetime(2017,7,27))
        self.assertEqual(port_values['portfolio_id'], '30003')
        self.assertAlmostEqual(port_values['number_of_units'], 77954611.1195737)
        self.assertAlmostEqual(port_values['unit_price'], 18.433547)
        self.assertAlmostEqual(port_values['nav'], 1436979977.22)



    def test_bal(self):
        filename = os.path.join(get_current_path(), 'samples', 'CLM BAL 2017-07-27.xls')
        port_values = {}
        output_dir = os.path.join(get_current_path(), 'samples')
        open_bal(filename, port_values, output_dir, 'bal')



    def test_bal2(self):
        filename = os.path.join(get_current_path(), 'samples', 'CLM GNT 2017-07-27 changed.xls')
        port_values = {}
        output_dir = os.path.join(get_current_path(), 'samples')
        open_bal(filename, port_values, output_dir, 'gnt')
