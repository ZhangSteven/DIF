"""
Test the open_bal() method to open the trustee Macau Balanced Fund.
"""

import unittest2
import os
from DIF.utility import get_current_path
from DIF.open_bal import open_bal



class TestBAL(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestBAL, self).__init__(*args, **kwargs)



    def test_consolidate_cash(self):
        filename = os.path.join(get_current_path(), 'samples', 'CLM BAL 2017-07-27.xls')
        port_values = {}
        output_dir = os.path.join(get_current_path(), 'samples')
        open_bal(filename, port_values, output_dir, 'bal')

