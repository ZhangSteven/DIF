# coding=utf-8
"""
Test methods open_expense.py

"""

import unittest2
import datetime
from xlrd import open_workbook
from DIF.utility import get_current_path
from DIF.open_summary import read_date, find_cell_string
from DIF.open_expense import read_expense_fields, read_expense_item, \
                            InvalidExpenseItem, ExpenseTotalNotMatch, \
                            validate_expense_sub_total, read_expense_sub_total, \
                            is_blank_line, read_expense, validate_expense_date, \
                            InconsistentExpenseDate

class TestExpense(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestExpense, self).__init__(*args, **kwargs)

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



    def test_read_date(self):
        """
        Test the read_date() function.
        """
        filename = get_current_path() + '\\samples\\expense_sample.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Expense Report')

        d = read_date(ws, 5, 1)
        self.assertEqual(d, datetime.datetime(2015,12,10))



    def test_expense_fields(self):

        filename = get_current_path() + '\\samples\\expense_sample.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Expense Report')

        # find the row where the fields are
        row = find_cell_string(ws, 0, 0, 'Value Date')
        fields = read_expense_fields(ws, row)
        fields_to_be = ['value_date', 'description', 'empty_field',
                        'empty_field', 'amount', 'currency',
                        'fx_rate', 'empty_field', 'hkd_equivalent']

        self.assertEqual(len(fields), 9)
        i = 0
        for fld in fields:
            self.assertEqual(fld, fields_to_be[i])
            i = i + 1


    def test_read_expense_item(self):
        filename = get_current_path() + '\\samples\\expense_sample.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Expense Report')

        row = find_cell_string(ws, 0, 0, 'Value Date')
        fields = read_expense_fields(ws, row)
        expenses = []

        row = 11    # read the first expense at A12
        read_expense_item(ws, row, fields, expenses)
        self.assertEqual(len(expenses), 1)
        self.validate_expense_item(expenses[0])



    def test_read_expense_item2(self):
        filename = get_current_path() + '\\samples\\expense_sample.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Expense Report')

        row = find_cell_string(ws, 0, 0, 'Value Date')
        fields = read_expense_fields(ws, row)
        expenses = []

        row = 13    # read the first expense at A14
                    # no expense item will be read
        read_expense_item(ws, row, fields, expenses)
        self.assertEqual(len(expenses), 0)



    def test_read_expense_item_error(self):
        filename = get_current_path() + '\\samples\\expense_sample.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Expense Report')

        row = find_cell_string(ws, 0, 0, 'Value Date')
        fields = read_expense_fields(ws, row)
        expenses = []

        row = 14    # read the expense at A15, suppose to generate error
        with self.assertRaises(InvalidExpenseItem):
            read_expense_item(ws, row, fields, expenses)
        


    def test_validate_expense_subtotal(self):
        filename = get_current_path() + '\\samples\\expense_sample.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Expense Report')

        row = find_cell_string(ws, 0, 0, 'Value Date')
        fields = read_expense_fields(ws, row)

        row = 11    # expense items start at A12
        expenses = []
        while (row < ws.nrows):
            try:
                read_expense_item(ws, row, fields, expenses)
            except InvalidExpenseItem:
                # this line is not a expense item, skip it
                pass

            row = row + 1
            if is_blank_line(ws, row, 9):   # end of the first expense section
                break

            # end of while loop

        row = 27    # read from I28
        expense_sub_total = read_expense_sub_total(ws, row)
        validate_expense_sub_total(expenses, expense_sub_total)

        expense_sub_total = expense_sub_total + 0.01    # create an error
        with self.assertRaises(ExpenseTotalNotMatch):
            validate_expense_sub_total(expenses, expense_sub_total)



    def test_read_expense(self):
        filename = get_current_path() + '\\samples\\expense_sample.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Expense Report')

        port_values = {}
        read_expense(ws, port_values)
        expenses = port_values['expense']
        self.assertEqual(len(expenses), 9)  # should have 9 items, those
                                            # with 0 amount are not included
        self.validate_expense_item(expenses[0])
        self.validate_expense_item2(expenses[8])
        expense_sub_total = 64144.11
        validate_expense_sub_total(expenses, expense_sub_total)



    def test_read_expense2(self):
        filename = get_current_path() + '\\samples\\expense_sample2.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Expense Report')

        port_values = {}
        read_expense(ws, port_values)
        expenses = port_values['expense']
        self.assertEqual(len(expenses), 10)  # should have 9 items, those
                                            # with 0 amount are not included

        self.validate_expense_item(expenses[0])
        self.validate_expense_item3(expenses[9])
        expense_sub_total = 65644.11
        validate_expense_sub_total(expenses, expense_sub_total)


    def test_validate_expense_date(self):
        filename = get_current_path() + '\\samples\\expense_sample2.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Expense Report')

        port_values = {}
        read_expense(ws, port_values)
        expenses = port_values['expense']

        d = datetime.datetime(2015,12,10)
        validate_expense_date(expenses, d)

        # change the date in one of the expense items
        exp_item = expenses[1]
        exp_item['value_date'] = d + datetime.timedelta(days=1)
        with self.assertRaises(InconsistentExpenseDate):
            validate_expense_date(expenses, d)



    def validate_expense_item(self, expense_item):
        """
        Validate the first expense item in expense_sample.xls
        """
        self.assertEqual(len(expense_item), 6)  # 6 data fields
        self.assertEqual(expense_item['value_date'], datetime.datetime(2015,12,10))
        self.assertEqual(expense_item['description'], 'Setup Fee')
        self.assertAlmostEqual(expense_item['amount'], 602.12)
        self.assertEqual(expense_item['currency'], 'HKD(港幣)')
        self.assertEqual(expense_item['fx_rate'], 1.0)
        self.assertAlmostEqual(expense_item['hkd_equivalent'], 602.12)



    def validate_expense_item2(self, expense_item):
        """
        Validate the last expense item in expense_sample.xls
        """
        self.assertEqual(len(expense_item), 6)  # 6 data fields
        self.assertEqual(expense_item['value_date'], datetime.datetime(2015,12,10))
        self.assertEqual(expense_item['description'], 'SFC Annual Fee')

        # special case, in the excel, the amount is 16.4383561643836,
        # but the hkd equivalent is 16.44, they are not equal, maybe due
        # to human error.
        self.assertAlmostEqual(expense_item['amount'], 16.44, places=2)
        self.assertEqual(expense_item['currency'], 'HKD(港幣)')
        self.assertEqual(expense_item['fx_rate'], 1.0)
        self.assertAlmostEqual(expense_item['hkd_equivalent'], 16.44)



    def validate_expense_item3(self, expense_item):
        """
        Validate the performance fee expense item in expense_sample2.xls
        """
        self.assertEqual(len(expense_item), 6)  # 6 data fields
        self.assertEqual(expense_item['value_date'], datetime.datetime(2015,12,10))
        self.assertEqual(expense_item['description'], 'Performance Fee')
        self.assertAlmostEqual(expense_item['amount'], 1500)
        self.assertEqual(expense_item['currency'], 'HKD(港幣)')
        self.assertEqual(expense_item['fx_rate'], 1.0)
        self.assertAlmostEqual(expense_item['hkd_equivalent'], 1500)





        