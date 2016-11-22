# coding=utf-8
# 
# opens the expense worksheet.
# 

from xlrd import open_workbook
from xlrd.xldate import xldate_as_datetime
from DIF.utility import logger, get_datemode, retrieve_or_create
from DIF.open_holding import is_empty_cell
from DIF.open_summary import read_date, find_cell_string, get_portfolio_date



class InvalidExpenseItem(Exception):
	"""
	An error condition when a row is found to not containing the right
	cell values for an expense item.
	"""
	pass



class ExpenseTotalNotMatch(Exception):
	"""
	To indicate when the sum of expense items does not equal to the
	sub total read from the spread sheet.
	"""
	pass



# class InconsistentExpenseDate(Exception):
# 	pass



def read_expense(ws, port_values):
	"""
	Read the expenses worksheet. To use the function:

	expenses = port_values['expense']
	for expense_item in expenses:
		... expense_item['date']...
		keys: date, description, amount, currency, 
		exchange_rate, hkd_equivalent
	"""
	row = find_cell_string(ws, 0, 0, 'Valuation Period :')
	# expense_date = read_date(ws, row, 1)
	n = find_cell_string(ws, row, 0, 'Value Date')
	row = row + n

	fields = read_expense_fields(ws, row)
	row = row + 1

	while (is_blank_line(ws, row, 9)):	# skip blank lines
		row = row + 1

	expenses = retrieve_or_create(port_values, 'expense')
	while (row < ws.nrows):
		try:
			read_expense_item(ws, row, fields, expenses)
		except InvalidExpenseItem:
			# this line is not a expense item, skip it
			pass

		row = row + 1
		if is_blank_line(ws, row, 9):	# end of the first expense section
			break

		# end of while loop

	while (is_blank_line(ws, row, 9)):	# skip blank lines
		row = row + 1

	expense_sub_total = read_expense_sub_total(ws, row)
	row = row + 1	# move to next line
	validate_expense_sub_total(expenses, expense_sub_total)

	while (is_blank_line(ws, row, 9)):	# skip blank lines
		row = row + 1

	# continue to read the next section of expense items (
	# the performance fee)
	while (row < ws.nrows):
		try:
			read_expense_item(ws, row, fields, expenses)
		except InvalidExpenseItem:
			# this line is not a expense item, skip it
			pass

		row = row + 1
		if is_blank_line(ws, row, 9):	# end of the expense section
			break

		# end of while loop

	while (is_blank_line(ws, row, 9)):	# skip blank lines
		row = row + 1

	# now read the sub total after performance fee is included
	expense_sub_total = read_expense_sub_total(ws, row)
	validate_expense_sub_total(expenses, expense_sub_total)



def is_blank_line(ws, row, n_cells):
	"""
	Tell whether the row is empty in the first n cells.
	"""
	for i in range(n_cells):
		if not is_empty_cell(ws, row, i):
			return False

	return True



def read_expense_fields(ws, row):
	"""
	Read the data fields for an expense position
	"""
	fields = []

	field_mapping = {'Value Date':'value_date', 'Description':'description', 
						'Amount':'amount', 'CCY':'currency', 'Rate':'fx_rate', 
						'HKD Equiv.':'hkd_equivalent', '':'empty_field'}

	for column in range(9):	# read up to column I
		
		cell_value = ws.cell_value(row, column)
		if not isinstance(cell_value, str):	# data field name needs to
											# be string
			logger.error('read_expense_fields(): invalid expense field: {0}'.
							format(cell_value))
			raise ValueError('expense field not a string')

		try:
			fld = field_mapping[str.strip(cell_value)]
		except KeyError:
			logger.error('read_expense_fields(): unexpected expense field: {0}'.
							format(cell_value))
			raise ValueError('unexpected expense field')

		fields.append(fld)


	return fields



def read_expense_item(ws, row, fields, expenses):
	"""
	Read the row and try to create an expense item based on content
	in the row. If successful, append it to the expenses list,
	if not, an InvalidExpenseItem exception is thrown.

	If amount of the expense is 0, then no expense item is created.
	"""

	exp_item = {}	# the expense item
	column = -1
	for fld in fields:
		column = column + 1
		if fld == 'empty_field':
			continue

		cell_value = ws.cell_value(row, column)

		if fld in ['value_date', 'amount', 'fx_rate', 'hkd_equivalent']:
			if not isinstance(cell_value, float):
				logger.error('read_expense_item(): field {0} is not float, value = {1}'.
							format(fld, cell_value))
				raise InvalidExpenseItem('invalid field type')

			if fld == 'value_date':
				datemode = get_datemode()
				exp_item[fld] = xldate_as_datetime(cell_value, datemode)
			elif fld == 'amount' and cell_value == 0:	# stop reading
				logger.warning('read_expense_item(): amount is 0 at row {0}'.
								format(row))
				return

			else:
				exp_item[fld] = cell_value

		elif fld in ['description', 'currency']:
			if not isinstance(cell_value, str):
				logger.error('read_expense_item(): field {0} is not string, value = {1}'.
							format(fld, cell_value))
				raise InvalidExpenseItem('invalid field type')

			exp_item[fld] = str.strip(cell_value)

		else:	# unexpected field type
			logger.error('read_expense_item(): field {0} is unexpected'.format(fld))
			raise InvalidExpenseItem('unexpected field type')

		# end of for loop

	expenses.append(exp_item)



def read_expense_sub_total(ws, row):
	sub_total = ws.cell_value(row, 8)
	if isinstance(sub_total, float):
		return sub_total
	else:
		logger.error('read_expense_sub_total(): subtotal is not a float, value = {0}'.
						format(sub_total))
		raise ValueError('invalid subtotal type')



def validate_expense_sub_total(expenses, expense_sub_total):
	"""
	Retrieve expense items from the expenses list, sum them up and
	compare to the expense_sub_total.
	"""
	hkd_expenses = [exp_item['hkd_equivalent'] for exp_item in expenses]
	if abs(sum(hkd_expenses) - expense_sub_total) < 0.000001:
		pass	# alright, do nothing

	else:
		logger.error('validate_expense_sub_total(): sum of expenses {0} does not match the sub total {1}'.
						format(sum(hkd_expenses), expense_sub_total))
		raise ExpenseTotalNotMatch()