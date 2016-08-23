# coding=utf-8
# 
# opens the expense worksheet.
# 

from xlrd import open_workbook
from xlrd.xldate import xldate_as_datetime
from DIF.utility import logger, get_datemode, retrieve_or_create
from DIF.open_holding import is_empty_cell



def read_expense(ws, port_values):
	"""
	Read the expenses worksheet. To use the function:

	expenses = port_values['expense']
	for expense_item in expenses:
		... expense_item['date']...
		keys: date, description, amount, currency, 
		exchange_rate, hkd_equivalent
	"""
	row = 0

	n, expense_date = read_date(ws, row)
	row = row + n

	while row < ws.nrows:
		cell_value = ws.cell_value(row, 0)
		
		row = row + 1	# end of while loop
	
	fields = read_expense_fields(ws, row)


	expenses = retrieve_or_create(port_values, 'expense')

	while (row < ws.nrows):
		while (is_blank_line(ws, row)):
			row = row + 1

		try:
			read_expense_item(ws, row, expenses)
		except (ValueError, TypeError):
			# this line is not a expense item, skip it
			pass

		row = row + 1
		# end of while loop


	

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
						'HKD Equiv.':'hkd_equivalent'}

	for column in range(9):	# read up to column I
		if is_empty_cell(ws, row, column):
			fld = 'empty_field'
			fields.append(fld)
			continue

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