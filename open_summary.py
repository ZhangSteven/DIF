# coding=utf-8
# 
# Read the portfolio summary section of the excel from trustee.
# 

from xlrd import open_workbook
from xlrd.xldate import xldate_as_datetime
import xlrd
import datetime, logging
from DIF.utility import logger, get_datemode



class CellNotFound(Exception):
	"""
	An error condition when we failed to find a cell in the worksheet.
	"""
	pass



# def open_excel_summary(file_name):
# 	"""
# 	Open the excel file, populate portfolio values into a dictionary.
# 	"""
# 	logger.debug('in open_excel_summary()')

# 	try:
# 		wb = open_workbook(filename=file_name)
# 	except Exception as e:
# 		logger.critical('DIF file {0} cannot be opened'.format(file_name))
# 		logger.exception('open_excel_summary()')
# 		raise

# 	# the place holder for DIF portfolio information
# 	port_values = {}

# 	# read portfolio summary
# 	try:
# 		sn = 'Portfolio Sum.'
# 		ws = wb.sheet_by_name(sn)
# 	except Exception as e:
# 		logger.critical('worksheet {0} cannot be opened'.format(sn))
# 		logger.exception('open_excel_summary()')
# 		raise

# 	try:
# 		read_portfolio_summary(ws, port_values)
# 	except Exception as e:
# 		logger.error('failed to populate portfolio summary.')
# 		raise

# 	show_portfolio_summary(port_values)
# 	logger.debug('out of open_excel_summary()')



def read_portfolio_summary(ws, port_values):
	"""
	Read the content of the worksheet containing portfolio summary, iterate
	through all its rows and columns to populate some portfolio values.
	"""
	logger.debug('in read_portfolio_summary()')

	row = find_cell_string(ws, 0, 0, 'Valuation Period :')
	# d = read_date(ws, row, 1)
	d = read_date(ws, row, 3)
	port_values['date'] = d

	# read the summary of cash and holdings
	n = read_cash_holding_total(ws, row, port_values)
	row = row + n

	n = find_cell_string(ws, row, 0, 'Total Units Held at this Valuation  Date')
	row = row + n 	# move to that row
	cell_value = ws.cell_value(row, 2)	# read value at column C
	populate_value(port_values, 'number_of_units', cell_value, row, 2)

	# the first 'unit price' is before performance fee,
	# so we do not use it
	n = find_cell_string(ws, row, 0, 'Unit Price')
	row = row + n

	n = find_cell_string(ws, row, 0, 'Net Asset Value')
	row = row + n
	cell_value = ws.cell_value(row, 9)	# read value at column C
	populate_value(port_values, 'nav', cell_value, row, 9)

	# the second 'unit price' after 'net asset value' is the
	# the one we want to use.
	n = find_cell_string(ws, row, 0, 'Unit Price')
	row = row + n
	cell_value = ws.cell_value(row, 2)	# read value at column C
	populate_value(port_values, 'unit_price', cell_value, row, 2)

	# for row in range(ws.nrows):
	# while row < ws.nrows:
			
	# 	# search the first column
	# 	cell_value = ws.cell_value(row, 0)
	# 	cell_type = ws.cell_type(row, 0)

	# 	if (cell_value.startswith('Total Units Held at this Valuation  Date')):
	# 		cell_value = ws.cell_value(row, 2)	# read value at column C
	# 		populate_value(port_values, 'number_of_units', cell_value, row, 2)

	# 	elif (cell_value.startswith('Unit Price')):
	# 		if count == 0:
	# 			# there are two cells in column A that shows 'Unit Price',
	# 			# but only the second cell contains the right value (after
	# 			# performance fee)
	# 			count = count + 1
	# 		else:
	# 			cell_value = ws.cell_value(row, 2)	# read value at column C
	# 			populate_value(port_values, 'unit_price', cell_value, row, 2)

	# 	elif (cell_value == 'Net Asset Value'):
	# 		cell_value = ws.cell_value(row, 9)
	# 		populate_value(port_values, 'nav', cell_value, row, 9)

	# 	row = row + 1
	# 	# end of while loop
			
	logger.debug('out of read_portfolio_summary()')



def populate_value(port_values, key, cell_value, row, column):
	"""
	For the number of units, nav and unit price, they have the same validation
	process, so we put it here.

	If cell_value is valid, assign it to the port_values dictionary. Otherwise
	throw an ValueError exception with the msg to indicate something is wrong.

	port_values	: the dictionary holding the portfolio values read from
					the excel.
	key			: needs to be a string, indicating the name of the value.
	"""
	logger.debug('in populate_value()')

	if (isinstance(cell_value, float)) and cell_value > 0:
		port_values[key] = cell_value
	else:
		logger.error('cell {0},{1} is not a valid {2}: {3}'
						.format(row, column, key, cell_value))
		raise ValueError(key)

	logger.debug('out of populate_value()')



def show_portfolio_summary(port_values):
	"""
	Show summary of the portfolio, read from the 'Portfolio Sum.' sheet.
	"""	
	for key in port_values:
		if key == 'nav':
			print('nav = {0}'.format(port_values['nav']))
		elif key == 'date':
			print('date = {0}'.format(port_values['date']))
		elif key == 'number_of_units':
			print('number_of_units = {0}'.format(port_values['number_of_units']))
		elif key == 'unit_price':
			print('unit_price = {0}'.format(port_values['unit_price']))



def read_date(ws, row, column):
	"""
	Find the date of valuation period.
	"""
	datemode = get_datemode()
	cell_value = ws.cell_value(row, column)
	try:
		d = xldate_as_datetime(cell_value, datemode)
	except:
		logger.error('read_date(): failed to convert value:{0} to date'.
						format(cell_value))
		raise

	return d



def find_cell_string(ws, row, column, cell_string):
	"""
	Starting from a given row, search in the give column, until the cell
	content starts with the cell_string.

	Returns how many more rows have been read besides the current row
	to find the cell.
	"""
	rows_read = 0

	while (row+rows_read < ws.nrows):
		cell_value = ws.cell_value(row+rows_read, 0)
		if isinstance(cell_value, str) and cell_value.startswith(cell_string):
			return rows_read

		rows_read = rows_read + 1
		# end of while loop

	
	# reached end of worksheet, but not found yet.
	logger.error('find_cell_string(): cell string {0} not found'.
					format(cell_string))
	raise CellNotFound('cell string not found')



def read_cash_holding_total(ws, row, port_values):
	"""
	Read the subtotal of cash, bond holding and equity holding.
	"""
	rows_read = find_cell_string(ws, row, 0, 'Current Portfolio')
	count = 0
	while row+rows_read < ws.nrows and count < 4:
		cell_value = ws.cell_value(row+rows_read, 0)
		if not isinstance(cell_value, str):
			continue

		count = count + 1	# assume we find one item
		target_value = ws.cell_value(row+rows_read, 7)	# column H
		
		if cell_value.startswith('Cash') and isinstance(target_value, float):
			port_values['cash_total'] = target_value
		elif cell_value.startswith('Debt Securities') and isinstance(target_value, float):
			debt_value = target_value
		elif cell_value.startswith('Debt Amortization') and isinstance(target_value, float):
			debt_amortization = target_value
		elif cell_value.startswith('Equities') and isinstance(target_value, float):
			port_values['equity_total'] = target_value
		else:
			count = count - 1	# item not found, reverse the count


		rows_read = rows_read + 1
		# end of while loop

	if count < 4:	# not all 4 sub totals found before end of file
		logger.error('read_cash_holding_total(): some subtotal is missing')
		raise CellNotFound

	port_values['bond_total'] = debt_value + debt_amortization
	return rows_read



def get_portfolio_date(port_values):
	"""
	Read the date of the summary.
	"""
	return port_values['date']