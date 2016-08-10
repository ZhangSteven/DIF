# coding=utf-8
# 
# Read the portfolio summary section of the excel from trustee.
# 

from xlrd import open_workbook
from xlrd.xldate import xldate_as_datetime
import xlrd

import datetime, logging

from DIF.utility import logger



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



def read_portfolio_summary(ws, port_values, datemode=0):
	"""
	Read the content of the worksheet containing portfolio summary, iterate
	through all its rows and columns to populate some portfolio values.
	"""
	logger.debug('in read_portfolio_summary()')

	count = 0
	for row in range(ws.nrows):
			
		# search the first column
		cell_value = ws.cell_value(row, 0)
		cell_type = ws.cell_type(row, 0)

		if (cell_value == 'Valuation Period : From'):
			# the date is in this row, column B
			cell_value = ws.cell_value(row, 1)
			cell_type = ws.cell_type(row, 1)
		
			if isinstance(cell_value, float):
				# Excel stores 'date' formatted cell as a float number, so we
				# expect a float value here.
				#
				# But sometimes, a date is formatted as "text" in a cell, then
				# it will be read as a string, in this case, we need to handle it
				# differently.
				d = xldate_as_datetime(cell_value, datemode)
				port_values['date'] = d

			else:
				logger.error('cell {0},1 is not a valid date: {1}'
								.format(row, cell_value))
				raise ValueError('date')	# 'date' to indicate what's 
											# wrong, used in test code.

		elif (cell_value.startswith('Total Units Held at this Valuation  Date')):
			cell_value = ws.cell_value(row, 2)	# read value at column C
			populate_value(port_values, 'number_of_units', cell_value, row, 2)

		elif (cell_value.startswith('Unit Price')):
			if count == 0:
				# there are two cells in column A that shows 'Unit Price',
				# but only the second cell contains the right value (after
				# performance fee)
				count = count + 1
			else:
				cell_value = ws.cell_value(row, 2)	# read value at column C
				populate_value(port_values, 'unit_price', cell_value, row, 2)

		elif (cell_value == 'Net Asset Value'):
			cell_value = ws.cell_value(row, 9)
			populate_value(port_values, 'nav', cell_value, row, 9)
			
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