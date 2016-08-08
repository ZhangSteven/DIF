# coding=utf-8
# 
# Read the portfolio summary section of the excel from trustee.
# 

from xlrd import open_workbook
from xlrd.xldate import xldate_as_datetime
import xlrd

import datetime, logging

from DIF.utility import logger



def open_excel_summary(file_name):
	"""
	Open the excel file, populate portfolio values into a dictionary.
	"""
	logger.log(logging.DEBUG, 'in open_excel_summary()')

	try:
		wb = open_workbook(filename=file_name)
	except Exception as e:
		logger.log(logging.CRITICAL, 'DIF file {0} cannot be opened'.format(file_name))
		raise

	# the place holder for DIF portfolio information
	port_values = {}

	# read portfolio summary
	try:
		sn = 'Portfolio Sum.'
		ws = wb.sheet_by_name(sn)
	except Exception as e:
		logger.log(logging.CRITICAL, 'worksheet {0} cannot be opened'.format(sn))
		# logger.log(logging.CRITICAL, repr(e))	# seems doesn't log anything?
		raise

	try:
		read_portfolio_summary(ws, port_values)
	except Exception as e:
		logger.log(logging.ERROR, 'failed to populate portfolio summary.')
		raise

	show_portfolio_summary(port_values)
	logger.log(logging.DEBUG, 'out of open_excel_summary()')



def read_portfolio_summary(ws, port_values, datemode=0):
	"""
	Read the content of the worksheet containing portfolio summary, iterate
	through all its rows and columns to populate some portfolio values.
	"""
	logger.log(logging.DEBUG, 'in read_portfolio_summary()')

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
				raise TypeError('read_portfolio_summary():cell {0},{1} not a valid date: {2}'
									.format(row, 1, cell_value))

	logger.log(logging.DEBUG, 'out of read_portfolio_summary()')



def show_portfolio_summary(port_values):
	"""
	Show summary of the portfolio, read from the 'Portfolio Sum.' sheet.
	"""	
	for key in port_values:
		if key == 'nav':
			print('nav = {0}'.format(port_values['nav']))
		elif key == 'date':
			print('date = {0}'.format(port_values['date']))