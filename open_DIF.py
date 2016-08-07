# coding=utf-8
# 
# This file is used to parse the diversified income fund excel files from
# trustee, read the necessary fields and save into a csv file for
# reconciliation with Advent Geneva.
#
# To use it, first we create an instance of the DIF class based on an
# input excel file:
#
# try:
# 	d = DIF(fn)	# fn is the input file
# except Exception:
#	... something wrong ...
# else:
# 	... continue processing ...
#
# then we can query the different attributes of the portfolio, like
# the following:
#
# 

from xlrd import open_workbook
from xlrd import XL_CELL_EMPTY, XL_CELL_DATE, XL_CELL_ERROR, XL_CELL_BLANK
from xlrd.xldate import xldate_as_datetime
import xlrd



def open_excel(file_name):
	"""
	Open the excel file, populate portfolio values into a dictionary.
	"""
	wb = open_workbook(filename=file_name)

	# open sheet for portfolio summary
	ws = wb.sheet_by_name('Portfolio Sum.')

	port_values = {}
	read_portfolio_summary(ws, port_values)



def read_portfolio_summary(ws, port_values, datemode=0):
	"""
	Read the content of the worksheet containing portfolio summary, iterate
	through all its rows and columns to populate some portfolio values.
	"""
	if isinstance(ws, xlrd.sheet.Sheet) == False:
		raise TypeError('read_portfolio_summary():Not a worksheet object')

	for row in range(ws.nrows):
			
		# search the first column
		cell_value = ws.cell_value(row, 0)
		cell_type = ws.cell_type(row, 0)

		if (cell_value == 'Valuation Period : From'):
			# the date is in this row, column B
			cell_value = ws.cell_value(row, 1)
			cell_type = ws.cell_type(row, 1)
			if (cell_type == XL_CELL_DATE):	# it is a date in Excel, now convert
											# it to python datetime object
				print(xldate_as_datetime(cell_value, datemode))
				break
			else:							# it is not of 'date' format,
											# something must be wrong
				raise TypeError('read_portfolio_summary():cell {0},{1} should be in excel date format'.format(row, 1))
