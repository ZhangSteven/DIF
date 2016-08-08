# coding=utf-8
# 
# Read the portfolio summary section of the excel from trustee.
# 

from xlrd import open_workbook
from xlrd.xldate import xldate_as_datetime
import xlrd
import datetime



def open_excel_summary(file_name):
	"""
	Open the excel file, populate portfolio values into a dictionary.
	"""
	try:
		wb = open_workbook(filename=file_name)
	except Exception as e:
		# do some logging here
		raise

	# the place holder for DIF portfolio information
	port_values = {}

	# read portfolio summary
	try:
		ws = wb.sheet_by_name('Portfolio Sum.')
		read_portfolio_summary(ws, port_values)
	except Exception as e:
		# do some logging here
		raise

	

def read_portfolio_summary(ws, port_values, datemode=0):
	"""
	Read the content of the worksheet containing portfolio summary, iterate
	through all its rows and columns to populate some portfolio values.
	"""
	for row in range(ws.nrows):
			
		# search the first column
		cell_value = ws.cell_value(row, 0)
		cell_type = ws.cell_type(row, 0)

		if (cell_value == 'Valuation Period : From'):
			# the date is in this row, column B
			cell_value = ws.cell_value(row, 1)
			cell_type = ws.cell_type(row, 1)
		
			if isinstance(cell_value, float):
					# Excel stores 'date' formatted cell as a float number, we need
					# to convert it to a python datetime.datetime object.
					#
					# But sometimes, a date is formatted as "text" in a cell, then
					# it will be read as a string, in this case, we need to handle it
					# differently.
					print(xldate_as_datetime(cell_value, datemode))
					break

			else:							
				raise TypeError('read_portfolio_summary():cell {0},{1} not a valid date: {2}'
									.format(row, 1, cell_value))



