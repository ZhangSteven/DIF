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

	port_values = {}

	# find sheets that contain cash
	sheet_names = wb.sheet_names()
	for sn in sheet_names:
		if len(sn) > 4 and sn[-4:] == '-BOC':
			print('read from sheet {0}'.format(sn))
			ws = wb.sheet_by_name(sn)
			read_cash(ws, port_values)

	# show cash accounts
	cash_accounts = port_values['cash_accounts']

	for id in cash_accounts:
		cash_account = cash_accounts[id]	# use account_number as key
		
		bank = cash_account['bank']			# retrieve bank name
		account_num = cash_account['account_num']	# retrieve account number
		# date = cash_account['date']			# retrieve date
		balance = cash_account['balance']	# retrieve balance
		currency = cash_account['currency']	# retrieve currency
		# type = cash_account['type']			# retrieve account type
		# fx_rate = cash_account['fx_rate']	# retrieve FX rate to HKD
		# HKD_equivelant = cash_account['HKD_equivelant']	# retrieve amount in HKD
		print(account_num, currency, balance)

def read_cash(ws, port_values, datemode=0):
	"""
	Read the worksheet with cash information. To retrieve cash information, 
	we do:

	cash_accounts = port_values['cash_accounts']	# get all cash accounts

	for id in cash_accounts:
		cash_account = cash_accounts[id]	# use integer as key
		
		bank = cash_account['bank']			# retrieve bank name
		account_num = cash_account['account_num']	# retrieve account number
		date = cash_account['date']			# retrieve date
		balance = cash_account['balance']	# retrieve balance
		currency = cash_account['currency']	# retrieve currency
		account_type = cash_account['account_type']	# retrieve account type
		fx_rate = cash_account['fx_rate']	# retrieve FX rate to HKD
		HKD_equivelant = cash_account['HKD_equivelant']	# retrieve amount in HKD
		
	"""
	if 'cash_accounts' in port_values:
		cash_accounts = port_values['cash_accounts']
	else:
		cash_accounts = {}
		port_values['cash_accounts'] = cash_accounts

	def get_value(row):
		"""
		Define this local function to retrieve value for each property of
		a cash account, the information is either in column B or C.
		"""
		cell_type = ws.cell_type(row, 1)
		if cell_type == XL_CELL_EMPTY or cell_type == XL_CELL_BLANK:
			# column B is empty, return value in column C
			return ws.cell_value(row, 2)
		else:
			return ws.cell_value(row, 1)

	# to store cash account information read from this worksheet
	this_account = {}
	id = len(cash_accounts.keys()) + 1
	cash_accounts[id] = this_account

	for row in range(ws.nrows):
				
			# search the first column
			cell_value = ws.cell_value(row, 0)
			cell_type = ws.cell_type(row, 0)

			if (isinstance(cell_value, str)):
				if len(cell_value) > 4 and cell_value[:4] == 'Bank':
					this_account['bank'] = get_value(row)

				elif len(cell_value) > 11 and cell_value[:11] == 'Account No.':
					this_account['account_num'] = get_value(row)
					
				elif len(cell_value) > 16 and cell_value[:16] == 'Account Currency':
					this_account['currency'] = get_value(row)

				elif len(cell_value) > 15 and cell_value[:15] == 'Account Balance':
					this_account['balance'] = get_value(row)

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
