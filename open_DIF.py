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
	try:
		wb = open_workbook(filename=file_name)
	except Exception as e:
		# do some logging here
		raise

	# the place holder for DIF portfolio
	port_values = {}

	# read portfolio summary
	try:
		ws = wb.sheet_by_name('Portfolio Sum.')
		read_portfolio_summary(ws, port_values)
	except Exception as e:
		# do some logging here
		raise

	# read cash account information
	sheet_names = wb.sheet_names()
	for sn in sheet_names:
		if len(sn) > 4 and sn[-4:] == '-BOC':
			print('read from sheet {0}'.format(sn))
			ws = wb.sheet_by_name(sn)

			try:
				read_cash(ws, port_values)
			except Exception as e:
				# do some logging here
				raise

	# verify we have read the correct value
	show_cash_accounts(port_values)

	

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
		HKD_equivalent = cash_account['hkd_equivalent']	# retrieve amount in HKD
		
	"""
	if 'cash_accounts' in port_values:
		cash_accounts = port_values['cash_accounts']
	else:
		cash_accounts = {}
		port_values['cash_accounts'] = cash_accounts

	def get_value(row, column=1):
		"""
		Define this local function to retrieve value for each property of
		a cash account, the information is either in column B or C.
		"""
		cell_type = ws.cell_type(row, column)
		if cell_type == XL_CELL_EMPTY or cell_type == XL_CELL_BLANK:
			# if this column is empty, return value in next column
			return ws.cell_value(row, column+1)
		else:
			return ws.cell_value(row, column)

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

				elif len(cell_value) > 12 and cell_value[:12] == 'Account Type':
					this_account['account_type'] = get_value(row)
					
				elif len(cell_value) > 16 and cell_value[:16] == 'Valuation Period':
					date_string = get_value(row, 2)
					this_account['date'] = xldate_as_datetime(date_string, datemode)

				elif len(cell_value) > 16 and cell_value[:16] == 'Account Currency':
					this_account['currency'] = get_value(row)

				elif len(cell_value) > 15 and cell_value[:15] == 'Account Balance':
					this_account['balance'] = get_value(row)

				elif len(cell_value) > 13 and cell_value[:13] == 'Exchange Rate':
					this_account['fx_rate'] = get_value(row)

				elif len(cell_value) > 9 and cell_value[:9] == 'HKD Equiv':
					this_account['hkd_equivalent'] = get_value(row)



def show_cash_accounts(port_values):
	cash_accounts = port_values['cash_accounts']

	for id in cash_accounts:
		cash_account = cash_accounts[id]	# use account_number as key
		
		bank = cash_account['bank']			# retrieve bank name
		account_num = cash_account['account_num']	# retrieve account number
		date = cash_account['date']			# retrieve date
		balance = cash_account['balance']	# retrieve balance
		currency = cash_account['currency']	# retrieve currency
		account_type = cash_account['account_type']			# retrieve account type
		fx_rate = cash_account['fx_rate']	# retrieve FX rate to HKD
		HKD_equivelant = cash_account['hkd_equivalent']	# retrieve amount in HKD
		print(bank, date, account_num, account_type, currency, 
				balance, fx_rate, HKD_equivelant)