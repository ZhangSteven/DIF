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
import datetime



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
		date = cash_account['date']			# retrieve date
		balance = cash_account['balance']	# retrieve balance
		currency = cash_account['currency']	# retrieve currency
		account_type = cash_account['account_type']			# retrieve account type
		fx_rate = cash_account['fx_rate']	# retrieve FX rate to HKD
		HKD_equivelant = cash_account['hkd_equivalent']	# retrieve amount in HKD
		print(bank, date, account_num, account_type, currency, 
				balance, fx_rate, HKD_equivelant)



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



def convert_to_date(date_string, fmt='dd/mm/yyyy'):
	"""
	Convert a date string to a Python datetime.date object in the user 
	defined format.
	"""
	if fmt=='dd/mm/yyyy':
		dates = date_string.split('/')
		if (len(dates) == 3):
			try:
				dates_int = [int(d) for d in dates]
				the_date = datetime.date(dates_int[2], dates_int[1], dates_int[0])
				return the_date
			except Exception as e:
				# some thing wrong in the conversion process
				raise ValueError('convert_to_date(): invalid date_string: {0}'.format(date_string))
		else:
			raise ValueError('convert_to_date(): invalid date_string: {0}'.format(date_string))
	
	else:
		# format not handled
		raise ValueError('convert_to_date(): invalid format: {0}'.format(fmt))
