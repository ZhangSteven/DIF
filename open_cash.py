# coding=utf-8
# 
# Read the cash sections of the excel file from trustee.
#
# 

from xlrd import open_workbook
from xlrd.xldate import xldate_as_datetime
import xlrd
import datetime
from DIF.utility import logger, retrieve_or_create



# def open_excel_cash(file_name):
# 	"""
# 	Open the excel file, populate portfolio values into a dictionary.
#
#	For testing only.
# 	"""
# 	wb = open_workbook(filename=file_name)

# 	port_values = {}

# 	# find sheets that contain cash
# 	sheet_names = wb.sheet_names()
# 	for sn in sheet_names:
# 		if len(sn) > 4 and sn[-4:] == '-BOC':
# 			print('read from sheet {0}'.format(sn))
# 			ws = wb.sheet_by_name(sn)
# 			read_cash(ws, port_values)

# 	# show cash accounts
# 	show_cash_accounts(port_values)



def read_cash(ws, port_values, datemode=0):
	"""
	Read the worksheet with cash information. To retrieve cash information, 
	we do:

	cash_accounts = port_values['cash_accounts']	# get all cash accounts

	for cash_account in cash_accounts:
		
		bank = cash_account['bank']			# retrieve bank name
		account_num = cash_account['account_num']	# retrieve account number
		date = cash_account['date']			# retrieve date
		balance = cash_account['balance']	# retrieve balance
		currency = cash_account['currency']	# retrieve currency
		account_type = cash_account['account_type']	# retrieve account type
		fx_rate = cash_account['fx_rate']	# retrieve FX rate to HKD
		HKD_equivalent = cash_account['hkd_equivalent']	# retrieve amount in HKD
		
	"""
	logger.debug('in read_cash()')

	cash_accounts = retrieve_or_create(port_values, 'cash_accounts')

	def get_value(row, column=1):
		"""
		Define this local function to retrieve value for each property of
		a cash account, the information is either in column B or C.
		"""
		cell_type = ws.cell_type(row, column)
		if cell_type == xlrd.XL_CELL_EMPTY or cell_type == xlrd.XL_CELL_BLANK:
			# if this column is empty, return value in next column
			return ws.cell_value(row, column+1)
		else:
			return ws.cell_value(row, column)

	# to store cash account information read from this worksheet
	this_account = {}
	cash_accounts.append(this_account)

	for row in range(ws.nrows):
				
		# search the first column
		cell_value = ws.cell_value(row, 0)
		cell_type = ws.cell_type(row, 0)

		if (isinstance(cell_value, str)):

			if cell_value.startswith('Bank'):
				this_account['bank'] = get_value(row)

			elif cell_value.startswith('Account No.'):
				this_account['account_num'] = get_value(row)

			elif cell_value.startswith('Account Type'):
				this_account['account_type'] = get_value(row)
				
			elif cell_value.startswith('Valuation Period'):
				date_string = get_value(row, 2)
				this_account['date'] = xldate_as_datetime(date_string, datemode)

			elif cell_value.startswith('Account Currency'):
				this_account['currency'] = get_value(row)

			elif cell_value.startswith('Account Balance'):
				this_account['balance'] = get_value(row)

			elif cell_value.startswith('Exchange Rate'):
				this_account['fx_rate'] = get_value(row)

			elif cell_value.startswith('HKD Equiv'):
				this_account['hkd_equivalent'] = get_value(row)

	logger.debug('out of read_cash()')



def convert_to_date(date_string, fmt='dd/mm/yyyy'):
	"""
	Convert a string to a Python datetime.date object with the format
	specified.
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
				logger.exception('convert_to_date(): invalid date_string: {0}'.format(date_string))
				raise ValueError('convert_to_date()')
		else:
			logger.exception('convert_to_date(): invalid date_string: {0}'.format(date_string))
			raise ValueError('convert_to_date()')
	
	else:
		# format not handled
		logger.exception('convert_to_date(): invalid date format: {0}'.format(fmt))
		raise ValueError('convert_to_date()')