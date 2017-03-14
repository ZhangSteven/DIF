# coding=utf-8
# 
# This file is used to parse the diversified income fund excel files from
# trustee, read the necessary fields and save into a csv file for
# reconciliation with Advent Geneva.
#

from xlrd import open_workbook
from DIF.open_cash import read_cash
from DIF.open_summary import read_portfolio_summary, get_portfolio_date
from DIF.open_holding import read_holding
from DIF.open_expense import read_expense
from DIF.utility import logger, get_input_directory
from investment_lookup.id_lookup import get_investment_Ids
import csv, os



class InconsistentValue(Exception):
	pass

class InvalidDatetimeFormat(Exception):
	pass

class InconsistentExpenseDate(Exception):
	"""
	To indicate the expense item date is not the same as the portfolio
	value date.
	"""
	pass

class InvalidTickerFormat(Exception):
	pass



def open_dif(file_name, port_values, output_dir=get_input_directory()):
	"""
	Open the excel file of the DIF fund. Read its cash positions, holdings,
	expenses, calculate its nav and verify it with the nav from the excel.
	"""
	wb = open_workbook(filename=file_name)

	ws = wb.sheet_by_name('Portfolio Sum.')
	read_portfolio_summary(ws, port_values)
	
	# find sheets that contain cash
	sheet_names = wb.sheet_names()
	for sn in sheet_names:
		if len(sn) > 4 and sn[-4:] == '-BOC':
		    ws = wb.sheet_by_name(sn)
		    read_cash(ws, port_values)
	
	ws = wb.sheet_by_name('Portfolio Val.')
	read_holding(ws, port_values)
	
	ws = wb.sheet_by_name('Expense Report')
	read_expense(ws, port_values)

	validate_expense_date(port_values)

	# make sure the holding and cash are read correctly
	validate_cash_and_holding(port_values)

	# output the cash and holdings into csv files.
	return write_csv(port_values, output_dir)



def validate_expense_date(port_values):
	"""
	See if the date of the expense items is the same as the portfolio
	value date.
	"""
	port_date = get_portfolio_date(port_values)
	expenses = port_values['expense']
	for exp_item in expenses:
		if (exp_item['value_date'] == port_date):
			pass
		else:
			logger.error('expense date does not match: expense item: {0}, date {1}, portfolio date {2}'.
							format(exp_item['description'], exp_item['value_date'], port_date))
			raise InconsistentExpenseDate()



def validate_cash_and_holding(port_values):
	"""
	Calculate subtotal of cash, bond holdings and equity holdings, then 
	compare to the value from the excel file.

	The difference used in testing (0.01 for cash, 0.2 for bond and 0.1
	for equity) are based on experience. Because we find these numbers are
	'just nice' to pass the test, if they are too big, then there is no point
	to do verfication, if too small, then it will make some trustee excels fail.
	Maybe this is due to the rounding of actual number before they are input 
	to excel.
	"""
	cash_total = calculate_cash_total(port_values)
	if abs(cash_total - port_values['cash_total']) > 0.01:
		logger.error('validate_cash_and_holding(): calculated cash total {0} is inconsistent with that from file {1}'.
						format(cash_total, port_values['cash_total']))
		raise InconsistentValue

	fx_table = retrieve_fx(port_values)
	
	bond_holding = port_values['bond']
	bond_subtotal = calculate_bond_total(bond_holding, fx_table)
	if abs(bond_subtotal - port_values['bond_total']) > 0.2:
		logger.error('validate_cash_and_holding(): calculated bond total {0} is inconsistent with that from file {1}'.
						format(bond_subtotal, port_values['bond_total']))
		raise InconsistentValue

	equity_holding = port_values['equity']
	equity_subtotal = calculate_equity_total(equity_holding, fx_table)
	if abs(equity_subtotal - port_values['equity_total']) > 0.1:
		logger.error('validate_cash_and_holding(): calculated equity total {0} is inconsistent with that from file {1}'.
						format(equity_subtotal, port_values['equity_total']))
		raise InconsistentValue



def calculate_cash_total(port_values):
	total = 0
	cash_accounts = port_values['cash_accounts']
	for cash_account in cash_accounts:
		total = total + cash_account['hkd_equivalent']

	return total



def calculate_bond_total(bond_holding, fx_table):
	"""
	capital repayment needs to be taken into account.
	"""
	total = 0
	for bond in bond_holding:
		fx = fx_table[bond['currency']]
		amount = bond['par_amount']/100
		if amount == 0:
			continue

		try:
			local_currency_total = amount * bond['price']
		except KeyError:	# 'price' is not there, then it must be HTM
			local_currency_total = amount * bond['amortized_cost']

		total = total + fx*(local_currency_total + bond['accrued_interest'])

	return total



def calculate_equity_total(equity_holding, fx_table):
	"""
	preferred shares amount should be divided by 100
	"""
	total = 0
	for equity in equity_holding:
		fx = fx_table[equity['currency']]
		amount = equity['number_of_shares']
		if amount == 0:
			continue

		# if not 'listed_location' in equity:	# it's preferred shares
		# 	amount = amount /100

		# total = total + fx * amount * equity['price']

		total = total + fx*equity['market_value']

	return total



def retrieve_fx(port_values):
	fx_table = {}
	cash_accounts = port_values['cash_accounts']
	for cash_account in cash_accounts:
		fx_table[cash_account['currency']] = cash_account['fx_rate']

	return fx_table



def create_csv_file_name(date_string, output_dir, file_suffix):
	"""
	Create the output csv file name based on the date string, as well as
	the file suffix: cash, afs_positions, or htm_positions
	"""
	csv_file = output_dir + '\\DIF_' + date_string + '_' + file_suffix + '.csv'
	return csv_file



def write_csv(port_values, output_dir=get_input_directory()):
	"""
	Write cash and holdings into csv files.
	"""	
	cash_file = write_cash_csv(port_values, output_dir)
	htm_file = write_htm_holding_csv(port_values, output_dir)
	afs_file = write_afs_holding_csv(port_values, output_dir)
	return [cash_file, htm_file, afs_file]



def write_cash_csv(port_values, output_dir):
	portfolio_date = get_portfolio_date(port_values)
	portfolio_date = convert_datetime_to_string(portfolio_date)
	cash_file = create_csv_file_name(portfolio_date, output_dir, 'cash')
	logger.debug('write_cash_csv(): {0}'.format(cash_file))

	with open(cash_file, 'w', newline='') as csvfile:
		file_writer = csv.writer(csvfile, delimiter='|')

		cash_accounts = port_values['cash_accounts']

		fields = ['account_type', 'account_num', 'currency', 'balance', 
					'fx_rate', 'hkd_equivalent']

		portfolio_date = get_portfolio_date(port_values)
		portfolio_date = convert_datetime_to_string(portfolio_date)

		file_writer.writerow(['portfolio', 'custodian', 'date'] + fields)
		for cash_account in cash_accounts:
			row = ['19437', 'BOCHK', portfolio_date]
			for fld in fields:
				item = cash_account[fld]
				if fld == 'date':
					item = portfolio_date
				row.append(item)

			file_writer.writerow(row)

	return cash_file



def write_htm_holding_csv(port_values, output_dir):
	"""
	Output the HTM positions
	"""
	portfolio_date = get_portfolio_date(port_values)
	portfolio_date = convert_datetime_to_string(portfolio_date)
	holding_file = create_csv_file_name(portfolio_date, output_dir, 'htm_positions')
	logger.debug('write_htm_holding_csv(): {0}'.format(holding_file))
		
	with open(holding_file, 'w', newline='') as csvfile:
		file_writer = csv.writer(csvfile, delimiter='|')
		bond_holding = port_values['bond']

		# pick all fields that HTM bond have
		fields = ['name', 'currency', 'accounting_treatment', 'par_amount', 
					'is_listed', 'listed_location', 'fx_on_trade_day', 
					'coupon_rate', 'coupon_start_date', 'maturity_date', 
					'average_cost', 'amortized_cost', 'book_cost', 
					'interest_bought', 'amortized_value', 'accrued_interest', 
					'amortized_gain_loss', 'fx_gain_loss']

		file_writer.writerow(['portfolio', 'date', 'custodian', 'geneva_investment_id', 
								'isin', 'bloomberg_figi'] + fields)
		
		for bond in bond_holding:
			if bond['par_amount'] == 0 or bond['accounting_treatment'] != 'HTM':
				continue

			row = ['19437', portfolio_date, 'BOCHK']
			investment_ids = get_investment_Ids('19437', 'ISIN', bond['isin'], 
												bond['accounting_treatment'])
			for id in investment_ids:
				row.append(id)

			for fld in fields:
				try:	# HTM and Trading bonds have slightly different fields,
						# e.g, HTM bonds have amortized_cost while Trading
						# bonds have price
					item = bond[fld]
					if fld == 'coupon_start_date' or fld == 'maturity_date':
						item = convert_datetime_to_string(item)
				except KeyError:
					item = ''

				row.append(item)

			file_writer.writerow(row)

	return holding_file



def write_afs_holding_csv(port_values, output_dir):
	"""
	Output the AFS positions, including trading bond and equity.
	"""
	portfolio_date = get_portfolio_date(port_values)
	portfolio_date = convert_datetime_to_string(portfolio_date)
	holding_file = create_csv_file_name(portfolio_date, output_dir, 'afs_positions')
	logger.debug('write_afs_holding_csv(): {0}'.format(holding_file))

	with open(holding_file, 'w', newline='') as csvfile:
		file_writer = csv.writer(csvfile, delimiter='|')
		bond_holding = port_values['bond']
		equity_holding = port_values['equity']

		# pick fields that are common to equity and trading bond
		fields = ['ticker', 'isin', 'bloomberg_figi', 'name', 'currency', 
					'accounting_treatment', 'quantity', 'average_cost', 
					'price', 'book_cost', 'market_value', 'market_gain_loss', 
					'fx_gain_loss']
			
		file_writer.writerow(['portfolio', 'date', 'custodian'] + fields)
		
		for bond in bond_holding:
			if bond['par_amount'] == 0 or bond['accounting_treatment'] != 'Trading':
				continue

			row = ['19437', portfolio_date, 'BOCHK']

			for fld in fields:
				if fld == 'quantity':
					fld = 'par_amount'
				try:	
					item = bond[fld]
				except KeyError:
					item = ''

				row.append(item)

			file_writer.writerow(row)


		for equity in equity_holding:
			if equity['number_of_shares'] == 0 or equity['accounting_treatment'] != 'Trading':
				continue

			row = ['19437', portfolio_date, 'BOCHK']

			for fld in fields:
				if fld == 'quantity':
					fld = 'number_of_shares'
				try:	
					item = equity[fld]
					if fld == 'ticker':
						item = convert_to_BLP_ticker(item)
				except KeyError:
					item = ''

				row.append(item)

			file_writer.writerow(row)

	return holding_file



def convert_datetime_to_string(dt, fmt='yyyy-mm-dd'):
	"""
	convert a datetime object to string according to the 
	format.
	"""
	if fmt == 'yyyy-mm-dd':
		return '{0}-{1}-{2}'.format(dt.year, dt.month, dt.day)

	else:
		logger.error('convert_datetime_to_string(): invalid format {0}'.
						format(fmt))
		raise InvalidDatetimeFormat



def convert_to_BLP_ticker(ticker):
	"""
	Convert a ticker in trustee's format to Bloomberg ticker format. E.g.,

	H0939: 939 HK
	H1186: 1186 HK
	N0011: 11 HK
	N2388: 2388 HK

	H probaly means "H shares", N probably means "normal shares", so the rule
	is to remove the leading "H" or "N", then remove any leading zeros, then
	append "HK" to finish the conversion.
	"""
	if len(ticker) == 5 and ticker[0] in ['H', 'N']:
		ticker = ticker[1:]
		if ticker.isdigit():
			i = 0
			for char in ticker:
				if char == '0':
					i = i + 1
				else:
					break

			if i < len(ticker):
				return ticker[i:] + ' HK'

	logger.error('convert_to_BLP_ticker(): invalid ticker format {0}'.format(ticker))
	raise InvalidTickerFormat



# we can execute the open_dif() from command line with an input file name
if __name__ == '__main__':
	import sys
	if len(sys.argv) < 2:
		print('use python open_dif.py <input_file>')
		sys.exit(1)

	filename = get_input_directory() + '\\' + sys.argv[1]
	if not os.path.exists(filename):
		print('{0} does not exist'.format(filename))
		sys.exit(1)

	port_values = {}
	try:
		open_dif(filename, port_values)
	except:
		logger.exception('open_dif():')
		print('something goes wrong, check log file.')
	else:
		print('OK')
