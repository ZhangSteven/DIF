# coding=utf-8
# 
# This file is used to parse the diversified income fund excel files from
# trustee, read the necessary fields and save into a csv file for
# reconciliation with Advent Geneva.
#

from xlrd import open_workbook
from DIF.open_cash import read_cash
from DIF.open_summary import read_portfolio_summary
from DIF.open_holding import read_holding
from DIF.open_expense import read_expense
from DIF.utility import logger, config, get_current_path
import csv



class InconsistentValue(Exception):
	pass



class InvalidDatetimeFormat(Exception):
	pass



def open_dif(file_name, port_values):
	"""
	Open the excel file of the DIF fund. Read its cash positions, holdings,
	expenses, calculate its nav and verify it with the nav from the excel.
	"""
	try:
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

		# make sure the holding and cash are read correctly
		validate_cash_and_holding(port_values)

		# output the cash, holdings into a csv file.
		write_csv(port_values)
	except:
		logger.exception('open_dif()')
		raise



def validate_cash_and_holding(port_values):
	"""
	Calculate subtotal of cash, bond holdings and equity holdings, then 
	compare to the value from the excel file.

	The difference used in testing (0.01 for cash, 0.05 for bond and 0.01
	for equity) are based on experience. Because we find these numbers are
	'just nice' to pass the test, if they are too big, then there is no point
	to do verfication, if too small, then it will make some excels fail.
	Maybe this is due to the rounding of actual number before they are input 
	to excel.
	"""
	cash_total = calculate_cash_total(port_values)
	if abs(cash_total - port_values['cash_total']) > 0.01:
		logger.error('validate_cash_holding(): calculated cash total {0} is inconsistent with that from file {1}'.
						format(cash_total, port_values['cash_total']))
		raise InconsistentValue

	fx_table = retrieve_fx(port_values)
	
	bond_holding = port_values['bond']
	bond_subtotal = calculate_bond_total(bond_holding, fx_table)
	if abs(bond_subtotal - port_values['bond_total']) > 0.05:
		logger.error('validate_cash_holding(): calculated bond total {0} is inconsistent with that from file {1}'.
						format(bond_subtotal, port_values['bond_total']))
		raise InconsistentValue

	equity_holding = port_values['equity']
	equity_subtotal = calculate_equity_total(equity_holding, fx_table)
	if abs(equity_subtotal - port_values['equity_total']) > 0.01:
		logger.error('validate_cash_holding(): calculated equity total {0} is inconsistent with that from file {1}'.
						format(equity_subtotal, port_values['equity_total']))
		raise InconsistentValue



def calculate_cash_total(port_values):
	total = 0
	cash_accounts = port_values['cash_accounts']
	for cash_account in cash_accounts:
		total = total + cash_account['hkd_equivalent']

	return total



# def calculate_holding_total(port_values):
# 	fx_table = retrieve_fx(port_values)
	
# 	bond_holding = port_values['bond']
# 	bond_subtotal = calculate_bond_total(bond_holding, fx_table)

# 	equity_holding = port_values['equity']
# 	equity_subtotal = calculate_equity_total(equity_holding, fx_table)

# 	return bond_subtotal + equity_subtotal
# 	# return bond_subtotal



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

		if not 'listed_location' in equity:	# it's preferred shares
			amount = amount /100

		total = total + fx * amount * equity['price']

	return total



def retrieve_fx(port_values):
	fx_table = {}
	cash_accounts = port_values['cash_accounts']
	for cash_account in cash_accounts:
		fx_table[cash_account['currency']] = cash_account['fx_rate']

	return fx_table



def write_csv(port_values):
	"""
	Write cash and holdings into csv files.
	"""	
	cash_file = get_current_path() + '\\cash.csv'
	write_cash_csv(cash_file, port_values)

	holding_file = get_current_path() + '\\bond_holding.csv'
	write_bond_holding_csv(holding_file, port_values)

	holding_file = get_current_path() + '\\equity_holding.csv'
	write_equity_holding_csv(holding_file, port_values)



def write_cash_csv(cash_file, port_values):
	with open(cash_file, 'w', newline='') as csvfile:
		file_writer = csv.writer(csvfile)

		cash_accounts = port_values['cash_accounts']

		fields = ['bank', 'date', 'account_type', 
					'account_num', 'currency', 'balance', 
					'fx_rate', 'hkd_equivalent']

		file_writer.writerow(fields)
		for cash_account in cash_accounts:
			row = []
			for fld in fields:
				item = cash_account[fld]
				if fld == 'date':
					item = convert_datetime_to_string(item)
				row.append(item)

			file_writer.writerow(row)



def write_bond_holding_csv(holding_file, port_values):
	with open(holding_file, 'w', newline='') as csvfile:
		file_writer = csv.writer(csvfile)

		fields = ['isin', 'name', 'currency', 'accounting_treatment', 
				'par_amount', 'is_listed', 'listed_location', 
                'fx_on_trade_day', 'coupon_rate', 'coupon_start_date', 
                'maturity_date', 'average_cost', 'amortized_cost', 
                'price', 'book_cost', 'interest_bought', 'amortized_value', 
                'market_value', 'accrued_interest', 'amortized_gain_loss', 
                'market_gain_loss', 'fx_gain_loss']

		file_writer.writerow(fields)
		bond_holding = port_values['bond']
		for bond in bond_holding:
			if bond['par_amount'] == 0:
				continue

			row = []
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



def write_equity_holding_csv(holding_file, port_values):
	with open(holding_file, 'w', newline='') as csvfile:
		file_writer = csv.writer(csvfile)

		fields = ['ticker', 'isin', 'name', 'currency', 'accounting_treatment', 
					'number_of_shares', 'currency', 'fx_on_trade_day', 
					'last_trade_date', 'average_cost', 'price', 'book_cost', 
                    'market_value', 'market_gain_loss', 'fx_gain_loss']

		file_writer.writerow(fields)
		equity_holding = port_values['equity']
		for equity in equity_holding:
			if equity['number_of_shares'] == 0:
				continue

			row = []
			for fld in fields:
				try:
					item = equity[fld]
					if fld == 'last_trade_date':
						item = convert_datetime_to_string(item)
				except KeyError:
					item = ''

				row.append(item)

			file_writer.writerow(row)



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