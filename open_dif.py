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
from DIF.utility import logger, config



class InconsistentValue(Exception):
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

		
	except:
		logger.exception('open_dif()')
		raise



def validate_cash_and_holding(port_values):
	"""
	Calculate subtotal of cash, bond holdings and equity holdings, then 
	compare to the value from the excel file.

	Based on experience, the difference between the subtotal value and the
	calculated subtotal is below 0.01 but above 0.001. Maybe this is due to
	the rounding of actual number before they are input to excel.
	"""
	cash_total = calculate_cash_total(port_values)
	if abs(cash_total - port_values['cash_total']) > 0.01:
		logger.error('validate_cash_holding(): calculated cash total {0} is inconsistent with that from file {1}'.
						format(cash_total, port_values['cash_total']))
		raise InconsistentValue

	fx_table = retrieve_fx(port_values)
	
	bond_holding = port_values['bond']
	bond_subtotal = calculate_bond_total(bond_holding, fx_table)
	if abs(bond_subtotal - port_values['bond_total']) > 0.01:
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



