# coding=utf-8
# 
# Parse the Macau balanced fund and Macau guarantee fund excel files 
# from trustee, read the necessary fields and save into a csv file for
# reconciliation with Advent Geneva.
#

from xlrd import open_workbook
import re
from DIF.open_cash import read_cash
from DIF.open_summary import find_cell_string, read_date, read_cash_holding_total, \
							populate_value
from DIF.open_holding import read_holding
from DIF.open_dif import validate_cash_and_holding, write_csv
from DIF.utility import get_input_directory
import logging
logger = logging.getLogger(__name__)



class FundNameNotFound(Exception):
	pass

class FundIdNotFound(Exception):
	pass

class PositionCustodianNotFound(Exception):
	pass
	


def open_bal(file_name, port_values, output_dir, output_prefix):
	"""
	Open the excel file of the trustee Macau balanced fund or guarantee fund.
	Read its cash positions, holdings and output to csv files.
	"""
	wb = open_workbook(filename=file_name)

	ws = wb.sheet_by_name('Portfolio Sum.')
	read_portfolio_summary(ws, port_values)
	
	# find sheets that contain cash
	sheet_names = wb.sheet_names()
	for sn in sheet_names:
		if sn.startswith('CA') or sn.startswith('SA'):
			ws = wb.sheet_by_name(sn)
			read_cash(ws, port_values)

	consolidate_cash(port_values)
	
	ws = wb.sheet_by_name('Portfolio Val.')
	read_holding(ws, port_values)

	# make sure the holding and cash are read correctly
	validate_cash_and_holding(port_values)

	# output the cash and holdings into csv files.
	return write_csv(port_values, output_dir, output_prefix)



def read_portfolio_summary(ws, port_values):
	"""
	Similar to the read_portfolio_summary() function in open_summary.py,
	this function reads the portfolio summary of balanced fund and guarantee
	fund. The difference compared to DIF portfolio summary is:

	1. There is no net asset value in balanced/guarantee fund.
	2. The unit price value is in column B instead of column C.

	Also, here we set the portfolio id and position custodian for Balanced
	and Guarantee fund.
	"""
	logger.debug('in read_portfolio_summary()')

	row = find_cell_string(ws, 0, 0, 'Fund Name')
	port_values['portfolio_id'] = get_fund_id(ws.cell_value(row, 0))
	port_values['position_custodian'] = get_position_custodian(port_values['portfolio_id'])

	n = find_cell_string(ws, row, 0, 'Valuation Period :')
	row = row + n
	# print('row={0}'.format(row))
	port_values['date'] = read_date(ws, row, 3)

	# read the summary of cash and holdings
	n = read_cash_holding_total(ws, row, port_values)
	row = row + n

	n = find_cell_string(ws, row, 0, 'Due to, Due from')
	row = row + n
	populate_value(port_values, 'nav', ws, row, 10)

	n = find_cell_string(ws, row, 0, 'Total Units Held at this Valuation  Date')
	row = row + n 	# move to that row
	populate_value(port_values, 'number_of_units', ws, row, 2)

	# the second 'unit price' after 'net asset value' is the
	# the one we want to use.
	n = find_cell_string(ws, row, 0, 'Unit Price')
	row = row + n
	populate_value(port_values, 'unit_price', ws, row, 2)



def get_fund_id(name_string):
	"""
	Extract fund name from the string, it may look like:

	Fund Name (基金名稱): CHINA LIFE MACAU BRANCH BALANCED OPEN FUND 中國人壽澳門分公司開放式平衡基金
	
	We need to get the portion starting from "China LIFE"
	"""
	m = re.search('Fund Name.*:\s*(.*)', name_string)
	if m is not None:
		fund_name = m.group(1).strip()
	else:
		logger.error('get_fund_id(): failed to extract fund name from {0}'.format(name_string))
		raise FundNameNotFound()

	if fund_name.startswith('CHINA LIFE MACAU BRANCH BALANCED'):
		return '30004'
	elif fund_name.startswith('CHINA LIFE MACAU BRANCH GUARANTEE'):
		return '30003'
	else:
		logger.error('get_fund_id(): failed to map to fund id: {0}'.format(fund_name))
		raise FundIdNotFound()



def get_position_custodian(portfolio_id):
	"""
	Return the custodian bank for its security holdings based on the portfolio
	id.
	"""
	c_map = {
		'30003':'ICBCMACAU',
		'30004':'ICBCMACAU'
	}
	try:
		return c_map[portfolio_id]
	except KeyError:
		logger.error('get_position_custodian(): failed to locate custodian for portfolio {0}'.format(portfolio_id))
		raise PositionCustodianNotFound()



def consolidate_cash(port_values):
	"""
	For the balanced fund or guarantee fund, combine the checking
	and savings account for the same currency in the same bank.
	"""
	new_cash_accounts = []
	cash_accounts = port_values['cash_accounts']
	for cash_account in cash_accounts:
		if find_n_merge(cash_account, new_cash_accounts):
			continue

		new_cash_accounts.append(cash_account)

	port_values['cash_accounts'] = new_cash_accounts

	

def find_n_merge(cash_account, cash_accounts):
	"""
	For the input cash account, if another cash account with the same
	bank and currency is found in cash accounts, then merge it to
	the existing cash account, then return true. If not, do nothing,
	reture false.
	"""
	for ca in cash_accounts:
		if cash_account['bank'] == ca['bank'] and \
			cash_account['currency'] == ca['currency']:
			ca['balance'] = ca['balance'] + cash_account['balance']
			ca['local_currency_equivalent'] = ca['local_currency_equivalent'] + \
												cash_account['local_currency_equivalent']
			return True

	return False
