# coding=utf-8
# 
# Read the holdings section of the excel file from trustee.
#
# 

from xlrd import open_workbook
from xlrd.xldate import xldate_as_datetime
import xlrd
import datetime
from DIF.utility import logger


def read_holding(ws, port_values, datemode=0):
	"""
	Read the worksheet with portfolio holdings. To retrieve holding, 
	we do:

	equity_holding = port_values['equity']
	for id in equity_holding:
		equity = equity_holding[id]
		... retrive equity values using the following key ...

		ticker, isin, name, number_of_shares, currency, listed_location, 
		fx_trade_date, last_trade_date, average_cost, price, book_cost,
		market_value

	bond_holding = port_values['bond']
	for id in bond_holding:
		bond = bond_holding[id]
		... retrive bond values using the following key ...

		isin, name, accounting_treatment, par_amount, currency, is_listed, 
		listed_location, fx_trade_date, coupon_rate, coupon_start_date, 
		maturity_date, average_cost, amortized_cost, price, book_cost,
		interest_bought, amortized_value, market_value, accrued_interest,
		amortized_gain_loss, market_gain_loss, fx_gain_loss

	Note a bond will not have all of the above fields, depending on
	its accounting treatment, HTM bonds have amortized_cost, amortized_value,
	amortized_gain_loss, and have price, market_value, market_gain_loss set to
	zero. Bonds for trading is the opposite.

	"""
	logger.debug('in read_holding()')

	for row in range(ws.nrows):
			
		# search the first column
		cell_value = ws.cell_value(row, 0)

		if isinstance(cell_value, str) and '.' in cell_value:
			tokens = cell_value.split('.')
			if len(tokens) > 1:
				if str.strip(tokens[1]).startswith('Debt Securities'):	# bond
					print('bond: {0}'.format(cell_value))
				elif str.strip(tokens[1]).startswith('Equities'):		# equity
					print('equity: {0}'.format(cell_value))

	logger.debug('out of read_holding()')


