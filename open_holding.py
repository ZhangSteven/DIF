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

	"""
	Now trying to read the holdings worksheet. The structure of data is as
	follows:

	Section (bond/equity/forward/xxx):
		sub section:
			holding1
			holding2
			...

		sub section:
			...

	Section
		sub section:

	"""

	row = 0
	while (row < ws.nrows):
			
		# search the first column
		cell_value = ws.cell_value(row, 0)

		if isinstance(cell_value, str) and '.' in cell_value:
			tokens = cell_value.split('.')
			if len(tokens) > 1:
				if str.strip(tokens[1]).startswith('Debt Securities'):	# bond
					logger.debug('bond: {0}'.format(cell_value))

					n = read_bond_section(ws, row)	# read n rows
					row = row + n

				elif str.strip(tokens[1]).startswith('Equities'):		# equity
					logger.debug('equity: {0}'.format(cell_value))

					n = read_equity_section(ws, row)
					row = row + n
		
		# move to next row
		row = row + 1

	logger.debug('out of read_holding()')




def read_bond_section(ws, row):
	"""
	Read a bond section in the worksheet (ws), starting on row number (row).

	Return the number of rows read in this function
	"""
	rows_read = 1

	while True:
		cell_value = ws.cell_value(row+rows_read, 0)
		
		# logger.debug(cell_value)
		if isinstance(cell_value, str) and cell_value.startswith('('):

			# detect the start of a subsection
			# a subsection looks like "(i) Held to Maturity (xxx)"
			i = cell_value.find(')', 1, len(cell_value)-1)
			if i > 0:	# the string looks like '(xxx) yyy'
				temp_str = str.strip(cell_value[i+1:])
				
				# logger.debug(temp_str)
				if temp_str.startswith('Held to Maturity'):	# found HTM sub sec
					logger.debug('HTM: {0}'.format(cell_value))
					n = read_bond_sub_section(ws, row+rows_read, 'HTM')
					rows_read = rows_read + n
				
				elif temp_str.startswith('Trading'):
					logger.debug('Trading: {0}'.format(cell_value))
					n = read_bond_sub_section(ws, row+rows_read, 'Trading')
					rows_read = rows_read + n

				else:
					# some other category other than HTM or Trading,
					# maybe needs to implement in the future
					pass

		elif isinstance(cell_value, str) and cell_value.startswith('Total'):
			# the section ends
			break

		rows_read = rows_read + 1	# move to next row

	return rows_read



def read_bond_sub_section(ws, row, category):
	"""
	Read a bond section in the worksheet (ws), starting on row number (row).

	Return the number of rows read in this function
	"""
	rows_read = 1

	while True:
		cell_value = ws.cell_value(row+rows_read, 0)
		
		# logger.debug(cell_value)
		if isinstance(cell_value, str) and cell_value.startswith('('):

			# detect the start of a bond holding position
			# a holding position looks like "(isin code) security name"
			i = cell_value.find(')', 1, len(cell_value)-1)
			if i > 0:	# the string looks like '(xxx) yyy'
				isin = cell_value[1:i]
				logger.debug(isin)

		elif isinstance(cell_value, str) and str.strip(cell_value) == '':
			# the subsection ends
			break

		rows_read = rows_read + 1	# move to next row

	return rows_read



def read_equity_section(ws, row):

	return 0