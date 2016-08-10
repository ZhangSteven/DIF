# coding=utf-8
# 
# This file is used to parse the diversified income fund excel files from
# trustee, read the necessary fields and save into a csv file for
# reconciliation with Advent Geneva.
#
# To use it, 
#
# try:
# 	  port_values = open_excel(file_name)	# file_name being the trustee excel
# except Exception:
#	  ... error handling ...
#
# then we can query the different attributes of the portfolio, like
# the following:
#
# nav = port_values['nav']
#
# cash_accounts = port_values['cash_accounts']
# for id in cash_accounts:
#	....
# 

from xlrd import open_workbook
from DIF.open_cash import read_cash, show_cash_accounts
from DIF.open_summary import read_portfolio_summary, show_portfolio_summary
from DIF.utility import logger



def open_excel(file_name):
	"""
	Open the excel file, populate portfolio values into a dictionary. 

	Return the dictionary containing portfolio values if everything is
	OK, otherwise return None.
	"""
	try:
		wb = open_workbook(filename=file_name)
	except Exception as e:
		logger.critical('DIF file {0} cannot be opened'.format(file_name))
		logger.exception('open_excel()')
		return None

	port_values = {}	# the place holder for DIF portfolio information

	sn = 'Portfolio Sum.'	# read the portfolio summary sheet
	try:
		ws = wb.sheet_by_name(sn)
		read_portfolio_summary(ws, port_values)
	except Exception as e:
		logger.critical('Sheet {0} cannot be opened'.format(sn))
		logger.exception('open_excel()')
		return None

	# verify we have read the correct value
	show_portfolio_summary(port_values)

	# read cash accounts from multiple sheets
	sheet_names = wb.sheet_names()
	for sn in sheet_names:
		if sn.endswith('-BOC'):
			# print('read from sheet {0}'.format(sn))
			ws = wb.sheet_by_name(sn)

			try:
				read_cash(ws, port_values)
			except Exception as e:
				return None

	# verify we have read the correct value
	show_cash_accounts(port_values)