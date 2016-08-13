# coding=utf-8
# 
# This file is used to parse the diversified income fund excel files from
# trustee, read the necessary fields and save into a csv file for
# reconciliation with Advent Geneva.
#
# To use it, 
#
# try:
#	  port_values = {}
# 	  open_excel(file_name, port_values)	# file_name being the trustee excel
# except:
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



def open_excel(file_name, port_values):
	"""
	Open the excel file, populate portfolio values based on contents of that
	file. 

	port_values is a dictionary object, the place holder for DIF portfolio 
	information.
	"""
	try:
		wb = open_workbook(filename=file_name)
	except:
		logger.critical('open_excel(): DIF file {0} cannot be opened'
							.format(file_name))
		logger.exception('open_excel()')
		raise

	sn = 'Portfolio Sum.'	# read the portfolio summary sheet
	try:
		ws = wb.sheet_by_name(sn)
	except:
		logger.error('open_excel(): Sheet {0} cannot be opened'
						.format(sn))
		logger.exception('open_excel()')
		raise

	read_portfolio_summary(ws, port_values)

	# verify we have read the correct value
	# show_portfolio_summary(port_values)

	# read cash accounts from multiple sheets
	sheet_names = wb.sheet_names()
	count = 0
	for sn in sheet_names:
		if sn.endswith('-BOC'):	# search for cash sheets
			count = count + 1
			# print('read from sheet {0}'.format(sn))
			ws = wb.sheet_by_name(sn)
			read_cash(ws, port_values)

	if (count == 0):
		logger.error('open_excel(): Failed to find cash sheets')
		raise Exception('no cash sheet')	# indicate something wrong

	# verify we have read the correct value
	# show_cash_accounts(port_values)
