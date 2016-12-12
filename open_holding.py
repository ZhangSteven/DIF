# coding=utf-8
# 
# Read the holdings section of the excel file from trustee.
#
# 

from xlrd import open_workbook
from xlrd.xldate import xldate_as_datetime
import xlrd
import datetime
from DIF.utility import logger, get_datemode, retrieve_or_create



class BadFieldName(Exception):
	pass

class BadAssetClass(Exception):
	pass



def read_holding(ws, port_values):
	"""
	Read the worksheet with portfolio holdings. To retrieve holding, 
	we do:

	equity_holding = port_values['equity']
	for equity in equity_holding:
		equity['ticker'], equity['name']
		... retrive equity values using the following key ...

		ticker, isin, accounting_treatment, name, number_of_shares, currency, 
		listed_location, fx_on_trade_day, last_trade_date, average_cost, price, 
		book_cost, market_value, market_gain_loss, fx_gain_loss

	bond_holding = port_values['bond']
	for bond in bond_holding:
		bond['isin'], bond['name']
		... retrive bond values using the following key ...

		isin, name, accounting_treatment, par_amount, currency, is_listed, 
		listed_location, fx_on_trade_day, coupon_rate, coupon_start_date, 
		maturity_date, average_cost, amortized_cost, price, book_cost,
		interest_bought, amortized_value, market_value, accrued_interest,
		amortized_gain_loss, market_gain_loss, fx_gain_loss

	Note a bond may not have all of the above fields, depending on
	its accounting treatment. A HTM bond has amortized_cost, amortized_value,
	amortized_gain_loss, while a trading bond has price, market_value, 
	market_gain_loss set to zero.

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

					currency = read_currency(cell_value)
					fields, n = read_bond_fields(ws, row)	# read the bond
					row = row + n							# field names

					n = read_section(ws, row, fields, 'bond', currency, port_values)
					row = row + n

				elif str.strip(tokens[1]).startswith('Equities'):		# equity
					logger.debug('equity: {0}'.format(cell_value))

					# equity_holding = retrieve_or_create(port_values, 'equity')
					currency = read_currency(cell_value)
					fields, n = read_equity_fields(ws, row)
					row = row + n

					n = read_section(ws, row, fields, 'equity', currency, port_values)
					row = row + n
		
		# move to next row
		row = row + 1

	logger.debug('out of read_holding()')



def read_currency(cell_value):
	"""
	Read the currency from the cell containing a section start, such as
	'V. Debt Securities - US$  (債務票據- 美元)',
	'V. Debt Securities - SGD  (債務票據- 星加坡元)',
	'X. Equities - USD (股票-美元)'

	From the above strings, the function return USD, SGD, USD
	"""
	temp_list = cell_value.split('-')
	token = temp_list[1]
	temp_list = token.split('(')
	currency = str.strip(temp_list[0])

	if currency == 'US$':
		currency = 'USD'	# make the correction

	return currency



def read_field_name(ws, row, column):
	"""
	Read a field name give its position.
	"""
	fld1 = ws.cell_value(row-1, column)
	fld2 = ws.cell_value(row, column)
	if isinstance(fld1, str) and isinstance(fld2, str):
		field = (str.strip(fld1), str.strip(fld2))
	else:
		logger.error('read_field_name(): invalid type in row {0}, {1} column {2}'.
						format(row-1, row, column))
		raise TypeError('bad field name type {0}, {1}'.format(fld1, fld2))

	# logger.debug(field)
	return field



def read_fields(ws, name_map, row):
	"""
	Read the field names for this bond section, it may be fields for held
	to maturity bond, or for trading bonds.
	"""
	rows_read = 1
	fields = []

	while (row+rows_read < ws.nrows):
			
		# search the first column
		cell_value = ws.cell_value(row+rows_read, 0)

		if cell_value == 'Description':
			for i in range(2, 17):
				field_tuple = read_field_name(ws, row+rows_read, i)
				try:
					fields.append(name_map[field_tuple])
				except:
					logger.exception('read_fields(): ')
					logger.error('read_fields(): bad field name at row {0}, column {1}, value = {2} {3}'.
									format(row+rows_read, i, ws.cell_value(row+rows_read-1, i),
											ws.cell_value(row+rows_read, i)))
					raise BadFieldName()

			break	# finished reading the fields

		# move to next row
		rows_read = rows_read + 1
	# end of while loop

	return fields, rows_read



def read_bond_fields(ws, row):
	"""
	Read the field names for this bond section, it may be fields for held
	to maturity bond, or for trading bonds.
	"""
	# rows_read = 1
	# fields = []

	name_map = {
		('', ''):'empty_field',

		# Bond fields section
		('票面值', 'Par Amt'):'par_amount',
		('上市 (是/否)', 'Listed (Y/N)'):'is_listed',
		('Primary', 'Exchange'):'listed_location',
		('(AVG) FX', 'for TXN'):'fx_on_trade_day',
		('Int.', 'Rate (%)'):'coupon_rate',
		('Int.', 'Start Day'):'coupon_start_date',
		('到期日', 'Maturity'):'maturity_date',
		('Cost', '(%)'):'average_cost',
		('Price', '(%)'):'price',
		('(Amortized)', '(%)'):'amortized_cost',
		('成本價', 'Book Cost'):'book_cost',
		('Int.', 'Bought'):'interest_bought',
		('市價', 'M. Value'):'market_value',
		('Adjusted Value', '(Amortized)'):'amortized_value',
		('應收利息', 'Accr. Int.'):'accrued_interest',
		('Year-End', 'Amortization'):'amortized_gain_loss',
		('Gain/(Loss)', 'M. Value'):'market_gain_loss',
		('FX', 'HKD Equiv.'):'fx_gain_loss'
	}

	return read_fields(ws, name_map, row)



def read_equity_fields(ws, row):
	"""
	Note the equity section contains listed equity and preferred shares. 
	For listed equity we'll see listed_location, but for preferred shares, 
	this field is missing.	
	"""

	name_map = {
		('', ''):'empty_field',

		('股數', 'Share'):'number_of_shares',
		('幣值', 'CCY'):'currency',
		('Location', 'of Listed'):'listed_location',
		('(AVG) FX', 'for TXN'):'fx_on_trade_day',
		('最後交易日', 'Latest V.D.'):'last_trade_date',
		('Avg.', 'Price'):'average_cost',
		('Market', 'Price'):'price',
		('成本價', 'Book Cost'):'book_cost',
		('市價', 'M. Value'):'market_value',
		('Gain/(Loss)', 'M. Value'):'market_gain_loss',
		('FX', 'HKD Equiv.'):'fx_gain_loss'
	}

	return read_fields(ws, name_map, row)



def adjust_fields(port_values, fields, asset_class, accounting_treatment):
	"""
	Sometimes there are missing fields in a section, say SGD bond HTM
	section tends to have empty columns for its data fields. To overcome
	this problem, we reuse the fields for the same type of holding from
	previous sections.
	"""
	# don't touch equity fields, because the fields for real equity and for
	# those preferred shares treated as equity are different.
	if asset_class == 'equity':
		return fields

	if not (asset_class, accounting_treatment) in port_values:
		port_values[(asset_class, accounting_treatment)] = fields
		return fields
	else:
		existing_fields = port_values[(asset_class, accounting_treatment)]
		if len(fields) != len(existing_fields):
			logger.warning('adjust_fields(): existing fields do not match with the fields passed in, reuse existing fields.')
			show_fields(existing_fields, fields)

		else:
			for i in range(len(fields)):
				if fields[i] != existing_fields[i]:
					logger.warning('adjust_fields(): existing fields do not match with the fields passed in, reuse existing fields.')
					show_fields(existing_fields, fields)

		return existing_fields



def show_fields(existing_fields, fields):
	for i in range(max(len(existing_fields), len(fields))):
		logger.info('{0} : {1}, {2}'.format(i, get_value(existing_fields, i), get_value(fields, i)))



def get_value(fields, index):
	try:
		return fields[index]
	except IndexError:
		return '<out of range>'



def read_section(ws, row, fields, asset_class, currency, port_values):
	"""
	Read a section in the worksheet (ws), starting on row number (row).
	fields being the list of fields to read from column C. For example,
	for HTM bond section, we expect to fields in the following order:

		par_amount, currency, is_listed, listed_location, fx_on_trade_day, 
		coupon_rate, coupon_start_date, maturity_date, average_cost, 
		amortized_cost, book_cost, interest_bought, amortized_value, 
		accrued_interest, amortized_gain_loss, fx_gain_loss

	for trading bonds, we expect to see fields in the following order:

		par_amount, currency, is_listed, listed_location, fx_on_trade_day, 
		coupon_rate, coupon_start_date, maturity_date, average_cost, 
		price, book_cost, interest_bought, market_value, accrued_interest,
		market_gain_loss, fx_gain_loss

	for listed equity, we expect to see fields in the following order:

		number_of_shares, currency, listed_location, fx_on_trade_day, 
		empty_field, last_trade_date, empty_field, average_cost, price, 
		book_cost, empty_field, market_value, empty_field, market_gain_loss, 
		fx_gain_loss

	for equity (preferred shares), we expect to see fields in the following order:
		number_of_shares, currency, empty_field, fx_on_trade_day, 
		empty_field, last_trade_date, empty_field, average_cost, price, 
		book_cost, empty_field, market_value, empty_field, market_gain_loss, 
		fx_gain_loss

	Return the number of rows read in this function
	"""

	# currently only handle these two types of asset class
	if not asset_class in ['equity', 'bond']:
		logger.error('read_section(): invalid asset class: {0}'.format(asset_class))
		raise BadAssetClass()

	rows_read = 1
	holding = retrieve_or_create(port_values, asset_class)

	while (row+rows_read < ws.nrows):
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
					accounting_treatment = 'HTM'
				
				elif temp_str.startswith('Trading'):
					accounting_treatment = 'Trading'

				else:
					# some other category other than HTM or Trading,
					# Needs to implement
					logger.error('read_section(): unhandled accounting treament at row {0} column 0, value = {1}'.
									format(row+rows_read, cell_value))
					raise ValueError('bad accounting treatment')

				fields = adjust_fields(port_values, fields, asset_class, accounting_treatment)
				n = read_sub_section(ws, row+rows_read, accounting_treatment, 
										fields, asset_class, currency, holding)
				rows_read = rows_read + n

		elif isinstance(cell_value, str) and cell_value.startswith('Total'):
			# the section ends
			break

		rows_read = rows_read + 1	# move to next row

	return rows_read



def read_sub_section(ws, row, accounting_treatment, fields, asset_class, currency, holding):
	"""
	Read a sub section in the worksheet (ws), starting on row number (row).

	Return the number of rows read in this function
	"""
	rows_read = 1

	while (row+rows_read < ws.nrows):
		cell_value = ws.cell_value(row+rows_read, 0)
		
		# logger.debug(cell_value)
		if isinstance(cell_value, str) and cell_value.startswith('('):

			# detect the start of a security holding position
			# a holding position looks like "(xxx) security name"
			i = cell_value.find(')', 1, len(cell_value)-1)
			if i > 0:	# the string looks like '(xxx) yyy'
				security = {}

				# start populating fields of a security, then save it to the 
				# security_holding list.
				token = cell_value[1:i]
				if (asset_class == 'bond'):
					security['isin'] = token
				elif (asset_class == 'equity'):
					if ('listed_location') in fields:	# it's listed equity
						security['ticker'] = token
					else:								# it's preferred shares
						security['isin'] = token

				security['name'] = cell_value
				security['currency'] = currency
				security['accounting_treatment'] = accounting_treatment

				column = 2	# now start reading fields in column C
				for field in fields:
					cell_value = ws.cell_value(row+rows_read, column)
					# logger.debug('{0},{1},{2}'.format(row+rows_read, column, cell_value))

					if field == 'empty_field':	# ignore this field, move to
												# next column
						column = column + 1
						continue

					field_value = validate_and_convert_field_value(field, cell_value)

					# if already has currency assigned (in the case of listed 
					# equity), check whether the currency value is inconsistent
					if field == 'currency' and 'currency'in security \
						and field_value != security['currency']:
								
						logger.error('read_sub_section(): inconsistent currency value at row {0}, column {1}'.
										format(row+rows_read, column))
						raise ValueError('inconsistent currency value')

					security[field] = field_value

					# if holding amount is zero, stop reading other fields.
					if (field == 'par_amount' or field == 'number_of_shares') \
						and field_value == 0:
						break

					column = column + 1
					# end of for loop
					
				holding.append(security)
				# logger.debug(isin)

		elif is_end_of_sub_section(ws, row+rows_read):
			# the subsection ends
			break

		rows_read = rows_read + 1	# move to next row
		# end of while loop

	return rows_read



def is_end_of_sub_section(ws, row):
	"""
	Tell whether this is the end of the sub section.
	
	If the first 4 columns are all empty, then it is a blank line, then
	it is the end of the sub section.
	"""
	for column in range(4):
		if not is_empty_cell(ws, row, column):
			return False

	return True



def is_empty_cell(ws, row, column):
	"""
	If the cell value is all white space or an empty string, then it is
	an empty cell.
	"""
	cell_value = ws.cell_value(row, column)
	if isinstance(cell_value, str) and str.strip(cell_value) == '':
		return True
	else:
		return False

	

def validate_and_convert_field_value(field, cell_value):
	"""
	Validate the field value read is of proper type, if yes then convert it
	to a proper value if necessary (e.g., an empty field to zero). If no,
	then raise an exception.
	"""

	field_value = cell_value

	# if the following field value type is not string, then
	# something must be wrong
	if field in ['is_listed', 'listed_location', 'currency'] \
		and not isinstance(cell_value, str):
		logger.error('validate_field_value(): field {0} should be a string: {1}'.
						format(field, cell_value))
		raise ValueError('bad field type: not a string')

	elif field in ['fx_on_trade_day', 'coupon_rate', 'average_cost', \
					'amortized_cost', 'price', 'book_cost', 'interest_bought', \
					'amortized_value', 'market_value', 'accrued_interest', \
					'amortized_gain_loss', 'market_gain_loss', 'fx_gain_loss', \
					'coupon_start_date', 'maturity_date', 'last_trade_date'] \
		and not isinstance(cell_value, float):
		logger.error('validate_field_value(): field {0} should be a float: {1}'.
						format(field, cell_value))
		raise ValueError('bad field type: not a float')

	elif field in ['par_amount', 'number_of_shares']:
		if isinstance(cell_value, float):
			# OK, no change
			pass
		elif isinstance(cell_value, str) and str.strip(cell_value) == '':
			# treat an empty holding as zero
			field_value = 0
		else:
			logger.error('validate_field_value(): field {0} should be a float or empty string: {1}'.
						format(field, cell_value))
			raise ValueError('bad field type: not a float or empty string')

	# convert float to python datetime object when necessary
	if field in ['coupon_start_date', 'maturity_date', 'last_trade_date']:
		datemode = get_datemode()
		field_value = xldate_as_datetime(cell_value, datemode)

	return field_value
