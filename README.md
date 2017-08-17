# DIF

Convert the trustee's DIF excel to Geneva's format, for reconciliation purpose.

+++++++++
Testing
+++++++++

We use nose2 to do all the unit testing. To run all tests, go the project directory and run "nose2". All test classes are in directory "test".

nose2 stops logging to the log file by default, instead it redirects the log messages to stdout. By default messages with logging level equal or above
logging.WARNING gets displayed. To display debug messages, do:

	nose2 --log-level DEBUG

To run test cases only in test_holding.py module, as it is under the test/ directory, do:

	nose2 -s test test_holding

To run a specific test method in test_holding.py, do:

	nose2 -s test test_holding.TestHolding.test_read_bond_fields_HTM

For more information, see:

see http://stackoverflow.com/questions/17890087/how-to-run-specific-test-in-nose2


+++++++++
To do
+++++++++

1. Modify the error checking code, to include the sheet name being read.


+++++++++++++++++++++
ver 0.24 @ 2017-8-17
+++++++++++++++++++++
1. Added a few lines to open_bal.read_portfolio_summary() function, to read NAV for balanced and guarantee funds.



+++++++++++++++++++++
ver 0.2301 @ 2017-8-16
+++++++++++++++++++++
1. Changed logging level from ERROR to WARNING when an expense item is not valid, in open_expense.py, because invalid expense items are simply ignored and not included in total expense calculations.



+++++++++++++++++++++
ver 0.23 @ 2017-8-16
+++++++++++++++++++++
1. Fixed a bug in logging. Previously it used utility.py to create a root logger, but now every module uses the standard way to obtain a logger:
	
	logger = logging.getLogger(__name__)

	when modules from other packages call open_dif.py, and if those modules configure their own root logger, there is a conflict. The new way of obtaining a logger solves this problem.



+++++++++++++++++++++
ver 0.2201 @ 2017-8-7
+++++++++++++++++++++
1. Fix a little bit in output csv filenames.



+++++++++++++++++++++
ver 0.22 @ 2017-8-7
+++++++++++++++++++++
1. Added support for trustee Macau Balanced fund and Guarantee fund.



+++++++++++++++++++++
ver 0.21 @ 2017-7-12
+++++++++++++++++++++
1. Changed open_dif() in open_dif.py, so that it passes the "Broker-MS" sheet as a cash sheet to the open_cash() function. This is because that sheet contains the futures cash balance.



++++++++++
ver 0.2
++++++++++
1. Disabled validate_expense_date() function in open_dif.py, because the expense date may or may not match the portfolio date. In DIF 2017-7-10 file, there is a bank charge dated 2016-11-1. Therefore we disable the function.

2. Test case updated to reflect that change.


++++++++++
ver 0.1901
++++++++++
1. Changed validate_cash_and_holding() function, so that it won't raise error unless the difference between the calculated equity total and the total in file is larger than 0.1 (previously 0.01). Because the DIF file on 2017-3-10 has a difference of 0.013.



++++++++++
ver 0.19
++++++++++
1. Changed write_csv() function to return the list of output csv files. This is required by the recon_helper package.



++++++++++
ver 0.18
++++++++++
1. Bug fix: Add output_dir to open_dif() function.



++++++++++
ver 0.17
++++++++++
1. Add an output_dir parameter to write_csv() function, so that it can work with the reconciliation_helper package. The output_dir parameter's default value is the input directory, so it stays backward compatible with ver 0.16, if working in standalone mode (python open_dif.py <input_file>), it still produces the same behaviour.



++++++++++
ver 0.16
++++++++++
1. Bug fix: previous equity holding validation uses amount/100 to calculate the total market value, however for DIF holdings on 2016-12-16, they mix fund holding with preferred shares holding, so this does not work anymore. Now we add up market value directly.



++++++++++
ver 0.16
++++++++++
1. Adjusted the tolerance value for bond total to 0.2 instead of 0.1. Because:

	1.1 For trustee file "CL Franklin DIF 2016-12-15.xls", the difference is 0.115.

	1.2 The are two decimal places for "bond value" and "bond amortization" in the NAV file, so rounding error should be at most 0.1+0.1 = 0.2, if trustee rounding is 0.119 to 0.11, just cut off the 3rd decimal point number.



++++++++++
ver 0.15
++++++++++
Tested with Geneva custom loaders (HTM position, AFS position, cash).

1. For the first HTM bond and Trading bond section, we save the data field names (usually they are good). For subsequent HTM bond/trading bond sections, we read the data field names and compare with the saved one, if they are different, leave a warning message in the log and reuse the saved one. For equity field names we don't do this because the equity section for real equity and for preferred shares treated as equity are different.

2. Change the read_bond_fields() and read_equity_fields() functions to make them easier to understand, change the if...elifs to a dictionary.



++++++++++
ver 0.14
++++++++++
Tested OK with Geneva custom loaders (HTM position, AFS position, cash)

1. Now we output htm positions and afs positions into two files, instead of bond holding and equity holding. This way, we can recon the price/market value of the afs positions, as trustee also uses Bloomberg price, we can use this to check whether geneva has downloaded the prices properly.

2. Change the delimiter to '|' instead of ',' for csv file, because when the security name contains a comma, although Python csvwrite wraps it with "" and Excel can read it properly, Geneva custom loader not able to parse it correctly.



++++++++++
ver 0.13
++++++++++
This version works with Geneva custom loaders (cash, position).

1. Change the output csv for bond positions, to handle both HTM and Trading positions.
2. Use investment_lookup.id_lookup's get_investment_ids() function to get the appriate investment IDs for HTM and Trading positions.
3. Modify equity and cash output csv to work with the custom data loaders of Geneva.



++++++++++
ver 0.12
++++++++++
1. Use config_logging's logging function instead of its own file logging.
2. Read input DIF file and output the csv files to the directory specified in the config file, instead of in the local directory.



+++++++++
ver 0.11
+++++++++

1. Add 3 columns "date", "portfolio" and "custodian" to the equity and bond holding files.

2. Convert the JP Morgan ticker "N0011", "H0939" to Bloomberg ticker format, 11 HK and 939 HK. The conversion assumes all stocks are HK stocks, ticker always start with "N" or "H".



+++++++++
ver 0.1
+++++++++

1. Can read cash, holdings (bond and equity), expenses from the trustee xls files.

2. Provides validation for cash and holding data, in open_dif.py.

3. Output cash, bond holding and equity holding to 3 csv files.

Usage:

	python open_dif.py <trustee_excel_file>

Note the trustee excel file must be put into the same directory as the open_dif.py



+++++++++
ver 0.01
+++++++++

Be able to read a sample xls file from trustee and read a few values, just to verify that the xlrd package works.
