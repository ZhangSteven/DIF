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

1. Separate the holdings output to 2 files: HTM and Trading (trading bond and equity), so the we can match the price/market value local of the latter.

2. How about not reading the columns for each holding section, but read columns for each type of holdings once and re-use it later, e.g., HTM bond, trading bond, equity.

3. For HKD equivalent, email the other party to change?

4. Modify the error checking code, to include the sheet name it's reading?



++++++++++
ver 0.14
++++++++++
1. Now we output htm positions and afs positions into two files, instead of bond holding and equity holding. This way, we can recon the price/market value of the afs positions, as trustee also uses Bloomberg price, we can use this to check whether geneva has downloaded the prices properly.

waiting to be tested



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
