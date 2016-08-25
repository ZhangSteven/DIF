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
