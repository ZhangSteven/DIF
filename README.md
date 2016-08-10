# DIF

Convert the trustee's DIF excel to Geneva's format, for reconciliation purpose.


Testing: go the project directory and run command "nose2". All test classes
are in directory "test".

Note: nose2 will by default stop logging to the log file, instead redirect the log messages to stdout. By default only logging level equal to or above
logging.ERROR will get displayed. To display debug messages, do:

	nose2 --log-level DEBUG

To run just a specific test_xxx.py module, suppose it is under the test/ directory, then do:

	nose2 -s test test_xxx

see http://stackoverflow.com/questions/17890087/how-to-run-specific-test-in-nose2


ver 0.01

Be able to read a sample xls file from trustee and read a few values, just to verify that the xlrd package works.
