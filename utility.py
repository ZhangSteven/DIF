# coding=utf-8
# 
import configparser, os



class InvalidDatamode(Exception):
	pass


def get_current_path():
	"""
	Get the absolute path to the directory where this module is in.

	This piece of code comes from:

	http://stackoverflow.com/questions/3430372/how-to-get-full-path-of-current-files-directory-in-python
	"""
	return os.path.dirname(os.path.abspath(__file__))



def _load_config(filename='dif.config'):
	"""
	Read the config file, convert it to a config object. The config file is 
	supposed to be located in the same directory as the py files, and the
	default name is "config".

	Caution: uncaught exceptions will happen if the config files are missing
	or named incorrectly.
	"""
	path = get_current_path()
	config_file = path + '\\' + filename
	# print(config_file)
	cfg = configparser.ConfigParser()
	cfg.read(config_file)
	return cfg



# initialized only once when this module is first imported by others
if not 'config' in globals():
	config = _load_config()



def get_base_directory():
	"""
	The directory where the log file resides.
	"""
	global config
	directory = config['logging']['directory']
	if directory == '':
		directory = get_current_path()

	return directory



def get_datemode():
	"""
	Read datemode from the config object and return it (in integer)
	"""
	global config
	d = config['excel']['datemode']
	try:
		datemode = int(d)
	except:
		logger.error('get_datemode(): invalid datemode value: {0}'.format(d))
		raise InvalidDatamode()

	return datemode



def get_input_directory():
	"""
	Where the input files reside.
	"""
	global config
	directory = config['input']['directory']
	if directory.strip() == '':
		directory = get_current_path()

	return directory



def retrieve_or_create(port_values, key):
	"""
	retrieve or create the holding objects (list of dictionary) from the 
	port_values object, the holding place for all items in the portfolio.
	"""

	if key in port_values:	# key exists, retrieve
		holding = port_values[key]	
	else:					# key doesn't exist, create
		if key in ['bond', 'equity', 'cash_accounts', 'expense']:
			holding = []
		else:
			# not implemented yet
			logger.error('retrieve_or_create(): invalid key: {0}'.format(key))
			raise ValueError('invalid_key')

		port_values[key] = holding

	return holding



def get_holding_fx(port_values):
	if not 'holding_fx' in port_values:
		port_values['holding_fx'] = {}

	return port_values['holding_fx']