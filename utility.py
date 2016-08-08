# coding=utf-8
# 
# This file contains utility functions to be used by other modules in
# this package, such as logging and config.
# 
import logging
import configparser
import os



# Initialization
# if not 'configs' in globals():
#     configs = {}
#     parse_config('config.txt', configs)
#     log_lvl = configs['loglevel']
#     if log_lvl == 'debug':
#         log_lvl = logging.DEBUG
#     elif log_lvl == 'info':
#         log_lvl = logging.INFO
#     elif log_lvl == 'warn':
#         log_lvl = logging.WARN
#     else:
#         log_lvl = logging.DEBUG
#     logger = my_custom_logger('main', configs['logfile'], log_lvl)
#     logger.info('stockapp started')


def get_current_path():
	"""
	Get the absolute path to the directory where this module is in.

	This piece of code comes from:

	http://stackoverflow.com/questions/3430372/how-to-get-full-path-of-current-files-directory-in-python
	"""
	return os.path.dirname(os.path.abspath(__file__))



def _load_config(filename='config'):
	"""
	Read the global config file, convert it to a config object. The config file
	is supposed to be located in the same directory as the py files, and named
	as "dif.config".

	Caution: uncaught exceptions will happen if the config files are missing
	or named incorrectly.
	"""
	path = get_current_path()
	config_file = path + '\\' + filename
	print(config_file)
	cfg = configparser.ConfigParser()
	cfg.read(config_file)
	return cfg



# initialize the config object if it's not there
if not 'config' in globals():
	config = _load_config()



def convert_log_level(log_level):
	"""
	Convert the log level specified in the config file to the numerical
	values required by the logging module.
	"""
	if log_level == 'debug':
		return logging.DEBUG
	elif log_level == 'info':
		return logging.INFO
	elif log_level == 'warning':
		return logging.WARNING
	elif log_level == 'error':
		return logging.ERROR
	elif log_level == 'critical':
		return logging.CRITICAL
	else:
		return logging.DEBUG



# def _create_logger(name, filename, loglevel):
def _create_logger():
    """ 
    Creates a logger based on the python logging package. Supposed to be 
    called only once.

    Original code from:
    http://stackoverflow.com/questions/7621897/python-logging-module-globally
    """

    # use the config object
    global config

    filename = config['logging']['log_file']
    filename = get_current_path() + '\\' + filename

    formatter = logging.Formatter(fmt='%(asctime)s - %(levelname)s - %(module)s - %(message)s')
    handler = logging.FileHandler(filename)
    handler.setFormatter(formatter)

    logger_name = config['logging']['logger_name']
    log_level = config['logging']['log_level']
    log_level = convert_log_level(log_level)

    logger = logging.getLogger(logger_name)
    logger.setLevel(log_level)
    logger.addHandler(handler)
    
    return logger



# initialize the logger if it's not there
if not 'logger' in globals():
	logger = _create_logger()

