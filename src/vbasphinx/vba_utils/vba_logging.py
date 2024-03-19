'''defines logger'''

import sys
import logging
import logging.handlers

def setup_logger(fname):
    '''sets up logging to screen and file named fname'''
    hdlr = logging.FileHandler(fname, 'w', encoding='utf-8')
    screen_hdlr = logging.StreamHandler(sys.stdout)
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    logger.addHandler(hdlr)
    logger.addHandler(screen_hdlr)
