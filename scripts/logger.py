"""Settup for a logger
"""
import logging

LOG_NAME = 'LOG record.log'

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

formatter = logging.Formatter('%(levelname)s:(%(asctime)s) In module: %(module)s at Line %(lineno)d --> %(message)s')
file_handler = logging.FileHandler(LOG_NAME)
file_handler.setFormatter(formatter)

logger.addHandler(file_handler)


if __name__ == '__main__':
    logger.error('Teste')