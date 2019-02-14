import logging


logger = logging.getLogger(__name__)

handler = logging.StreamHandler()

format = logging.Formatter('%(levelname)s - %(name)s - (line: %(lineno)s) - %(message)s')

handler.setFormatter(format)

logger.addHandler(handler)
logger.setLevel(logging.DEBUG)
