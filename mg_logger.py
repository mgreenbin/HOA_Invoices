import logging

formatstr = "%(asctime)s %(levelname)-8s [%(funcName)s:%(lineno)d] %(message)s"
logging.basicConfig(
    format=formatstr,
    level=logging.DEBUG,
    filename="HOA_Invoices.log")

logger = logging.getLogger("App.log")
logger.warning("This is a warning: Get ready, it's about to BLOW")
logger.debug("Debug: find those buggers!")
logger.critical("It has reached CRITICAL MASS: It's over now Bud!")
print("End Logging.....")
