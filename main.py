from processors.outlook_processor import OutlookProcessor
from utils.logger import setup_logger
from utils.loader import loading
import logging

if __name__ == "__main__":
    setup_logger()

    logging.info("Application started")
    loading("Starting automation", 6)

    processor = OutlookProcessor()
    processor.run()

    loading("Finalizing", 4)
    logging.info("Application finished")