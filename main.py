from processors.outlook_processor import OutlookProcessor

from utils.logger import setup_logger
from utils.loader import loading
import logging


if __name__ == "__main__":

    setup_logger()

    logging.info("Application started")
    loading("Starting automation", 6)

    outlook = OutlookProcessor()
    outlook.run()

    loading("Outlook processing completed", done_text="Emails downloaded...")

    logging.info("Outlook processing finished")

    logging.info("Excel export started")
    loading("Generating report", 4)

  

    loading("Finalizing report", done_text="Successfully Exported...")

    logging.info("Application finished")