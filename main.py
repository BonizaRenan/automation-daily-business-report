from processors.outlook_processor import OutlookProcessor
from processors.report_email_processor import ReportEmailProcess

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
    loading("Generating reported", 4)

    savereport = ReportEmailProcess()
    savereport.run()

    loading("Finalizing report", done_text="Successfully Report...")
    logging.info("Application finished")

    