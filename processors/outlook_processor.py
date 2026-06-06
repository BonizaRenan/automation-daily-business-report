from services.outlook_service import OutlookService
from services.excel_merge_service import ExcelMergeService
from utils.config_loader import load_config
from datetime import datetime
import logging
import shutil
import os

from helpers.file_helper import (
    create_folder,
    create_daily_folder,
    get_valid_files,
    get_excel_files,
)


class OutlookProcessor:

    BASE_FOLDER = "C:/Users/Renan Boniza/Desktop/Automation/daily-business-report/shared/download"

    def __init__(self):
        self.outlook = OutlookService()
        self.excel = ExcelMergeService()
        self.config = load_config()

    # =========================
    # MAIN ENTRY
    # =========================
    def run(self):

        logging.info("Application started")

        folder_date = datetime.now().strftime("%m-%d-%Y")
        daily_root = create_daily_folder(self.BASE_FOLDER, folder_date)

        results = []

        for item in self.config.get("Subject", []):
            try:
                result = self.run_job(item, daily_root)
                if result:
                    results.append(result)

            except Exception as e:
                subject = item.get("email")
                logging.error(f"Error processing {subject}: {e}")

        logging.info("Application finished")

    # =========================
    # DISPATCHER
    # =========================
    def run_job(self, item, daily_root):

        subject = item.get("email")
        date_format = item.get("date_format")
        file_format = item.get("file_format")

        subject_with_date = f"{subject} - {datetime.now().strftime(date_format)}"

        logging.info(f"Processing: {subject_with_date}")

        handlers = {
            "single": self._handle_single,
            "merge": self._handle_merge,
            "generated": self._handle_generated
        }

        handler = handlers.get(file_format)

        if not handler:
            logging.error(f"{subject} - Unknown format: {file_format}")
            return

        return handler(item, subject_with_date, daily_root)

    # =========================
    # SINGLE
    # =========================
    def _handle_single(self, item, subject, daily_root):

        download_folder = daily_root

        email_found = self.outlook.download_emails(
            subject_filter=subject,
            download_folder=download_folder
        )

        files = get_valid_files(download_folder)

        if not email_found:
            logging.error(f"{subject} - No Email Received")
            return

        if not files:
            logging.error(f"{subject} - No File Received")
            return

    # =========================
    # MERGE
    # =========================
    def _handle_merge(self, item, subject, daily_root):

        folder = create_folder(daily_root, item)

        self.outlook.download_emails(
            subject_filter=subject,
            download_folder=folder
        )

        files = get_excel_files(folder)

        if not files:
            logging.error(f"{subject} - No files to merge")
            return

        output_prefix = item.get("file_name", "output")

        merged_file = os.path.join(folder, f"{output_prefix}_Merged.xlsx")

        self.excel.merge_files(files, merged_file, item)

        final_output = os.path.join(
            daily_root,
            f"{output_prefix}{datetime.now().strftime('%m%d')}.xlsx"
        )

        shutil.move(merged_file, final_output)

        # DELETE TEMP FOLDER
        try:
            if os.path.exists(folder):
                shutil.rmtree(folder)
                logging.info(f"Deleted temp folder: {folder}")
        except Exception as e:
            logging.info(f"Failed to delete folder {folder}: {e}")

    # =========================
    # GENERATED
    # =========================
    def _handle_generated(self, item, subject, daily_root):

        logging.info("Running selenium automation...")

        # TODO: selenium logic here