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

    BASE_FOLDER = os.path.abspath("shared/download")

    def __init__(self):

        self.config = load_config()

        self.outlook = OutlookService()
        self.excel = ExcelMergeService()

    # =========================
    # MAIN ENTRY
    # =========================
    def run(self):

        logging.info("Application started")

        folder_date = datetime.now().strftime("%m-%d-%Y")
        daily_root = create_daily_folder(self.BASE_FOLDER, folder_date)

        results = []

        for item in self.config.get("Subject", []):

            if not item:
                continue

            try:
                result = self.run_job(item, daily_root)
                if result:
                    results.append(result)

            except Exception as e:
                logging.error(f"Error processing {item.get('email')}: {e}")

        logging.info("Application finished")
        return results

    # =========================
    # DISPATCHER
    # =========================
    def run_job(self, item, daily_root):

        subject = item.get("email")
        date_format = item.get("date_format", "%m-%d-%Y")
        file_format = item.get("file_format")

        if not subject:
            logging.error("Missing subject/email in config item")
            return False

        subject_with_date = f"{subject} - {datetime.now().strftime(date_format)}"

        logging.info(f"Processing: {subject_with_date} ({file_format})")

        handlers = {
            "single": self._handle_single,
            "merge": self._handle_merge,
            "generated": self._handle_generated
        }

        handler = handlers.get(file_format)

        if not handler:
            logging.error(f"{subject} - Unknown format: {file_format}")
            return False

        return handler(item, subject_with_date, daily_root)

    # =========================
    # SINGLE
    # =========================
    def _handle_single(self, item, subject, daily_root):

        download_folder = daily_root

        # ✅ folder already guaranteed by helper, but safe check
        os.makedirs(download_folder, exist_ok=True)

        email_found = self.outlook.download_emails(
            subject_filter=subject,
            download_folder=download_folder
        )

        files = get_valid_files(download_folder)

        if not email_found:
            logging.error(f"{subject} - No email received")
            return False

        if not files:
            logging.error(f"{subject} - No valid files found")
            return False

        logging.info(f"{subject} - Single completed")
        return True

    # =========================
    # MERGE
    # =========================
    def _handle_merge(self, item, subject, daily_root):

        folder = create_folder(daily_root, item)

        email_found = self.outlook.download_emails(subject, folder)

        files = get_excel_files(folder)

        if not email_found:
            logging.error(f"{subject} - No email received")
            return False

        if not files:
            logging.error(f"{subject} - No files to merge")
            return False

        output_prefix = item.get("file_name", "output")

        merged_file = os.path.join(folder, f"{output_prefix}_Merged.xlsx")

        self.excel.merge_files(files, merged_file, item)

        final_output = os.path.join(
            daily_root,
            f"{output_prefix}{datetime.now().strftime('%m%d')}.xlsx"
        )

        shutil.move(merged_file, final_output)

        # cleanup safely
        try:
            if os.path.exists(folder):
                shutil.rmtree(folder)
        except Exception as e:
            logging.warning(f"Failed cleanup folder {folder}: {e}")

        logging.info(f"{subject} - Merge completed")
        return True

    # =========================
    # GENERATED
    # =========================
    def _handle_generated(self, item, subject, daily_root):

        logging.info(f"{subject} - Running selenium automation...")
        return True