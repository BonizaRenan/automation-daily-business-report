from services.outlook_service import OutlookService
from services.excel_service import ExcelService
from utils.config_loader import load_config
from utils.loader import loading
from datetime import datetime
import logging
import shutil
import os


from helpers.file_helper import (
    create_folder,
    create_daily_folder,
    get_valid_files,
    get_excel_files,
    build_result,
    generate_report
)


class OutlookProcessor:

    BASE_FOLDER = "C:/Users/Renan Boniza/Desktop/Automation/daily-business-report/shared/download"

    def __init__(self):
        self.outlook = OutlookService()
        self.excel = ExcelService()
        self.config = load_config()

    # =========================
    # MAIN ENTRY
    # =========================
    def run(self):
        loading("Running automation", done_text="Successfully Automated...")

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
                results.append(build_result(subject, f"Error: {str(e)}"))

        generate_report(self.excel, results, daily_root)

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
            return build_result(subject, f"Unknown format: {file_format}")

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
            return build_result(subject, "No Email Received")

        if not files:
            return build_result(subject, "No File Received")

        return build_result(subject, "Received", download_folder)

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
            return build_result(subject, "No Files Found")

        output_prefix = item.get("file_name", "output")

        merged_file = os.path.join(folder, f"{output_prefix}_Merged.xlsx")

        self.excel.merge_files(files, merged_file, item)

        final_output = os.path.join(
            daily_root,
            f"{output_prefix}_{datetime.now().strftime('%m%d')}.xlsx"
        )

        shutil.move(merged_file, final_output)

        return build_result(subject, "Received", final_output)

    # =========================
    # GENERATED
    # =========================
    def _handle_generated(self, item, subject, daily_root):

        logging.info("Running selenium automation...")

        # TODO: selenium logic here

        return build_result(subject or "Generated Report", "Generated", daily_root)