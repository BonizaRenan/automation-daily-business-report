from services.outlook_service import OutlookService
from services.excel_service import ExcelService
from utils.config_loader import load_config
import logging
import os


class OutlookProcessor:
    def __init__(self):
        self.outlook = OutlookService()
        self.excel = ExcelService()
        self.config = load_config()

    def run(self):
        for item in self.config["Subject"]:
            subject = item["Email"]
            folder = item["Folder"]

            logging.info(f"Processing: {subject}")

            if subject == "Test one":
                logging.info("Special case triggered")
                item["Status"] = "Merged"
            else:
                self.run_download_file(subject, folder, item)

        # Export results after processing
        self.excel.export(self.config["Subject"])

    def run_download_file(self, subject, folder, item):
        email_found = self.outlook.download_emails(
            subject_filter=subject,
            download_folder=folder
        )

        if not email_found:
            item["Status"] = "No Email Received"
            # logging.warning(f"No Email Found: {subject}")
            return

        if not os.path.exists(folder):
            item["Status"] = "No Folder Found"
            # logging.warning(f"Folder not found: {folder}")
            return

        files = os.listdir(folder)
        valid_files = [f for f in files if f.endswith((".xlsx", ".pdf", ".csv"))]

        if valid_files:
            item["Status"] = "Received"
        else:
            item["Status"] = "No File Received"