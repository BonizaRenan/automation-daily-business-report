from services.outlook_service import OutlookService
from services.excel_service import ExcelService
from utils.config_loader import load_config
from datetime import datetime
import logging
import os
import shutil


class OutlookProcessor:

    def __init__(self):
        self.outlook = OutlookService()
        self.excel = ExcelService()
        self.config = load_config()

    # =========================
    # CARTON CLEAN UP HANDLER
    # =========================
    def handle_carton_cleanup(self, subject, folder):
        folder_path = os.path.join(folder, "Carton_Clean_Up")
        os.makedirs(folder_path, exist_ok=True)

        self.outlook.download_emails(
            subject_filter=subject,
            download_folder=folder_path
        )

        files = os.listdir(folder_path)

        carton_files = [
            os.path.join(folder_path, f)
            for f in files
            if f.endswith(".xlsx") and "merged" not in f.lower()
        ]

        if not carton_files:
            return {
                "Subject": subject,
                "Status": "No Files Found"
            }

        temp_output = os.path.join(folder_path, "Carton_Clean_Up_Merged.xlsx")

        self.excel.merge_carton_files(carton_files, temp_output)

        parent_folder = os.path.dirname(folder_path)

        final_output = os.path.join(
            parent_folder,
            f"Carton_Clean_Up_{datetime.now().strftime('%m%d%Y')}.xlsx"
        )

        shutil.move(temp_output, final_output)

        return {
            "Subject": subject,
            "Status": "Received",
            "File": final_output
        }

    # =========================
    # MAIN PROCESS
    # =========================
    def run(self):

        formatted_date = datetime.now().strftime("%m/%d/%Y")
        results = []

        for item in self.config["Subject"]:
            subject = item["Email"]
            folder = item["Folder"]

            subject_with_date = f"{subject} {formatted_date}"
            logging.info(f"Processing: {subject_with_date}")

            # =========================
            # SPECIAL CASE
            # =========================
            if subject == "Carton Clean Up":
                result = self.handle_carton_cleanup(subject_with_date, folder)
                results.append(result)
                continue

            # =========================
            # DEFAULT FLOW
            # =========================
            email_found = self.outlook.download_emails(
                subject_filter=subject_with_date,
                download_folder=folder
            )

            if not email_found:
                results.append({
                    "Subject": subject,
                    "Status": "No Email Received"
                })
                continue

            if not os.path.exists(folder):
                results.append({
                    "Subject": subject,
                    "Status": "No Folder Found"
                })
                continue

            files = os.listdir(folder)
            valid_files = [
                os.path.join(folder, f)
                for f in files
                if f.endswith((".xlsx", ".pdf", ".csv"))
            ]

            results.append({
                "Subject": subject,
                "Status": "Received" if valid_files else "No File Received",
                "FilePath": folder if valid_files else []
                })

    # =========================
    # EXPORT FINAL REPORT
    # =========================
        self.excel.export(results)