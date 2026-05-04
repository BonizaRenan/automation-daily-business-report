from services.outlook_service import OutlookService
from services.excel_service import ExcelService
from utils.config_loader import load_config
from utils.loader import loading
from datetime import datetime
import logging
import shutil
import os


class OutlookProcessor:

    def __init__(data):
        data.outlook = OutlookService()
        data.excel = ExcelService()
        data.config = load_config()

    def run(data):

        loading("Running automation", done_text="Successfully Automated...")

        formatted_date = datetime.now().strftime("%m/%d/%Y")
        folder_date = datetime.now().strftime("%m-%d-%Y")

        results = []

        base_folder = "C:/Users/Renan Boniza/Desktop/Automation/daily-business-report/shared/download"


        daily_root = os.path.join(base_folder, f"Daily Business Report {folder_date}")
        os.makedirs(daily_root, exist_ok=True)

        for item in data.config["Subject"]:
            subject = item["Email"]
            subject_with_date = f"{subject} {formatted_date}"

            logging.info(f"Processing: {subject_with_date}")

            # =========================
            # CARTON CLEAN UP
            # =========================
            if subject == "Carton Clean Up":
                result = data.handle_carton_cleanup(subject_with_date, daily_root)
                results.append(result)
                continue

            # =========================
            # NORMAL DOWNLOAD
            # =========================
            email_found = data.outlook.download_emails(
                subject_filter=subject_with_date,
                download_folder=daily_root
            )

            if not email_found:
                results.append({
                    "Subject": subject_with_date,
                    "Status": "No Email Received",
                    "File": ""
                })
                continue

            valid_files = [
                f for f in os.listdir(daily_root)
                if f.endswith((".xlsx", ".pdf", ".csv", ".xlsm"))
            ]

            results.append({
                "Subject": subject_with_date,
                "Status": "Received" if valid_files else "No File Received",
                "File": daily_root if valid_files else ""
            })

     
        report_file = os.path.join(daily_root, "Export_Report.xlsm")

        data.excel.export(results, output_file=report_file)

        logging.info(f"FINAL REPORT CREATED: {report_file}")


    def handle_carton_cleanup(self, subject, daily_root):

        folder_path = os.path.join(daily_root, "Carton_Clean_Up")
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
                "Status": "No Files Found",
                "File": ""
            }

        temp_output = os.path.join(folder_path, "Carton_Clean_Up_Merged.xlsx")

        self.excel.merge_carton_files(carton_files, temp_output)

        final_output = os.path.join(
            daily_root,
            f"Carton_Clean_Up_{datetime.now().strftime('%m%d%Y')}.xlsx"
        )

        shutil.move(temp_output, final_output)

        return {
            "Subject": subject,
            "Status": "Received",
            "File": daily_root
        }