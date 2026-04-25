from services.outlook_service import outlookemails
from services.excel_service import export_to_excel
from datetime import datetime
import json
import os

now = datetime.now()
formatted_date = now.strftime("%m/%d/%Y")


def load_config(path="shared/Json/subject.json"):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def run_from_json():
    config = load_config()

    for item in config["Subject"]:
        subject = item["Email"]
        folder = item["Folder"]

        print(f"Processing: {subject}")

        if subject == "Test one":
            print("asdasdadsd")
            item["Status"] = "Merged"
        else:
            run_download_file(subject, folder, item)

    export_to_excel(config["Subject"])


def run_download_file(subject, folder, item):
    # Download email attachments from Outlook
    email_found = outlookemails(
        subject_filter=subject,
        download_folder=folder
    )

    if not email_found:
        item["Status"] = "No Email Received"
        print(f"No Email Found: {subject}")
        return

    # Check downloaded files
    if not os.path.exists(folder):
        item["Status"] = "No Folder Found"
        return

    files = os.listdir(folder)
    valid_files = [f for f in files if f.endswith((".xlsx", ".pdf", ".csv"))]

    if valid_files:
        item["Status"] = "Received"
    else:
        item["Status"] = "No File Received"