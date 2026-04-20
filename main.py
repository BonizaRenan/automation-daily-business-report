from services.outlook_service import outlookemails
from services.sharepoint_service import sharepoint
from services.excel_service import export_to_excel
import json
import os

def load_config(path="utils/Json/subject.json"):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def run_from_json():
    config = load_config()

    for item in config["Subject"]:
        subject = item["Email"]
        folder = item["Folder"]

        print(f"Processing: {subject}")

        # Outlook download
        email_found = outlookemails(
            subject_filter=subject,
            download_folder=folder
        )

        if not email_found:
            item["Status"] = "No Email Received"
            continue  # skip file checking

        # Check files if email exists
        files = os.listdir(folder) if os.path.exists(folder) else []

        valid_files = [f for f in files if f.endswith((".xlsx", ".pdf", ".csv"))]

        if valid_files:
            item["Status"] = "Received"
        else:
            item["Status"] = "No File Received"

    # 3. Save JSON
    with open("utils/Json/subject.json", "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4)

    # 4. Export Excel
    export_to_excel(config["Subject"])


def main():
     run_from_json()
    # sharepoint()


if __name__ == "__main__":
    main()