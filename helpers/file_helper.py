import os
import logging

def create_folder(daily_root, item):
    folder_name = item.get("file_name")

    if not folder_name:
        raise ValueError(f"Missing file_name in config: {item}")

    path = os.path.join(daily_root, folder_name)
    os.makedirs(path, exist_ok=True)
    return path


def create_daily_folder(base_folder, folder_date):
    path = os.path.join(base_folder, f"Daily Business Report {folder_date}")
    os.makedirs(path, exist_ok=True)
    return path


def get_valid_files(folder):
    return [
        f for f in os.listdir(folder)
        if f.endswith((".xlsx", ".pdf", ".csv", ".xlsm", ".xls"))
    ]


def get_excel_files(folder):
    return [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.endswith((".xlsx", ".xlsm", ".xls")) and "merged" not in f.lower()
    ]


def build_result(subject, status, file_path=""):
    return {
        "Subject": subject,
        "Status": status,
        "File": file_path
    }


def generate_report(excel_service, results, daily_root):
    report_file = os.path.join(daily_root, "Export_Report.xlsm")
    excel_service.export(results, output_file=report_file)
    logging.info(f"FINAL REPORT CREATED: {report_file}")