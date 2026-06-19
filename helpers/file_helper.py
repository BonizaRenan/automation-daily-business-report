import os


# =========================
# SAFE FOLDER CREATION (DAILY ROOT)
# =========================
def create_daily_folder(base_folder, folder_date):

    if not base_folder:
        raise ValueError("base_folder is empty or None")

    path = os.path.join(base_folder, f"Daily Business Report {folder_date}")

    os.makedirs(path, exist_ok=True)

    return os.path.abspath(path)


# =========================
# SAFE SUB FOLDER CREATION
# =========================
def create_folder(daily_root, item):

    if not daily_root:
        raise ValueError("daily_root is empty or None")

    folder_name = item.get("file_name")

    if not folder_name:
        raise ValueError(f"Missing file_name in config: {item}")

    # sanitize folder name (VERY IMPORTANT in Windows)
    folder_name = "".join(c for c in folder_name if c not in r'<>:"/\|?*')

    path = os.path.join(daily_root, folder_name)

    os.makedirs(path, exist_ok=True)

    return os.path.abspath(path)


# =========================
# VALID FILE CHECK
# =========================
def get_valid_files(folder):

    if not folder or not os.path.exists(folder):
        return []

    valid_ext = (".xlsx", ".pdf", ".csv", ".xlsm", ".xls")

    return [
        f for f in os.listdir(folder)
        if f.lower().endswith(valid_ext)
        and not f.startswith("~$")  # ignore temp excel files
    ]


# =========================
# EXCEL FILE FILTER (FOR MERGE)
# =========================
def get_excel_files(folder):

    if not folder or not os.path.exists(folder):
        return []

    valid_ext = (".xlsx", ".xlsm", ".xls", ".csv")

    files = []

    for f in os.listdir(folder):

        file_lower = f.lower()

        if not file_lower.endswith(valid_ext):
            continue

        # ignore merged + temp files
        if "merged" in file_lower:
            continue

        if f.startswith("~$"):
            continue

        files.append(os.path.join(folder, f))

    return files