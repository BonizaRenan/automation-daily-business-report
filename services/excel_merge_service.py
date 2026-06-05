from openpyxl import load_workbook, Workbook
import logging
import os




class ExcelMergeService:

    # =========================
    # MERGE FILES
    # =========================
    def merge_files(self, file_list, output_file, config_item=None):

        priority_rules = (config_item or {}).get("priority_rules", {})

        def get_priority(file_path):
            name = os.path.basename(file_path).lower()

            for key, priority in priority_rules.items():
                if key.lower() in name:
                    return priority

            return 99

        file_list = sorted(file_list, key=get_priority)

        new_wb = Workbook()
        new_wb.remove(new_wb.active)

        for file in file_list:
            try:
                wb = load_workbook(file)
                ws = wb.active

                sheet_name = os.path.splitext(os.path.basename(file))[0][:31]
                new_ws = new_wb.create_sheet(title=sheet_name)

                for row in ws.iter_rows(values_only=True):
                    new_ws.append(row)

            except Exception as e:
                logging.error(f"Failed merging file {file}: {e}")

        new_wb.save(output_file)

        logging.info(f"Merged file created: {output_file}")