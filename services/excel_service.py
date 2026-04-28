from openpyxl import load_workbook, Workbook
from datetime import datetime
import logging
import pandas as pd
import os


class ExcelService:

    def clean_value(self, value):
        if value == []:
            return ""
        if isinstance(value, list):
            return ", ".join(map(str, value))
        if isinstance(value, dict):
            return str(value)
        return value


    def export(self, data, output_file=None):
        logging.info("Exporting results to Excel")

        df = pd.DataFrame(data)

        if output_file is None:
            today = datetime.now().strftime("%m-%d-%Y")
            folder_name = f"Daily Business Report {today}"
            output_folder = os.path.join("shared", "download", folder_name)
            os.makedirs(output_folder, exist_ok=True)

            output_file = os.path.join(output_folder, "Export_Report.xlsm")

        template_path = "shared/download/template.xlsm"

        wb = load_workbook(template_path, keep_vba=True)
        ws = wb.active

        ws.delete_rows(1, ws.max_row)

        # headers
        for col_index, column in enumerate(df.columns, start=1):
            ws.cell(row=1, column=col_index, value=column)

        # data rows
        for row_index, row in enumerate(df.itertuples(index=False), start=2):
            for col_index, value in enumerate(row, start=1):
                ws.cell(
                    row=row_index,
                    column=col_index,
                    value=self.clean_value(value)
                )

        wb.save(output_file)

        logging.info(f"Export completed: {output_file}")


    def merge_carton_files(self, file_list, output_file):

        new_wb = Workbook()
        new_wb.remove(new_wb.active)

        for file in file_list:

            wb = load_workbook(file)
            ws = wb.active

            sheet_name = os.path.splitext(os.path.basename(file))[0][:31]

            new_ws = new_wb.create_sheet(title=sheet_name)

            for row in ws.iter_rows(values_only=True):
                new_ws.append(row)

        new_wb.save(output_file)