from openpyxl import load_workbook, Workbook
import logging
import pandas as pd
import os


class ExcelService:

    def export(self, data, output_file="shared/download/Export_Report.xlsx"):
        logging.info("Exporting results to Excel")

        df = pd.DataFrame(data)
        df.to_excel(output_file, index=False)

        logging.info(f"Export completed: {output_file}")

        
    
    def merge_carton_files(self, file_list, output_file):

        new_wb = Workbook()
        new_wb.remove(new_wb.active)  # remove default sheet

        for file in file_list:

            wb = load_workbook(file)
            ws = wb.active

            sheet_name = os.path.splitext(os.path.basename(file))[0][:31]

            new_ws = new_wb.create_sheet(title=sheet_name)

            for row in ws.iter_rows(values_only=True):
                new_ws.append(row)

            new_wb.save(output_file)