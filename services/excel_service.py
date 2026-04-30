import logging
import pandas as pd


class ExcelService:

    def export(self, data, output_file="shared/download/Export_Report.xlsx"):
        logging.info("Exporting results to Excel")

        df = pd.DataFrame(data)
        df.to_excel(output_file, index=False)

        logging.info(f"Export completed: {output_file}")