import pandas as pd

def export_to_excel(data, output_file="utils/download/Outlook_Report.xlsx"):
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)
    print(f"Excel generated: {output_file}")