from datetime import datetime
import pandas as pd
import win32com.client as win32
import logging



class ReportEmailService:

    # =====================================================
    # TITLE ROW
    # =====================================================
    def build_title_row(self, report_date):

        return f"""
        <tr>
            <td colspan='3'
                style='background-color:#70AD47;
                color:white;
                font-weight:bold;
                padding:6px;
                border:1px solid black;'>
                Daily Business Report
            </td>

            <td
                style='background-color:#70AD47;
                color:white;
                font-weight:bold;
                padding:6px;
                border:1px solid black;
                text-align:right;'>
                Report Date :
            </td>

            <td
                style='background-color:#70AD47;
                color:white;
                font-weight:bold;
                padding:6px;
                border:1px solid black;
                text-align:center;'>
                {report_date.strftime("%m/%d/%Y")}
            </td>
        </tr>
        """


    # =====================================================
    # GENERIC ROW BUILDER
    # =====================================================
    def build_row(self, row, is_header=False):

        tag = "th" if is_header else "td"

        if is_header:
            style = """
                background-color:#FFF2CC;
                font-weight:bold;
                padding:6px;
                text-align:left;
                border:1px solid black;
            """
        else:
            style = """
                padding:6px;
                border:1px solid black;
            """

        html = "<tr>"

        for value in row:

            if pd.isna(value):
                value = ""

            html += f"""
            <{tag} style='{style}'>
                {value}
            </{tag}>
            """

        html += "</tr>"

        return html


    # =====================================================
    # DATAFRAME TO HTML
    # =====================================================
    def dataframe_to_html(self, df):

        html = ""

        # HEADER
        html += self.build_row(df.columns, True)

        # ROWS
        for _, row in df.iterrows():
            html += self.build_row(row)

        return html


    # =====================================================
    # SEND EMAIL
    # =====================================================
    def send_email(self):

        today = datetime.now()

        # =====================================================
        # EXCEL FILE
        # =====================================================
        excel_file = (
            fr"shared\download"
            fr"\Daily Business Report {today.strftime('%m-%d-%Y')}"
            fr"\Daily_Business_Report-{today.strftime('%m-%d-%Y')}.xlsx"
        )

        # =====================================================
        # MAIN TABLE
        # =====================================================
        df_main = pd.read_excel(
            excel_file,
            sheet_name=0,
            usecols="A:E",
            skiprows=1,
            nrows=11
        )

        # =====================================================
        # SECOND TABLE
        # =====================================================
        df_app = pd.read_excel(
            excel_file,
            sheet_name=0,
            usecols="A:B",
            skiprows=14,
            nrows=2,
            header=None
        )

        # =====================================================
        # RENAME COLUMNS
        # =====================================================
        df_main.columns = [
            "Name of Task",
            "Team",
            "File Name",
            "Link to Report Folder",
            "Frequency"
        ]

        # =====================================================
        # BUILD HTML BODY
        # =====================================================
        str_body = """
        <html>
        <body style='font-family:Calibri;font-size:11pt;'>
        """

        # =====================================================
        # FIRST TABLE
        # =====================================================
        str_body += """
        <table
            style='border-collapse:collapse;
            font-family:Calibri;
            font-size:11pt;'
            border='1'>
        """

        str_body += self.build_title_row(today)

        str_body += self.dataframe_to_html(df_main)

        str_body += "</table>"

        str_body += "<br><br>"

        # =====================================================
        # SECOND TABLE
        # =====================================================
        str_body += """
        <table
            style='border-collapse:collapse;
            font-family:Calibri;
            font-size:11pt;'
            border='1'>
        """

        # HEADER
        str_body += """
        <tr>

            <th style='background-color:#FFF2CC;
            padding:6px;
            border:1px solid black;'>
                Application Virtual Date
            </th>

            <th style='background-color:#FFF2CC;
            padding:6px;
            border:1px solid black;'>
                Team
            </th>

            <th style='background-color:#FFF2CC;
            padding:6px;
            border:1px solid black;'>
                Date
            </th>

        </tr>
        """

        # DATA ROWS
        for _, row in df_app.iterrows():

            application_name = row[0]
            team_name = row[1]

            str_body += f"""
            <tr>

                <td style='padding:6px;border:1px solid black;'>
                    {application_name}
                </td>

                <td style='padding:6px;border:1px solid black;'>
                    {team_name}
                </td>

                <td style='padding:6px;border:1px solid black;text-align:center;'>
                    {today.strftime("%m/%d/%Y")}
                </td>

            </tr>
            """

        str_body += "</table>"

        # =====================================================
        # CLOSE HTML
        # =====================================================
        str_body += """
        </body>
        </html>
        """

        # =====================================================
        # OUTLOOK
        # =====================================================
        outlook = win32.Dispatch("Outlook.Application")

        mail = outlook.CreateItem(0)

         
        mail.To = "cookiesncream2026@outlook.com; paulotrevordregacho@outlook.com"

        mail.Subject = (
            f"Daily Business Report - "
            f"{today.strftime('%m/%d/%Y')}"
        )

        mail.HTMLBody = str_body

        # SHOW EMAIL
        mail.Display()

        # AUTO SEND
        # mail.Send()