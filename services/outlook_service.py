import win32com.client
import os
import logging


class OutlookService:

    def download_emails(self, subject_filter, download_folder):
        logging.info(f"Downloading emails for subject: {subject_filter}")

        os.makedirs(download_folder, exist_ok=True)

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items

        messages_found = False

        for message in messages:
            try:
                subject = str(getattr(message, "Subject", ""))

                if subject_filter.lower() in subject.lower():
                    messages_found = True
                    attachments = message.Attachments

                    for att in attachments:
                        filename = att.FileName.lower()

                        if filename.endswith((".xlsx", ".xls", ".csv", ".pdf")):
                            file_path = os.path.join(download_folder, att.FileName)
                            att.SaveASFile(file_path)

                            logging.info(f"Saved file: {att.FileName}")

            except Exception as e:
                logging.error(f"Skipping invalid item: {e}")

        if not messages_found:
            logging.warning(f"No emails found for subject: {subject_filter}")

        return messages_found