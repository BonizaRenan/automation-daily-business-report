import win32com.client
import logging
import re
import os


class OutlookService:

    def normalize_subject(self, subject: str) -> str:
        if not subject:
            return ""

        subject = subject.lower()
        subject = re.sub(r"^(fw:|fwd:|re:)\s*", "", subject, flags=re.IGNORECASE)

        return subject.strip()

    def download_emails(self, subject_filter, download_folder):
        logging.info(f"Downloading emails for subject: {subject_filter}")

        os.makedirs(download_folder, exist_ok=True)

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)

        messages = inbox.Items
        messages_found = False

        subject_filter_norm = subject_filter.lower()
        valid_extensions = (".xlsx", ".xls", ".csv", ".pdf", ".xlsm", ".xlsb")

        # ✅ Log scanning ONLY ONCE
        if not hasattr(self, "_scan_logged"):
            logging.info(f"Scanning {len(messages)} emails...")
            self._scan_logged = True

        for message in messages:
            try:
                subject = str(getattr(message, "Subject", ""))
                normalized_subject = self.normalize_subject(subject)

                if subject_filter_norm in normalized_subject:
                    messages_found = True

                    attachments = message.Attachments
                    logging.info(f"Matched email: {subject} | Attachments: {attachments.Count}")

                    if attachments.Count == 0:
                        logging.warning(f"No attachments in: {subject}")

                    for att in attachments:
                        file_name = att.FileName


                        if file_name.lower().endswith(valid_extensions):
                            file_path = os.path.join(download_folder, file_name)

                            att.SaveASFile(file_path)

                            logging.info(f"Saved → {file_path}")
                        else:
                            logging.warning(f"Skipped file type: {file_name}")

            except Exception as e:
                logging.error(f"Error processing email: {e}")

        if not messages_found:
            logging.warning(f"No emails found for subject: {subject_filter}")

        return messages_found