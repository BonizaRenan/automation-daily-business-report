import win32com.client
import logging
import re
import os
from datetime import datetime, timedelta


class OutlookService:

    def normalize_subject(self, subject: str) -> str:
        if not subject:
            return ""

        subject = subject.lower()
        subject = re.sub(r"^(fw:|fwd:|re:)\s*", "", subject, flags=re.IGNORECASE)
        return subject.strip()

    def safe_filename(self, name: str) -> str:
        return re.sub(r'[\\/*?:"<>|]', "_", name)

    def download_emails(self, subject_filter, download_folder):

        logging.info(f"Downloading emails for subject: {subject_filter}")

        # ✅ ALWAYS ensure folder exists (CRITICAL FIX)
        os.makedirs(download_folder, exist_ok=True)

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)

        items = inbox.Items
        items.Sort("[ReceivedTime]", True)

        cutoff = datetime.now() - timedelta(days=2)
        cutoff = cutoff.replace(tzinfo=None)

        subject_filter_norm = self.normalize_subject(subject_filter)

        messages_found = False

        logging.info("Scanning emails (2-day window only)...")

        for i in range(1, items.Count + 1):

            try:
                message = items.Item(i)

                msg_time = message.ReceivedTime
                msg_time = msg_time.replace(tzinfo=None)

                # ✅ STOP EARLY (PERFORMANCE BOOST)
                if msg_time < cutoff:
                    break

                subject = str(getattr(message, "Subject", "") or "")
                normalized_subject = self.normalize_subject(subject)

                if subject_filter_norm not in normalized_subject:
                    continue

                messages_found = True

                attachments = message.Attachments

                logging.info(
                    f"Matched: {subject} | "
                    f"{msg_time} | Attachments: {attachments.Count}"
                )

                if attachments.Count == 0:
                    continue

                for att in attachments:

                    file_name = self.safe_filename(att.FileName)

                    file_path = os.path.join(download_folder, file_name)

                    # overwrite protection
                    if os.path.exists(file_path):
                        base, ext = os.path.splitext(file_name)
                        timestamp = datetime.now().strftime("%H%M%S")

                        file_path = os.path.join(
                            download_folder,
                            f"{base}_{timestamp}{ext}"
                        )

                    try:
                        att.SaveAsFile(file_path)
                        logging.info(f"Saved → {file_path}")

                    except Exception as e:
                        logging.error(f"Failed to save attachment {file_name}: {e}")

            except Exception as e:
                logging.error(f"Error processing email: {e}")

        if not messages_found:
            logging.warning(f"No emails found for subject: {subject_filter}")

        return messages_found