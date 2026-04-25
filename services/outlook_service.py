import win32com.client
import os


def outlookemails(subject_filter, download_folder):
    os.makedirs(download_folder, exist_ok=True)
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    messages_found = False
    for message in messages:
        try:
        # Ensure it's a real email item
            subject = str(getattr(message, "Subject", ""))
            if subject_filter.lower() in subject.lower():
                messages_found = True
                attachments = message.Attachments
                for att in attachments:
                    filename = att.FileName.lower()
                if filename.endswith((".xlsx", ".xls", ".csv", ".pdf")):
                    file_path = os.path.join(download_folder, att.FileName)
                    att.SaveASFile(file_path)
                    print(f"Saved file: {att.FileName}")
        except Exception as e:
            print("Skipping invalid item:", e)
    return messages_found


 