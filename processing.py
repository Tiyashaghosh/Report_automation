from attachments import save_attachments,conversion,freeze_panes
from smtp_handler import send_email
import pandas as pd
import os

def process_email(message):  # Function which reads the mails, save attachments and converts them accordingly
    try:
        print("-----------------------------------")
        print(f"From : {message.get('From')}")
        print(f"To : {message.get('To')}")
        print(f"Bcc : {message.get('Bcc')}")
        print(f"Date : {message.get('Date')}")
        print(f"Subject: {message.get('Subject')}")

        body_data = None
        converted_files = []

        for part in message.walk():
            # Read the body of the email
            if body_data is None and (
                    part.get_content_type() in ['text/plain', 'text/html']) and part.get_content_disposition() is None:
                body_data = part.get_payload(decode=True).decode(errors='ignore')
                print("Body: \n", body_data)


            if part.get_filename():
                original_file_path = save_attachments(part, part.get_filename())
                print(f"✅ {original_file_path}")
                if original_file_path:
                    converted = conversion(file_path=original_file_path)
                    converted = freeze_panes(converted)
                    if converted:
                        converted_files.append(converted)

        if converted_files:
            while True:
                file_name = input("Please provide the recipient list file name (provide the extension as well): ").strip()

                if not file_name:
                    print("File name cannot be empty. Please input a valid file name.")
                    continue
                if not os.path.isfile(file_name):
                    print("File doesn't exist. Try again.")
                    continue

                emails = (pd.read_csv(file_name)["email_id"].dropna().drop_duplicates()).tolist()
                break
            send_email(
                email_list=emails,
                subject=message['Subject'],
                body=body_data or " ",
                attachment_list=converted_files
            )

    except Exception as e:
        print("Error processing message:", e)
