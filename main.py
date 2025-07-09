import email
from imap_handler import connect,search_emails
from processing import process_email
from email.utils import parsedate_to_datetime
def main():
    imap = connect()
    if not imap:
        print("Could not connect to server.")
        return

    try:
        email_ids = search_emails(imap)
        print(f"Found {len(email_ids)} emails.")

        latest_time = None
        latest_email = None
        latest_mail_id =  None
        for mail_id in email_ids:
            try:
                status, msg_data = imap.fetch(mail_id, '(RFC822)')
                if status == 'OK':
                    raw_email = msg_data[0][1]  # Tuple in a list
                    message = email.message_from_bytes(raw_email)
                    date = parsedate_to_datetime(message['date'])
                    if latest_time is None or date > latest_time:
                        latest_email = message
                        latest_time = date
                        latest_mail_id = mail_id
            except Exception as inner_e:
                print(f"Failed to process email ID {mail_id}: {inner_e}")
        process_email(latest_email)
        imap.store(latest_mail_id, '+FLAGS',
                   '\\Seen')  # After processing the email it marks it as read to not process it again
    finally:
        try:
            imap.logout()
        except Exception as logout_err:
            print("Error during logout:", logout_err)


if __name__ == "__main__":
    main()
