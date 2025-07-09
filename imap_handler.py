import imaplib
from  config import  imap_server,imap_port,my_email,my_password



# Define a function to connect to the server
def connect():
    try:
        imap = imaplib.IMAP4_SSL(imap_server, imap_port)
        imap.login(my_email, my_password)
        imap.select("Inbox")
        return imap
    except imaplib.IMAP4.error:
        print("Login failed. Please check your email and password.")
    except Exception as e:
        print("Error connecting to the server.", e)
        return None


# Function to search the required email ids
def search_emails(imap):
    subject_keyword = input("Enter part of the subject to search for: ").strip()

    _, mail_ids = imap.search(None, 'UNSEEN', 'FROM', '"xyz@domain.com"', 'SUBJECT',f'"{subject_keyword}"')

    return [mail_id.decode() for mail_id in mail_ids[0].split()]

