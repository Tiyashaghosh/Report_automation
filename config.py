from dotenv import load_dotenv, find_dotenv
import os



dotenv_path = find_dotenv()
load_dotenv(dotenv_path)
my_email = os.getenv('EMAIL')  # Load the email it will read mails from
my_password = os.getenv('APP_PASSWORD')  # Load the app password
imap_server = os.getenv('IMAP_SERVER')  # Imap server the code connects to
imap_port = int(os.getenv('IMAP_PORT'))
smtp_server = os.getenv('SMTP_SERVER')    # Smtp server the code connects to
smtp_port = int(os.getenv('SMTP_PORT'))   # Smtp port - 587
download_folder = os.getenv('DOWNLOAD_PATH')