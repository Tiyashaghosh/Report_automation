# Report_automation
Using python and it's libraries this project is build to read emails containing attachments which are html based file with .xls extension. 
📦 Python Libraries & Technologies Used
📬 Email Handling
imaplib – Connects to the email server using the IMAP protocol to read emails.

smtplib – Sends emails using the SMTP protocol.

email – Parses and creates email messages using email.message.EmailMessage.

email.message_from_bytes() – Converts raw byte data into a readable email object.

🔐 Environment Management
python-dotenv (dotenv) – Loads configuration variables from a .env file (e.g., credentials, server configs).

📁 File Handling & Conversion
os – Handles file and path operations (like checking existence, deleting files).

python-magic (magic) – Detects MIME types of attachments to ensure correct handling.

pywin32 (win32com.client) – Automates Microsoft Excel via COM interface to convert file formats (e.g., .xls to .xlsx). Requires Windows and MS Office.

📊 Data Handling
pandas – Reads and processes email lists from a CSV file (emails.csv).

⏱️ Timing & Network
time – Adds delays between operations (e.g., email sending).

socket – Manages network-level exceptions (e.g., SMTP resolution failures).
