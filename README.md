# Report_automation
Using python and it's libraries this project is build to read emails containing attachments which are html based file with .xls extension. 
📦 Python Libraries & Technologies Used
1. Email Handling
imaplib – Connects to the email server using the IMAP protocol to read emails.

smtplib – Sends emails using the SMTP protocol.

email – Parses and creates email messages (email.message.EmailMessage).

email.message_from_bytes() – Converts raw byte data into a readable email object.

2. Environment Management
dotenv (python-dotenv) – Loads configuration variables from a .env file to keep sensitive info (email credentials, server configs) safe and configurable.

3. File Handling & Conversion
os – File and path operations (checking existence, removing files).

magic (usually python-magic) – Detects MIME types of files to determine proper attachment handling.

win32com.client (from pywin32) – Interacts with Windows COM objects. Used here to control Microsoft Excel for converting files to .xlsx.

4. Data Handling
pandas – Reads email lists from a CSV file (emails.csv).

5. Time & Network
time – Adds delays, helpful for sending emails smoothly.

socket – Handles low-level network error exceptions during SMTP communication.

