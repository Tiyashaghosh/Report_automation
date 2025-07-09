📩 Report Automation
This project uses Python and its libraries to automatically read emails containing attachments—specifically .xls files that are actually HTML-based reports.
The script downloads these attachments, converts them to .xlsx, freezes pane and forwards them to a user chosen list of recipients.

📦ech Stack
Python 3.x

IMAP & SMTP (for email handling)

openpyxl, pandas (for Excel operations)

win32com (Excel automation on Windows)

dotenv (for environment variable management)

python-magic (for MIME type detection)
## 🚀 Features

- Connects to IMAP and SMTP servers securely using credentials from `.env`.
- Searches unseen emails based on a subject keyword and sender domain.
- Downloads and saves email attachments.
- Converts legacy `.xls` Excel files to `.xlsx` using `win32com`.
- Prompts the user to freeze panes (like "A1", "B5") in the processed Excel file.
- Sends the final files to a user-defined list of recipients (from a CSV).
- Marks processed emails as **read** to avoid reprocessing.

  ## 🧠 Project Structure

📦project-root
┣ 📄attachments.py # Download & convert email attachments
┣ 📄config.py # Loads environment variables from .env
┣ 📄imap_handler.py # Handles IMAP connection & email fetching
┣ 📄smtp_handler.py # Sends email using SMTP with attachments
┣ 📄processing.py # Main email parsing and file handling logic
┣ 📄main.py # Script entry point
┣ 📄Requirements.txt # Python dependencies
┗ 📄.env # Your local environment configuration

Install the dependencies
pip install -r Requirements.txt

Set up your .env file
EMAIL=your_email@gmail.com
APP_PASSWORD=your_app_password
IMAP_SERVER=imap.gmail.com
IMAP_PORT=993
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
DOWNLOAD_PATH=absolute_or_relative_path_to_save_files

Sample Email CSV Format
email_id
john@example.com
jane@example.com

❗Important Notes
This script works only on Windows due to the use of win32com.client.

It expects email attachments in Excel .xls format.

Your environment must have Microsoft Excel installed for conversion.

