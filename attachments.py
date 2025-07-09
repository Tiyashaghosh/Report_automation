import win32com.client as win32
import os
import time
from openpyxl.utils import coordinate_to_tuple
from config import download_folder
import openpyxl

# Function to save email attachments in their original format
def save_attachments(part, file_name):
    try:
        original_file_path = os.path.join(download_folder, file_name)
        if not os.path.exists(original_file_path):
            with open(original_file_path, 'wb') as f:
                f.write(part.get_payload(decode=True))
                f.flush()
                time.sleep(0.5)
            return original_file_path
        else:
            print(f"File already exists: {file_name}")
            return original_file_path  # or return None if you want to skip
    except Exception as e:
        print(f"Error saving attachment: {file_name}", e)
        return None


# Function to convert these attachments
def conversion(file_path):
    if not os.path.exists(file_path):
        print(f"File does not exist: {file_path}")
        return None

    # file_name = os.path.basename(file_path)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False  # Suppress extension mismatch popup

    try:
        print(f"✅ Opening file : {file_path}")
        wb = excel.Workbooks.Open(file_path)
        new = os.path.splitext(file_path)[0]
        # print(new)
        new_path = new + '.xlsx'
        wb.SaveAs(new_path, FileFormat=51)
        wb.Close(SaveChanges=False)
        os.remove(file_path)
        print(f"✅ Converted successfully to: {new_path}")
        return new_path
    except Exception as e:
        print(f"Excel conversion failed: {e}")
    finally:
        excel.Quit()

def freeze_panes(converted_file_path):
    wb = openpyxl.load_workbook(converted_file_path)
    ws = wb.active
    while True:
        user_input = input(f"Please enter the cell number to freeze for the file : {os.path.basename(converted_file_path)}: ")
        try:
            coordinate_to_tuple(user_input.upper())
            ws.freeze_panes = user_input.upper()
            break
        except ValueError:
            print("Invalid cell reference.Please try again (A1,B5).")


    wb.save(converted_file_path)
    print(f"Successfully froze panes at : {user_input.upper()} for file {os.path.basename(converted_file_path)}" )
    return converted_file_path

# def recipient_list(converted_file_path):
#     file_name = os.path.basename(converted_file_path)
#     emails = input(f"Who do you want to send this file to? {file_name} (Please enter the emails with a space): ")
#     receiver_list = [email.strip() for email in emails.split() if email.strip()]
#     return receiver_list