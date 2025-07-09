import smtplib
import socket
from email.message import EmailMessage
import time
from config import smtp_server,smtp_port,my_email,my_password
import os
import magic



# Function which will send emails
def send_email(email_list, subject, body, attachment_list):
    try:
        # Connect to the SMTP server
        with smtplib.SMTP(smtp_server, smtp_port) as connection:
            connection.starttls()

            try:
                connection.login(my_email, my_password)
            except smtplib.SMTPAuthenticationError:
                print("Login failed. Please check your email or app password.")
                return
            except smtplib.SMTPException as e:
                print(f"Unexpected SMTP login error: {e}")
                return

            for email_id in email_list:
                try:
                    msg = EmailMessage()
                    msg['Subject'] = ''.join(subject.splitlines())
                    msg['From'] = my_email
                    msg['To'] = email_id
                    msg.set_content(body)

                    for attachment in attachment_list:
                        try:
                            maintype, subtype = magic.from_file(attachment, mime=True).split('/')
                            filename = os.path.basename(attachment)

                            with open(attachment, 'rb') as f:
                                file_data = f.read()

                            msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=filename)
                        except FileNotFoundError:
                            print(f"Attachment not found: {attachment}")
                        except Exception as e:
                            print(f"Failed to attach file {attachment}: {e}")

                    connection.send_message(msg)
                    time.sleep(1)
                    print(f"Email sent to: {email_id}")

                except smtplib.SMTPRecipientsRefused:
                    print(f"Recipient refused: {email_id}")
                except smtplib.SMTPDataError as e:
                    print(f"SMTP data error for {email_id}: {e}")
                except Exception as inner_e:
                    print(f"Failed to send to {email_id}: {inner_e}")

    except socket.gaierror:
        print("Network error: Unable to resolve SMTP server address.")
    except socket.timeout:
        print("Network timeout while connecting to SMTP server.")
    except smtplib.SMTPConnectError:
        print("Failed to connect to SMTP server.")
    except Exception as e:
        print(f"SMTP connection or setup failed: {e}")