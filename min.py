



import os
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import time

# Input sender's email and password securely
sender_email = input("Enter Your Email ID: ")
sender_pass = input("Enter Your Password: ")  # Secure password entry

# Specify the Excel file path (ensure it points to the actual file)
excel_file = r"C:\Users\Leader.Din\Desktop\kk\email.xlsx"  # Update with the correct filename

# Try loading the Excel file
try:
    df = pd.read_excel(excel_file)
except FileNotFoundError:
    print(f"Error: File not found at {excel_file}")
    exit()
except PermissionError:
    print(f"Error: Permission denied for file at {excel_file}")
    exit()
except Exception as e:
    print(f"Error: Unable to read the Excel file - {e}")
    exit()

# Check if required columns are present
required_columns = {"EMAIL_ID", "Files to be attached", "NAME"}
if not required_columns.issubset(df.columns):
    print(f"Error: Excel file must contain the columns: {', '.join(required_columns)}")
    exit()

# Extract data from the DataFrame
receivers_email = df["EMAIL_ID"].values
attachments = df["Files to be attached"].values  # Ensure paths in Excel are accurate
names = df["NAME"].values

# Create a zip object to iterate over rows
zipped = zip(receivers_email, attachments, names)

# SMTP server configuration
smtp_server = 'smtp.gmail.com'
smtp_port = 465

# Get the current working directory (where the script is running)
script_dir = os.path.dirname(os.path.realpath(__file__))

# Input the scheduled time
send_time_input = input("Enter the time to send emails (HH:MM or HH:MM:SS, 24-hour format): ")
try:
    # Attempt to parse with seconds
    send_time = datetime.strptime(send_time_input, "%H:%M:%S").time()
except ValueError:
    try:
        # Attempt to parse without seconds
        send_time = datetime.strptime(send_time_input, "%H:%M").time()
    except ValueError:
        print("Error: Invalid time format. Please use HH:MM or HH:MM:SS (24-hour format).")
        exit()

# Wait until the scheduled time
while datetime.now().time() < send_time:
    time.sleep(1)  # Check the time every second

# Process each recipient
for receiver, attachment, name in zipped:
    try:
        # Create the email message with both plain text and HTML
        msg = MIMEMultipart("alternative")
        msg['From'] = sender_email
        msg['To'] = receiver
        msg['Subject'] = "Test Mail"

        # Text version of the email
        text = f"Hello {name},\n\nI have something for you. Please find the attachment."

        # HTML version of the email
        html = f"""
        <html>
            <body>
                <p>Hello {name},</p>
                <p>I have something for you. Please find the attachment.</p>
                <p>Best regards,<br>Your Name</p>
            </body>
        </html>
        """

        # Attach both plain text and HTML
        msg.attach(MIMEText(text, "plain"))
        msg.attach(MIMEText(html, "html"))

        # Handle file attachment paths correctly
        attachment_path = os.path.join(script_dir, attachment)

        # Check if the file exists
        if not os.path.exists(attachment_path):
            # Try appending '.pdf' if no file extension is found
            attachment_path_with_extension = attachment_path + '.pdf'
            if os.path.exists(attachment_path_with_extension):
                attachment_path = attachment_path_with_extension
                print(f"Found attachment with '.pdf' extension: {attachment_path}")
            else:
                print(f"Attachment NOT found: {attachment_path}")  # Debugging line
                continue  # Skip to the next recipient if attachment is not found

        # Attach the file
        with open(attachment_path, 'rb') as f:
            file_part = MIMEBase('application', 'octet-stream')
            file_part.set_payload(f.read())
            encoders.encode_base64(file_part)
            file_part.add_header(
                'Content-Disposition',
                f'attachment; filename="{os.path.basename(attachment_path)}"'
            )
            msg.attach(file_part)

        # Send the email
        with smtplib.SMTP_SSL(smtp_server, smtp_port) as smtp:
            smtp.login(sender_email, sender_pass)
            smtp.send_message(msg)

        print(f"Email sent successfully to {receiver}.")
    except Exception as e:
        print(f"Error sending email to {receiver}: {e}")

print("All emails have been processed.")
