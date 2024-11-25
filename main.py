# import library that we need in this project
import os
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import time
import tkinter as tk
from tkinter import filedialog, messagebox

# SMTP server configuration for Gmail
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 465

# Function to schedule email sending at a specific date and time
def schedule_email(send_datetime_input):
    send_datetime = None
    if len(send_datetime_input) == 16:  # Format YYYY-MM-DD HH:MM
        send_datetime = datetime.strptime(send_datetime_input, "%Y-%m-%d %H:%M")
    elif len(send_datetime_input) == 19:  # Format YYYY-MM-DD HH:MM:SS
        send_datetime = datetime.strptime(send_datetime_input, "%Y-%m-%d %H:%M:%S")

    if send_datetime is None:
        messagebox.showerror("Error", "Invalid date-time format. Please use YYYY-MM-DD HH:MM or YYYY-MM-DD HH:MM:SS.")
        return

    while datetime.now() < send_datetime:
        time.sleep(1)
    print("It's time to send the emails!")

# Function to send a single email with a custom subject
def send_email(sender_email, sender_pass, receiver_email, name, subject, attachment_path):
    msg = MIMEMultipart("alternative")
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject

    # Email body
    text = f"Hello {name},\n\nI have something for you. Please find the attachment."
    html = f"""
    <html>
        <body>
            <p>Hello {name},</p>
            <p>I have something for you. Please find the attachment.</p>
            <p>Best regards,<br>Your Team</p>
        </body>
    </html>
    """
    msg.attach(MIMEText(text, "plain"))
    msg.attach(MIMEText(html, "html"))

    # Attach file
    if not os.path.exists(attachment_path):
        print(f"Attachment NOT found: {attachment_path}")
        return

    file_part = MIMEBase('application', 'octet-stream')
    f = open(attachment_path, 'rb')
    file_part.set_payload(f.read())
    encoders.encode_base64(file_part)
    file_part.add_header(
        'Content-Disposition',
        f'attachment; filename="{os.path.basename(attachment_path)}"'
    )
    msg.attach(file_part)
    f.close()

    # Send email
    smtp = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
    smtp.login(sender_email, sender_pass)
    smtp.send_message(msg)
    smtp.quit()

    print(f"Email sent successfully to {receiver_email}.")

# Function to send bulk emails
def send_bulk_emails(sender_email, sender_pass, excel_file):
    base_directory = os.path.dirname(excel_file)
    df = pd.read_excel(excel_file)

    # Define the required columns
    required_columns = {"EMAIL_ID", "NAME", "SUBJECT", "Files to be attached"}
    if not required_columns.issubset(df.columns):
        missing_columns = required_columns - set(df.columns)
        messagebox.showerror("Error", f"Excel file is missing required columns: {', '.join(missing_columns)}")
        return

    for _, row in df.iterrows():
        receiver_email = row['EMAIL_ID']
        name = row['NAME']
        subject = row['SUBJECT']
        attachment = row['Files to be attached']
        attachment_path = os.path.join(base_directory, attachment)

        if not os.path.exists(attachment_path):
            print(f"Attachment NOT found: {attachment_path}. Skipping email to {receiver_email}.")
            continue

        send_email(sender_email, sender_pass, receiver_email, name, subject, attachment_path)

# Function to handle GUI inputs and execute the main email function
def gui_send_emails():
    sender_email = entry_sender_email.get()
    sender_pass = entry_sender_pass.get()
    send_date = entry_send_date.get()
    send_time = entry_send_time.get()
    send_datetime_input = f"{send_date} {send_time}"

    schedule_email(send_datetime_input)  # Schedule email sending
    mode = var_mode.get()  # Get mode (one or many)

    if mode == "one":
        receiver_email = entry_receiver_email.get()
        name = entry_receiver_name.get()
        subject = entry_subject.get()
        attachment_path = entry_attachment_path.get()

        send_email(sender_email, sender_pass, receiver_email, name, subject, attachment_path)
        messagebox.showinfo("Success", "Email sent successfully!")

    elif mode == "many":
        excel_file = entry_excel_file.get()
        send_bulk_emails(sender_email, sender_pass, excel_file)
        messagebox.showinfo("Success", "Bulk emails sent successfully!")

    else:
        messagebox.showerror("Error", "Invalid mode selected. Please choose 'one' or 'many'.")

# Function to open file dialog for selecting attachment
def browse_attachment():
    file_path = filedialog.askopenfilename()
    entry_attachment_path.delete(0, tk.END)
    entry_attachment_path.insert(0, file_path)

# Function to open file dialog for selecting Excel file
def browse_excel():
    file_path = filedialog.askopenfilename()
    entry_excel_file.delete(0, tk.END)
    entry_excel_file.insert(0, file_path)

# Set up the Tkinter window
window = tk.Tk()
window.title("Email Automation Script")

# Create and place GUI elements
tk.Label(window, text="Sender Email:").grid(row=0, column=0)
entry_sender_email = tk.Entry(window, width=40)
entry_sender_email.grid(row=0, column=1)

tk.Label(window, text="Sender Password:").grid(row=1, column=0)
entry_sender_pass = tk.Entry(window, width=40, show="*")
entry_sender_pass.grid(row=1, column=1)

tk.Label(window, text="Send Date (YYYY-MM-DD):").grid(row=2, column=0)
entry_send_date = tk.Entry(window, width=40)
entry_send_date.grid(row=2, column=1)

tk.Label(window, text="Send Time (HH:MM):").grid(row=3, column=0)
entry_send_time = tk.Entry(window, width=40)
entry_send_time.grid(row=3, column=1)

tk.Label(window, text="Mode (one/many):").grid(row=4, column=0)
var_mode = tk.StringVar(value="one")
tk.Radiobutton(window, text="One", variable=var_mode, value="one").grid(row=4, column=1)
tk.Radiobutton(window, text="Many", variable=var_mode, value="many").grid(row=4, column=2)

# Input fields for single email
tk.Label(window, text="Receiver Email:").grid(row=5, column=0)
entry_receiver_email = tk.Entry(window, width=40)
entry_receiver_email.grid(row=5, column=1)

tk.Label(window, text="Receiver Name:").grid(row=6, column=0)
entry_receiver_name = tk.Entry(window, width=40)
entry_receiver_name.grid(row=6, column=1)

tk.Label(window, text="Subject:").grid(row=7, column=0)
entry_subject = tk.Entry(window, width=40)
entry_subject.grid(row=7, column=1)

tk.Label(window, text="Attachment Path:").grid(row=8, column=0)
entry_attachment_path = tk.Entry(window, width=40)
entry_attachment_path.grid(row=8, column=1)
tk.Button(window, text="Browse", command=browse_attachment).grid(row=8, column=2)

# Input fields for bulk email
tk.Label(window, text="Excel File (for many):").grid(row=9, column=0)
entry_excel_file = tk.Entry(window, width=40)
entry_excel_file.grid(row=9, column=1)
tk.Button(window, text="Browse", command=browse_excel).grid(row=9, column=2)

# Send email button
tk.Button(window, text="Send Emails", command=gui_send_emails).grid(row=10, column=0, columnspan=3)

# Start the Tkinter event loop
window.mainloop()
