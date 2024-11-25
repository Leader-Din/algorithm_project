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
from threading import Thread

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
    with open(attachment_path, 'rb') as f:
        file_part.set_payload(f.read())
    encoders.encode_base64(file_part)
    file_part.add_header(
        'Content-Disposition',
        f'attachment; filename="{os.path.basename(attachment_path)}"'
    )
    msg.attach(file_part)

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

import tkinter as tk
from tkinter import filedialog, messagebox

# Function to handle GUI inputs and execute the main email function
def gui_send_emails():
    sender_email = entry_sender_email.get()
    sender_pass = entry_sender_pass.get()
    send_date = entry_send_date.get()
    send_time = entry_send_time.get()
    send_datetime_input = f"{send_date} {send_time}"

    # Display user input for scheduling in terminal
    print(f"Scheduled send datetime: {send_datetime_input}")
    print(f"Sender Email: {sender_email}")
    print(f"Sender Password: {sender_pass}")
    print(f"Send Date: {send_date}")
    print(f"Send Time: {send_time}")
    print(f"Mode: {var_mode.get()}")  # Displaying the mode (one or many)

    schedule_email(send_datetime_input)  # Schedule email sending
    mode = var_mode.get()  # Get mode (one or many)

    if mode == "one":
        receiver_email = entry_receiver_email.get()
        name = entry_receiver_name.get()
        subject = entry_subject.get()
        attachment_path = entry_attachment_path.get()

        # Display user input for single email in terminal
        print(f"Sending single email to {receiver_email}")
        print(f"Subject: {subject}")
        print(f"Attachment path: {attachment_path}")

        send_email(sender_email, sender_pass, receiver_email, name, subject, attachment_path)
        messagebox.showinfo("Success", "Email sent successfully!")

    elif mode == "many":
        excel_file = entry_excel_file.get()

        # Display user input for bulk email in terminal
        print(f"Sending bulk emails using Excel file: {excel_file}")

        send_bulk_emails(sender_email, sender_pass, excel_file)
        messagebox.showinfo("Success", "Bulk emails sent successfully!")

    else:
        messagebox.showerror("Error", "Invalid mode selected. Please choose 'one' or 'many'.")

import tkinter as tk
from tkinter import filedialog, messagebox

# Function to browse attachment file
def browse_attachment():
    file_path = filedialog.askopenfilename()
    entry_attachment_path.delete(0, tk.END)
    entry_attachment_path.insert(0, file_path)

# Function to browse excel file
def browse_excel():
    file_path = filedialog.askopenfilename()
    entry_excel_file.delete(0, tk.END)
    entry_excel_file.insert(0, file_path)

# Function to handle GUI inputs and execute the main email function
def gui_send_emails():
    sender_email = entry_sender_email.get()
    sender_pass = entry_sender_pass.get()
    send_date = entry_send_date.get()
    send_time = entry_send_time.get()
    send_datetime_input = f"{send_date} {send_time}"

    # Display user input for scheduling in terminal
    print(f"Scheduled send datetime: {send_datetime_input}")
    print(f"Sender Email: {sender_email}")
    print(f"Sender Password: {sender_pass}")
    print(f"Send Date: {send_date}")
    print(f"Send Time: {send_time}")
    print(f"Mode: {var_mode.get()}")  # Displaying the mode (one or many)

    schedule_email(send_datetime_input)  # Schedule email sending
    mode = var_mode.get()  # Get mode (one or many)

    if mode == "one":
        receiver_email = entry_receiver_email.get()
        name = entry_receiver_name.get()
        subject = entry_subject.get()
        attachment_path = entry_attachment_path.get()

        # Display user input for single email in terminal
        print(f"Sending single email to {receiver_email}")
        print(f"Subject: {subject}")
        print(f"Attachment path: {attachment_path}")

        send_email(sender_email, sender_pass, receiver_email, name, subject, attachment_path)
        messagebox.showinfo("Success", "Email sent successfully!")

    elif mode == "many":
        excel_file = entry_excel_file.get()

        # Display user input for bulk email in terminal
        print(f"Sending bulk emails using Excel file: {excel_file}")

        send_bulk_emails(sender_email, sender_pass, excel_file)
        messagebox.showinfo("Success", "Bulk emails sent successfully!")

    else:
        messagebox.showerror("Error", "Invalid mode selected. Please choose 'one' or 'many'.")

# Set up the Tkinter window
window = tk.Tk()
window.title("Email Automation Script")
icon_path = "image.png"  # Replace with the correct icon file path
if os.path.exists(icon_path):
    icon = tk.PhotoImage(file=icon_path)
    window.iconphoto(True, icon)


# Customize the window background color
window.configure(bg='#f4f4f4')

# Create and place GUI elements with styling
tk.Label(window, text="Sender Email:", bg='#f4f4f4', font=("Arial", 10)).grid(row=0, column=0, padx=10, pady=5, sticky="e")
entry_sender_email = tk.Entry(window, width=40, font=("Arial", 10))
entry_sender_email.grid(row=0, column=1, padx=10, pady=5)

tk.Label(window, text="Sender Password:", bg='#f4f4f4', font=("Arial", 10)).grid(row=1, column=0, padx=10, pady=5, sticky="e")
entry_sender_pass = tk.Entry(window, width=40, show="*", font=("Arial", 10))
entry_sender_pass.grid(row=1, column=1, padx=10, pady=5)

tk.Label(window, text="Send Date (YYYY-MM-DD):", bg='#f4f4f4', font=("Arial", 10)).grid(row=2, column=0, padx=10, pady=5, sticky="e")
entry_send_date = tk.Entry(window, width=40, font=("Arial", 10))
entry_send_date.grid(row=2, column=1, padx=10, pady=5)

tk.Label(window, text="Send Time (HH:MM):", bg='#f4f4f4', font=("Arial", 10)).grid(row=3, column=0, padx=10, pady=5, sticky="e")
entry_send_time = tk.Entry(window, width=40, font=("Arial", 10))
entry_send_time.grid(row=3, column=1, padx=10, pady=5)

tk.Label(window, text="Mode (one/many):", bg='#f4f4f4', font=("Arial", 10)).grid(row=4, column=0, padx=10, pady=5, sticky="e")
var_mode = tk.StringVar(value="one")
tk.Radiobutton(window, text="One", variable=var_mode, value="one", bg='#f4f4f4', font=("Arial", 10)).grid(row=4, column=1, padx=10, pady=5, sticky="w")
tk.Radiobutton(window, text="Many", variable=var_mode, value="many", bg='#f4f4f4', font=("Arial", 10)).grid(row=4, column=1, padx=10, pady=5, sticky="e")

# For the 'one' mode (single email) fields
tk.Label(window, text="Receiver Email:", bg='#f4f4f4', font=("Arial", 10)).grid(row=5, column=0, padx=10, pady=5, sticky="e")
entry_receiver_email = tk.Entry(window, width=40, font=("Arial", 10))
entry_receiver_email.grid(row=5, column=1, padx=10, pady=5)

tk.Label(window, text="Receiver Name:", bg='#f4f4f4', font=("Arial", 10)).grid(row=6, column=0, padx=10, pady=5, sticky="e")
entry_receiver_name = tk.Entry(window, width=40, font=("Arial", 10))
entry_receiver_name.grid(row=6, column=1, padx=10, pady=5)

tk.Label(window, text="Subject:", bg='#f4f4f4', font=("Arial", 10)).grid(row=7, column=0, padx=10, pady=5, sticky="e")
entry_subject = tk.Entry(window, width=40, font=("Arial", 10))
entry_subject.grid(row=7, column=1, padx=10, pady=5)

tk.Label(window, text="Attachment Path:", bg='#f4f4f4', font=("Arial", 10)).grid(row=8, column=0, padx=10, pady=5, sticky="e")
entry_attachment_path = tk.Entry(window, width=40, font=("Arial", 10))
entry_attachment_path.grid(row=8, column=1, padx=10, pady=5)

tk.Button(window,width=10,text="Browse", command=browse_attachment, font=("Arial", 10), bg="#4CAF50", fg="white").grid(row=8, column=2, padx=10, pady=5)

# For the 'many' mode (bulk email) field
tk.Label(window, text="Excel File:", bg='#f4f4f4', font=("Arial", 10)).grid(row=9, column=0, padx=10, pady=5, sticky="e")
entry_excel_file = tk.Entry(window, width=40, font=("Arial", 10))
entry_excel_file.grid(row=9, column=1, padx=10, pady=5)

tk.Button(window,width=10, text="Browse", command=browse_excel, font=("Arial", 10), bg="#4CAF50", fg="white").grid(row=9, column=2, padx=10, pady=5)

# Common button styling
button_style = {
    "font": ("Arial", 12, "bold"),
    "bg": "#FF5733",  # Same background color as Send Emails button
    "fg": "white"
}

# Send button with consistent styling and center alignment
tk.Button(window, text="Send Emails", command=gui_send_emails, **button_style).grid(row=10, column=1, padx=10, pady=20)

# Exit button with the same style and right alignment
tk.Button(window, width=10, text="Exit", command=window.quit, **button_style).grid(row=10, column=2, padx=10, pady=20, sticky="e")


# Start the GUI loop
window.mainloop()


