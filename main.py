
# ================================================================================

# Importing required libraries
import os
import pandas as pd  # For handling Excel files
import smtplib  # For sending emails
from email.mime.text import MIMEText  # For creating email content
from email.mime.multipart import MIMEMultipart  # For creating multipart emails
from email.mime.base import MIMEBase  # For handling attachments
from email import encoders  # For encoding attachments
from datetime import datetime  # For working with date and time
import time  # For handling delays
import tkinter as tk  # For GUI creation
from tkinter import filedialog, messagebox  # For file dialog and message boxes
from threading import Thread  # For handling multithreading

# SMTP server configuration for Gmail
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 465

# Function to schedule email sending at a specific time

def schedule_email(send_time_input):
    """
    Schedule email to be sent at a specific date and time.
    :param send_time_input: Scheduled date-time in the format 'YYYY-MM-DD HH:MM' or 'YYYY-MM-DD HH:MM:SS'
    """
    try:
        send_time = datetime.strptime(send_time_input, "%Y-%m-%d %H:%M:%S")
    except ValueError:
        try:
            send_time = datetime.strptime(send_time_input, "%Y-%m-%d %H:%M")
        except ValueError:
            raise ValueError("Invalid datetime format. Please use 'YYYY-MM-DD HH:MM' or 'YYYY-MM-DD HH:MM:SS'.")
    
    current_time = datetime.now()
    
    # Ensure the scheduled time is in the future
    if send_time <= current_time:
        raise ValueError("Scheduled time must be in the future.")
    
    # Wait until the scheduled time
    while datetime.now() < send_time:
        time.sleep(1)
    
    print("It's time to send the emails!")

# Function to send a single email
def send_email(sender_email, sender_pass, receiver_email, name, attachment_path):
    try:
        # Create a multipart email
        msg = MIMEMultipart("alternative")
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = "Test Mail"

        # Define email content (plain text and HTML)
        text = f"Hello {name},\n\nI have something for you. Please find the attachment."
        html = f"""
        <html>
            <body>
                <p>Hello {name},</p>
                <p>I have something for you. Please find the attachment.</p>
                <p>Best regards,<br>Leader Din</p>
            </body>
        </html>
        """
        # Attach plain text and HTML content
        msg.attach(MIMEText(text, "plain"))
        msg.attach(MIMEText(html, "html"))

        # Attach a file if the path is valid
        if os.path.exists(attachment_path):
            with open(attachment_path, 'rb') as f:
                file_part = MIMEBase('application', 'octet-stream')
                file_part.set_payload(f.read())
                encoders.encode_base64(file_part)
                file_part.add_header(
                    'Content-Disposition',
                    f'attachment; filename="{os.path.basename(attachment_path)}"'
                )
                msg.attach(file_part)
        else:
            print(f"Attachment NOT found: {attachment_path}")
            return

        # Connect to SMTP server and send the email
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.login(sender_email, sender_pass)
            smtp.send_message(msg)

        print(f"Email sent successfully to {receiver_email}.")
    except Exception as e:
        print(f"Error sending email to {receiver_email}: {e}")

# Function to send bulk emails using an Excel file
def send_bulk_emails(sender_email, sender_pass, df, base_directory):
    for _, row in df.iterrows():
        receiver_email = row['EMAIL_ID']
        name = row['NAME']
        attachment = row['Files to be attached']
        attachment_path = os.path.join(base_directory, attachment)

        # Check if the attachment exists
        if not os.path.exists(attachment_path):
            print(f"Attachment NOT found: {attachment_path}")
            continue

        # Send email for each row
        send_email(sender_email, sender_pass, receiver_email, name, attachment_path)




# Function to load and validate an Excel file
def load_excel_file(file_path):
    try:
        df = pd.read_excel(file_path)
        required_columns = {"EMAIL_ID", "Files to be attached", "NAME"}
        # Check if required columns are present
        if not required_columns.issubset(df.columns):
            raise ValueError(f"Excel file must contain columns: {', '.join(required_columns)}")
        return df
    except Exception as e:
        raise ValueError(f"Error reading Excel file: {e}")

# Function to handle email sending based on mode (one/many)
def start_sending_emails(mode, sender_email, sender_pass, receiver_email=None, name=None, attachment_path=None, excel_file=None, send_time_input=None):
    try:
        # Schedule email sending
        schedule_email(send_time_input)
        if mode == "one":
            # Send a single email
            send_email(sender_email, sender_pass, receiver_email, name, attachment_path)
        elif mode == "many":
            # Send bulk emails
            base_directory = os.path.dirname(excel_file)
            df = load_excel_file(excel_file)
            send_bulk_emails(sender_email, sender_pass, df, base_directory)
        messagebox.showinfo("Success", "Emails sent successfully!")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# GUI creation for the email automation tool
def main():
    def browse_file(entry_widget):
        # Open a file dialog to select a file
        file_path = filedialog.askopenfilename()
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)

    def submit():
        # Collect input values from the GUI
        mode = mode_var.get()
        sender_email = email_entry.get()
        sender_pass = password_entry.get()
        send_time_input = time_entry.get()

        if mode == "one":
            receiver_email = recipient_email_entry.get()
            name = recipient_name_entry.get()
            attachment_path = attachment_entry.get()
            # Start email sending in a separate thread
            Thread(target=start_sending_emails, args=(mode, sender_email, sender_pass, receiver_email, name, attachment_path, None, send_time_input)).start()

        elif mode == "many":
            excel_file = excel_entry.get()
            # Start bulk email sending in a separate thread
            Thread(target=start_sending_emails, args=(mode, sender_email, sender_pass, None, None, None, excel_file, send_time_input)).start()

    # GUI window setup
    root = tk.Tk()
    root.title("Email Automation Sending")
    root.geometry("600x350")
    root.resizable(False, False)

    # GUI elements for user inputs
    tk.Label(root, text="Your Email:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
    email_entry = tk.Entry(root, width=40)
    email_entry.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(root, text="Your Password:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
    password_entry = tk.Entry(root, width=40, show="*")
    password_entry.grid(row=1, column=1, padx=5, pady=5)

    tk.Label(root, text="Mode:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
    mode_var = tk.StringVar(value="one")
    tk.Radiobutton(root, text="One", variable=mode_var, value="one").grid(row=2, column=1, sticky=tk.W)
    tk.Radiobutton(root, text="Many", variable=mode_var, value="many").grid(row=2, column=2, sticky=tk.W)

    tk.Label(root, text="Recipient Email:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
    recipient_email_entry = tk.Entry(root, width=40)
    recipient_email_entry.grid(row=3, column=1, padx=5, pady=5)

    tk.Label(root, text="Recipient Name:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
    recipient_name_entry = tk.Entry(root, width=40)
    recipient_name_entry.grid(row=4, column=1, padx=5, pady=5)




    tk.Label(root, text="Attachment Path:").grid(row=5, column=0, sticky=tk.W, padx=5, pady=5)
    attachment_entry = tk.Entry(root, width=40)
    attachment_entry.grid(row=5, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=lambda: browse_file(attachment_entry)).grid(row=5, column=2, padx=5, pady=5)

    tk.Label(root, text="Excel File:").grid(row=6, column=0, sticky=tk.W, padx=5, pady=5)
    excel_entry = tk.Entry(root, width=40)
    excel_entry.grid(row=6, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=lambda: browse_file(excel_entry)).grid(row=6, column=2, padx=5, pady=5)

    tk.Label(root, text="Send Time (YYYY-MM-DD HH:MM:SS):").grid(row=7, column=0, sticky=tk.W, padx=5, pady=5)
    time_entry = tk.Entry(root, width=40)
    time_entry.grid(row=7, column=1, padx=5, pady=5)

    tk.Button(root, text="Submit", command=submit).grid(row=8, column=0, columnspan=3, pady=10)

    root.mainloop()

# Run the GUI application
if __name__ == "__main__":
    main()