
# ================================================================================

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

# SMTP server configuration
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 465

def schedule_email(send_time_input):
    try:
        send_time = datetime.strptime(send_time_input, "%H:%M:%S").time()
    except ValueError:
        try:
            send_time = datetime.strptime(send_time_input, "%H:%M").time()
        except ValueError:
            raise ValueError("Invalid time format. Please use HH:MM or HH:MM:SS.")
    while datetime.now().time() < send_time:
        time.sleep(1)
    print("It's time to send the emails!")


def send_email(sender_email, sender_pass, receiver_email, name, attachment_path):
    try:
        msg = MIMEMultipart("alternative")
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = "Test Mail"

        text = f"Hello {name},\n\nI have something for you. Please find the attachment."
        html = f"""
        <html>
            <body>
                <p>Hello {name},</p>
                <p>I have something for you. Please find the attachment.</p>
                <p>Best regards,<br>Your Name</p>
            </body>
        </html>
        """
        msg.attach(MIMEText(text, "plain"))
        msg.attach(MIMEText(html, "html"))

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

        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.login(sender_email, sender_pass)
            smtp.send_message(msg)

        print(f"Email sent successfully to {receiver_email}.")
    except Exception as e:
        print(f"Error sending email to {receiver_email}: {e}")


def send_bulk_emails(sender_email, sender_pass, df, base_directory):
    for _, row in df.iterrows():
        receiver_email = row['EMAIL_ID']
        name = row['NAME']
        attachment = row['Files to be attached']
        attachment_path = os.path.join(base_directory, attachment)

        if not os.path.exists(attachment_path):
            print(f"Attachment NOT found: {attachment_path}")
            continue

        send_email(sender_email, sender_pass, receiver_email, name, attachment_path)


def load_excel_file(file_path):
    try:
        df = pd.read_excel(file_path)
        required_columns = {"EMAIL_ID", "Files to be attached", "NAME"}
        if not required_columns.issubset(df.columns):
            raise ValueError(f"Excel file must contain columns: {', '.join(required_columns)}")
        return df
    except Exception as e:
        raise ValueError(f"Error reading Excel file: {e}")


def start_sending_emails(mode, sender_email, sender_pass, receiver_email=None, name=None, attachment_path=None, excel_file=None, send_datetime_input=None, email_list=None):
    try:
        # Call the schedule_email function to wait until the specified time
        schedule_email(send_datetime_input)
        
        # Send the email (either one email, many emails from the excel sheet, or all emails from the list)
        if mode == "one":
            send_email(sender_email, sender_pass, receiver_email, name, attachment_path)
        elif mode == "many":
            base_directory = os.path.dirname(excel_file)
            df = load_excel_file(excel_file)
            send_bulk_emails(sender_email, sender_pass, df, base_directory)
        elif mode == "all":
            # If we have an email list to send to all
            if email_list:
                send_to_all(sender_email, sender_pass, email_list, attachment_path)
            else:
                raise ValueError("Email list must be provided for 'all' mode.")

        messagebox.showinfo("Success", "Emails sent successfully!")
    except Exception as e:
        messagebox.showerror("Error", str(e))


def main():
    def browse_file(entry_widget):
        file_path = filedialog.askopenfilename()
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)

    def submit():
        mode = mode_var.get()
        sender_email = email_entry.get()
        sender_pass = password_entry.get()
        send_time_input = time_entry.get()

        if mode == "one":
            receiver_email = recipient_email_entry.get()
            name = recipient_name_entry.get()
            attachment_path = attachment_entry.get()
            Thread(target=start_sending_emails, args=(mode, sender_email, sender_pass, receiver_email, name, attachment_path, None, send_time_input)).start()

        elif mode == "all":
            # For "all" mode, we can either use the Excel file or a manual list
            if manual_list_entry.get():
                # Parse the manual list input (Email, Name format)
                emails = manual_list_entry.get().split(",")
                for email in emails:
                    name = email.split(":")[0]
                    email_list.append((email.split(":")[1], name))
                attachment_path = attachment_entry.get()
                Thread(target=start_sending_emails, args=(mode, sender_email, sender_pass, None, None, attachment_path, None, send_datetime_input, email_list)).start()

    root = tk.Tk()
    root.title("Email Scheduler")
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
    tk.Radiobutton(root, text="All", variable=mode_var, value="all").grid(row=2, column=3, sticky=tk.W)


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

    tk.Label(root, text="Send Time (HH:MM:SS):").grid(row=7, column=0, sticky=tk.W, padx=5, pady=5)
    time_entry = tk.Entry(root, width=40)
    time_entry.grid(row=7, column=1, padx=5, pady=5)

    tk.Button(root, text="Submit", command=submit).grid(row=8, column=0, columnspan=3, pady=10)

    root.mainloop()


if __name__ == "__main__":
    main()
