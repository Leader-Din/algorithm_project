import os
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tkinter import Tk, Canvas, Frame, Label, Entry, Button, filedialog, OptionMenu, StringVar
from tkinter import messagebox
from tkinter import ttk  # Import ttk for Combobox widget
import time

# Email credentials
SMTP_SERVER = 'smtp.gmail.com'  # Change for different providers
SMTP_PORT = 587

def send_email(server, email_address, recipient_name, recipient_email, due_date, invoice_no, amount, custom_message):
    subject = "Invoice Details"
   
    # Email body with invoice information
    body = f"""
    Dear {recipient_name},\n\n
    I just wanted to drop your price {amount} USD in respect of our invoice {invoice_no} is due 
    for payment on {due_date}. I would be really grateful if you could confirm that everything is on track 
    for payment.\n\n
    {custom_message} \n\n
    Best regards,\n
    Your Name
    """
    
    # Initialize the MIMEMultipart message object
    msg = MIMEMultipart()
    msg['From'] = email_address
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        server.sendmail(email_address, recipient_email, msg.as_string())
        print(f"Email sent to {recipient_name} ({recipient_email})")
    except Exception as e:
        print(f"Failed to send email to {recipient_name}: {e}")

def terminal_mode():
    print("You are in Terminal Mode")
    send_option = input("Do you want to send to an individual or group? (individual/group): ").strip().lower()

    if send_option == "individual":
        recipient_name = input("Enter recipient name: ").strip()
        recipient_email = input("Enter recipient email: ").strip()
        due_date = input("Enter due date: ").strip()
        invoice_no = input("Enter invoice number: ").strip()
        amount = input("Enter amount: ").strip()
        custom_message = input("Enter custom message (optional): ").strip()

        email_address = input("Enter your email address: ").strip()
        email_password = input("Enter your email password: ").strip()

        try:
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            server.starttls()
            server.login(email_address, email_password)
        except Exception as e:
            print(f"Failed to connect to the email server: {e}")
            return

        time.sleep(3)  # Pause before sending the email
        send_email(server, email_address, recipient_name, recipient_email, due_date, invoice_no, amount, custom_message)

        server.quit()
        print(f"Email sent to {recipient_name} ({recipient_email})")

    elif send_option == "group":
        excel_file = input("Enter the path to the Excel file: ").strip()
        if not os.path.exists(excel_file):
            print("Invalid file path. Exiting...")
            return

        try:
            df = pd.read_excel(excel_file)
        except Exception as e:
            print(f"Failed to read the Excel file: {e}")
            return

        email_address = input("Enter your email address: ").strip()
        email_password = input("Enter your email password: ").strip()

        try:
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            server.starttls()
            server.login(email_address, email_password)
        except Exception as e:
            print(f"Failed to connect to the email server: {e}")
            return

        for _, row in df.iterrows():
            recipient_name = row['Name']
            recipient_email = row['Email']
            due_date = row['DueDate']
            invoice_no = row['InvoiceNo']
            amount = row['Amount']
            custom_message = row['Message'] if 'Message' in row else ""

            time.sleep(3)  # Pause before sending the next email
            send_email(server, email_address, recipient_name, recipient_email, due_date, invoice_no, amount, custom_message)

        server.quit()
        print("Emails have been sent successfully.")
    else:
        print("Invalid choice. Exiting...")

def gui_mode():
    def start_email_sending():
        send_option = mode_option.get()

        if send_option == "individual":
            recipient_name = name_entry.get().strip()
            recipient_email = email_entry_individual.get().strip()
            due_date = due_date_entry.get().strip()
            invoice_no = invoice_no_entry.get().strip()
            amount = amount_entry.get().strip()
            custom_message = custom_message_entry.get().strip()

            email_address = email_entry.get().strip()
            email_password = password_entry.get().strip()

            if not email_address or not email_password:
                messagebox.showerror("Error", "Please provide email credentials.")
                return

            try:
                server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
                server.starttls()
                server.login(email_address, email_password)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to connect to the email server: {e}")
                return

            time.sleep(3)  # Pause before sending the email
            send_email(server, email_address, recipient_name, recipient_email, due_date, invoice_no, amount, custom_message)

            server.quit()
            messagebox.showinfo("Success", f"Email sent to {recipient_name} ({recipient_email})")

        elif send_option == "group":
            excel_file = file_entry.get().strip()
            if not excel_file or not os.path.exists(excel_file):
                messagebox.showerror("Error", "Please provide a valid Excel file path.")
                return

            try:
                df = pd.read_excel(excel_file)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read the Excel file: {e}")
                return

            email_address = email_entry.get().strip()
            email_password = password_entry.get().strip()

            if not email_address or not email_password:
                messagebox.showerror("Error", "Please provide email credentials.")
                return

            try:
                server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
                server.starttls()
                server.login(email_address, email_password)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to connect to the email server: {e}")
                return

            for _, row in df.iterrows():
                recipient_name = row['Name']
                recipient_email = row['Email']
                due_date = row['DueDate']
                invoice_no = row['InvoiceNo']
                amount = row['Amount']
                custom_message = row['Message'] if 'Message' in row else ""

                time.sleep(3)  # Pause before sending the next email
                send_email(server, email_address, recipient_name, recipient_email, due_date, invoice_no, amount, custom_message)

            server.quit()
            messagebox.showinfo("Success", "Emails have been sent successfully.")
        else:
            messagebox.showerror("Error", "Invalid choice. Exiting...")

    root = Tk()
    root.title("Email Sender")
    root.geometry("600x400")

    canvas = Canvas(root, height=400, width=600)
    canvas.pack()

    frame = Frame(canvas, bg="white")
    frame.place(relwidth=0.8, relheight=0.8, relx=0.1, rely=0.1)

    # Dropdown for choosing sending mode (Individual or Group)
    Label(frame, text="Select sending mode:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
    mode_option = ttk.Combobox(frame, values=["individual", "group"], state="readonly")
    mode_option.grid(row=0, column=1, padx=10, pady=10)

    # Fields for individual sending
    Label(frame, text="Recipient Name:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
    name_entry = Entry(frame, width=40)
    name_entry.grid(row=1, column=1, padx=10, pady=10)

    Label(frame, text="Recipient Email (individual):").grid(row=2, column=0, padx=10, pady=10, sticky="w")
    email_entry_individual = Entry(frame, width=40)
    email_entry_individual.grid(row=2, column=1, padx=10, pady=10)

    Label(frame, text="Due Date:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
    due_date_entry = Entry(frame, width=40)
    due_date_entry.grid(row=3, column=1, padx=10, pady=10)

    Label(frame, text="Invoice No:").grid(row=4, column=0, padx=10, pady=10, sticky="w")
    invoice_no_entry = Entry(frame, width=40)
    invoice_no_entry.grid(row=4, column=1, padx=10, pady=10)

    Label(frame, text="Amount:").grid(row=5, column=0, padx=10, pady=10, sticky="w")
    amount_entry = Entry(frame, width=40)
    amount_entry.grid(row=5, column=1, padx=10, pady=10)

    Label(frame, text="Custom Message:").grid(row=6, column=0, padx=10, pady=10, sticky="w")
    custom_message_entry = Entry(frame, width=40)
    custom_message_entry.grid(row=6, column=1, padx=10, pady=10)

    # Fields for email credentials
    Label(frame, text="Your Email Address:").grid(row=7, column=0, padx=10, pady=10, sticky="w")
    email_entry = Entry(frame, width=40)
    email_entry.grid(row=7, column=1, padx=10, pady=10)

    Label(frame, text="Your Email Password:").grid(row=8, column=0, padx=10, pady=10, sticky="w")
    password_entry = Entry(frame, width=40, show="*")
    password_entry.grid(row=8, column=1, padx=10, pady=10)

    # Fields for group sending
    Label(frame, text="Excel File Path:").grid(row=9, column=0, padx=10, pady=10, sticky="w")
    file_entry = Entry(frame, width=40)
    file_entry.grid(row=9, column=1, padx=10, pady=10)

    # Button to start sending emails
    send_button = Button(frame, text="Send Emails", command=start_email_sending)
    send_button.grid(row=10, columnspan=2, pady=20)

    root.mainloop()

def main():
    mode = input("Select Mode (terminal/gui): ").strip().lower()
    if mode == "terminal":
        terminal_mode()
    elif mode == "gui":
        gui_mode()
    else:
        print("Invalid choice. Exiting...")

if __name__ == "__main__":
    main()

