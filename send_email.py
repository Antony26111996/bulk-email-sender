import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from tkinter import messagebox, Tk, Label, Button, Text, END, Entry, filedialog
import os  # Import os to get file sizes

# Function to send the email
def send_bulk_email(subject, message, recipient_emails, sender_email, sender_password, attachments):
    try:
        # Set up the SMTP server
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()  # Secure the connection
        server.login(sender_email, sender_password)

        for recipient in recipient_emails:
            # Create a MIMEMultipart message
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient
            msg['Subject'] = subject

            # Add the email body
            msg.attach(MIMEText(message, 'plain'))

            # Attach files
            for attachment in attachments:
                with open(attachment, 'rb') as file:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(file.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename={attachment.split("/")[-1]}')
                    msg.attach(part)

            # Send the email
            server.send_message(msg)

        server.quit()  # Terminate the SMTP session
        messagebox.showinfo("Success", "Emails sent successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Function to get the form data and send the emails
def send_emails_from_ui():
    subject = subject_entry.get()
    message = message_box.get("1.0", END)
    recipients = recipient_entry.get().split(',')  # Comma-separated emails
    sender_email = sender_email_entry.get()
    sender_password = sender_password_entry.get()

    if subject and message and recipients and sender_email and sender_password:
        send_bulk_email(subject, message, recipients, sender_email, sender_password, attachments)
    else:
        messagebox.showerror("Error", "All fields must be filled out!")

# Function to select attachments
def attach_files():
    global attachments
    files = filedialog.askopenfilenames(title="Select Files")
    attachments.extend(files)
    
    # Update the attachment label with file names and sizes
    attachment_info = "\n".join([f"{os.path.basename(file)} ({os.path.getsize(file) / 1024:.2f} KB)" for file in attachments])
    attachment_label.config(text=f"Attachments:\n{attachment_info}")

# Create the UI
def create_ui():
    global attachments, attachment_label
    attachments = []  # List to hold attachment file paths

    window = Tk()
    window.title("Bulk Email Sender")

    # Labels
    Label(window, text="Sender Email:").grid(row=0, column=0)
    Label(window, text="Sender Password:").grid(row=1, column=0)
    Label(window, text="Recipients (comma separated):").grid(row=2, column=0)
    Label(window, text="Subject:").grid(row=3, column=0)
    Label(window, text="Message:").grid(row=4, column=0)

    # Entry widgets
    global sender_email_entry, sender_password_entry, recipient_entry, subject_entry, message_box
    sender_email_entry = Entry(window, width=50)
    sender_password_entry = Entry(window, show="*", width=50)
    recipient_entry = Entry(window, width=50)
    subject_entry = Entry(window, width=50)
    message_box = Text(window, height=10, width=50)

    # Place the entry widgets
    sender_email_entry.grid(row=0, column=1)
    sender_password_entry.grid(row=1, column=1)
    recipient_entry.grid(row=2, column=1)
    subject_entry.grid(row=3, column=1)
    message_box.grid(row=4, column=1)

    # Attach files button
    attach_button = Button(window, text="Attach Files", command=attach_files)
    attach_button.grid(row=5, column=0)

    # Label to show number of attachments
    attachment_label = Label(window, text="Attachments: 0 files selected.")
    attachment_label.grid(row=5, column=1)

    # Send button
    send_button = Button(window, text="Send Emails", command=send_emails_from_ui)
    send_button.grid(row=6, column=1)

    window.mainloop()

if __name__ == "__main__":
    create_ui()
