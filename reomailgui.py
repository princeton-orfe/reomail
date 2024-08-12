import tkinter as tk
from tkinter import filedialog, messagebox
#from O365 import Account
from azure.identity import InteractiveBrowserCredential
from msal import ConfidentialClientApplication
from decouple import config
import requests
import csv

# Authentication details
client_id = config('CLIENT_ID')
tenant_id = config('TENANT_ID')

authority = f"https://login.microsoftonline.com/{tenant_id}"
scopes = ['https://graph.microsoft.com/Mail.Send']

# def authenticate_account():
#     global account
#     account = Account((client_id, client_secret), auth_flow_type='authorization', tenant_id=tenant_id, scopes=scopes)
#     if not account.is_authenticated:
#         if not account.authenticate():
#             messagebox.showerror("Authentication Error", "Failed to authenticate.")
#             return False
#     return True

def authenticate_account():
    global token
    client_id = config('CLIENT_ID')
    tenant_id = config('TENANT_ID')

    scopes = ['https://graph.microsoft.com/.default']

    credential = InteractiveBrowserCredential(client_id=client_id)
    token = credential.get_token(*scopes)
    
    if not token:
        messagebox.showerror("Authentication Error", "Failed to authenticate.")
        return False
    
    return True

# def send_email():
#     subject = subject_entry.get()
#     body_file = body_file_path.get()
#     
#     try:
#         with open(body_file, 'r') as f:
#             msg_body_content = f.read()
#     except FileNotFoundError:
#         messagebox.showerror("File Error", "Body file not found.")
#         return
#     
#     to_emails, _, _ = read_emails_from_csv(to_file_path.get()) if to_file_path.get() else [], [], []
#     bcc_emails, _, _ = read_emails_from_csv(bcc_file_path.get()) if bcc_file_path.get() else [], [], []
#     
#     if not to_emails:
#         to_emails = [email.strip() for email in to_entry.get().split(',')]
#     
#     if len(to_emails) + len(bcc_emails) > 300:
#         messagebox.showerror("Recipient Limit Exceeded", "Total recipients exceed 300.")
#         return
#     
#     if not authenticate_account():
#         return
# 
#     m = account.new_message()
#     m.to.add(to_emails)
#     m.bcc.add(bcc_emails)
#     m.subject = subject
#     m.body = msg_body_content
# 
#     attachments = ['banner1.jpg', 'banner2.jpg', 'banner3.jpg', 'pu-logo.png', 'orfe.png']
#     
#     for attachment_name in attachments:
#         m.attachments.add(attachment_name)
#         attachment = m.attachments[-1]
#         attachment.is_inline = True
#         attachment.content_id = attachment_name
# 
#     if not m.send():
#         messagebox.showerror("Send Error", "Message sending failed.")
#     else:
#         messagebox.showinfo("Success", "Message sent successfully.")
# 

def send_email():
    subject = subject_entry.get()
    body_file = body_file_path.get()

    try:
        with open(body_file, 'r') as f:
            msg_body_content = f.read()
    except FileNotFoundError:
        messagebox.showerror("File Error", "Body file not found.")
        return

    to_emails, _, _ = read_emails_from_csv(to_file_path.get()) if to_file_path.get() else [], [], []
    bcc_emails, _, _ = read_emails_from_csv(bcc_file_path.get()) if bcc_file_path.get() else [], [], []

    if not to_emails:
        to_emails = [email.strip() for email in to_entry.get().split(',')]

    if len(to_emails) + len(bcc_emails) > 300:
        messagebox.showerror("Recipient Limit Exceeded", "Total recipients exceed 300.")
        return

    if not authenticate_account():
        return

    headers = {
        'Authorization': f'Bearer {token.token}',
        'Content-Type': 'application/json'
    }

    email_data = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": msg_body_content
            },
            "toRecipients": [{"emailAddress": {"address": email}} for email in to_emails],
            "bccRecipients": [{"emailAddress": {"address": email}} for email in bcc_emails],
        }
    }

#    attachments = ['banner1.jpg', 'banner2.jpg', 'banner3.jpg', 'pu-logo.png', 'orfe.png']
    attachments = []
    email_data['message']['attachments'] = []

    for attachment_name in attachments:
        with open(attachment_name, "rb") as attachment_file:
            encoded_content = base64.b64encode(attachment_file.read()).decode('utf-8')
            email_data['message']['attachments'].append({
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": attachment_name,
                "contentBytes": encoded_content,
                "isInline": True,
                "contentId": attachment_name
            })

    response = requests.post(
        'https://graph.microsoft.com/v1.0/me/sendMail',
        headers=headers,
        json=email_data
    )

    if not response.ok:
        messagebox.showerror("Send Error", "Message sending failed.")
    else:
        messagebox.showinfo("Success", "Message sent successfully.")



def browse_file_path(entry):
    filename = filedialog.askopenfilename()
    entry.set(filename)

def read_emails_from_csv(file_path):
    emails = []
    first_names = []
    last_names = []
    
    with open(file_path, 'r') as csv_file:
        reader = csv.DictReader(csv_file)
        if 'email' not in reader.fieldnames:
            raise Exception("The CSV file requires a header row with at least one column specifying 'email'.")
        
        for row in reader:
            email = row.get('email', '').strip()
            first_name = row.get('first_name', '').strip()
            last_name = row.get('last_name', '').strip()
            
            emails.append(email)
            first_names.append(first_name)
            last_names.append(last_name)
    
    return emails, first_names, last_names

# GUI setup
root = tk.Tk()
root.title("HTML Email Sender")

tk.Label(root, text="Subject:").grid(row=0, column=0)
subject_entry = tk.Entry(root, width=50)
subject_entry.grid(row=0, column=1)

tk.Label(root, text="Body HTML File:").grid(row=1, column=0)
body_file_path = tk.StringVar()
tk.Entry(root, textvariable=body_file_path, width=50).grid(row=1, column=1)
tk.Button(root, text="Browse", command=lambda: browse_file_path(body_file_path)).grid(row=1, column=2)

tk.Label(root, text="To Email File (CSV):").grid(row=2, column=0)
to_file_path = tk.StringVar()
tk.Entry(root, textvariable=to_file_path, width=50).grid(row=2, column=1)
tk.Button(root, text="Browse", command=lambda: browse_file_path(to_file_path)).grid(row=2, column=2)

tk.Label(root, text="BCC Email File (CSV):").grid(row=3, column=0)
bcc_file_path = tk.StringVar()
tk.Entry(root, textvariable=bcc_file_path, width=50).grid(row=3, column=1)
tk.Button(root, text="Browse", command=lambda: browse_file_path(bcc_file_path)).grid(row=3, column=2)

tk.Label(root, text="To Emails (Comma Separated):").grid(row=4, column=0)
to_entry = tk.Entry(root, width=50)
to_entry.grid(row=4, column=1)

tk.Button(root, text="Send Email", command=send_email).grid(row=5, column=1)

root.mainloop()
