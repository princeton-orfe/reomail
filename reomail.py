import argparse
import webbrowser
from O365 import Account
from decouple import config
import os.path
import sys
import csv

client_id = config('CLIENT_ID')
client_secret = config('CLIENT_SECRET')
tenant_id = config('TENANT_ID')

scopes = ['https://graph.microsoft.com/Mail.Send']

# Create argument parser
parser = argparse.ArgumentParser(description='Send an HTML email using O365 with optional recipients from CSV files.')
parser.add_argument('--tofile', help='CSV file containing "to" email addresses')
parser.add_argument('--bccfile', help='CSV file containing "bcc" email addresses')
parser.add_argument('--subject', help='Subject of the email')
parser.add_argument('--bodyfile', default='body.html', help='HTML file containing the email body content, default is body.html')
parser.add_argument('--merge', action='store_true', help="Uses the CSV tofile or address list to send one message per recipient.  Causes bcc addresses or bccfile to be ignored.")
args = parser.parse_args()

try:
    with open(args.bodyfile) as f:
        msg_body_content = f.read()
except FileNotFoundError:
    print(f'Provide a file named body.html or specify a custom file with the --bodyfile option.')
    sys.exit(1)

# Prompt for the subject with a warning if blank
subject = args.subject or input('Enter the subject of the email: ')

if not subject:
    confirm = input('The subject is blank. Do you want to continue? (yes/no): ')
    if confirm.lower() != 'yes':
        print('Email not sent. Aborted.')
        sys.exit(1)

# Function to read emails from a CSV file
def read_emails_from_csv(file_path):
    emails = []
    with open(file_path, 'r') as csv_file:
        reader = csv.reader(csv_file)
        for row in reader:
            if row:
                emails.append(row[0].strip())
    return emails

# Parse CSV files for "to" and "bcc" recipients if provided
to_emails = []
bcc_emails = []

if args.tofile:
    to_emails = read_emails_from_csv(args.tofile)
if args.bccfile:
    bcc_emails = read_emails_from_csv(args.bccfile)



# Prompt the user for recipients if not provided in CSV

if not to_emails:
    to_input = input('Enter "to" email addresses (comma-separated, max 300):')
    to_emails = [email.strip() for email in to_input.split(',')]

if not bcc_emails:
    bcc_input = input('Enter "bcc" email addresses (comma-separated, max 300):')
    bcc_emails = [email.strip() for email in bcc_input.split(',')]

if len(to_emails) + len(bcc_emails) > 300:
    print('Total recipients (to + bcc) exceeds 300. Please reduce the number of recipients.')
    sys.exit(1)

# Authenticate the account - client cred flow?
account = Account((client_id, client_secret), auth_flow_type='authorization', tenant_id=tenant_id, scopes=scopes)
if not account.is_authenticated:
    if not account.authenticate():
        print('Authentication failed.')
        sys.exit(1)
print('Authenticated')

# if not account.is_authenticated:
#     consent_url, _ = account.con.get_authorization_url(redirect_uri='http://localhost:8000/oauth/callback')
#     print(f'Visit the following URL to give consent to send mail: {consent_url}')
#     webbrowser.open(consent_url)
#     callback_url = input('Paste the authenticated URL here: ')
#     if not account.is_authenticated:
#         if not account.authenticate(callback=callback_url):
#             print('Authentication failed.')
#             sys.exit(1)

# Create a new message
m = account.new_message()
m.to.add(to_emails)
m.bcc.add(bcc_emails)
m.subject = subject
m.body = msg_body_content

attachments = ['banner1.jpg', 'banner2.jpg', 'banner3.jpg', 'pu-logo.png', 'orfe.png']

# Add attachments with inline properties and content IDs
for attachment_name in attachments:
    m.attachments.add(attachment_name)
    attachment = m.attachments[-1] # Get the last added attachment
    attachment.is_inline = True
    attachment.content_id = attachment_name

# Send the message
if not m.send():
    print('Message sending failed.')
else:
    print('Message sent successfully.')
