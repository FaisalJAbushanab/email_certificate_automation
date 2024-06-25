import base64
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import pandas as pd
from requests import HTTPError

SCOPES = ['https://www.googleapis.com/auth/gmail.send']
creds = None

def authenticate():
    global creds
    flow = InstalledAppFlow.from_client_secrets_file('cred.json', SCOPES)
    creds = flow.run_local_server(port=0)

def send_email(service, recipient, attachment_path=None):
    # Create a multipart message
    message = MIMEMultipart()
    message['to'] = recipient # Email recipient
    # Email subject
    message['subject'] = 'شهادة اتمام حضور برنامج أخصائي الأمن السيبراني - مركز المبدعون'
    
    # Attach the message body
    msg = ''

    # Open the file in read mode
    with open('msg.html', 'r', encoding='utf-8') as file:
        # Read the contents of the file
        msg = file.read()
        # print(text)

    message.attach(MIMEText(msg, 'html'))

    # Attach the file if provided
    if attachment_path:
        attachment_name = os.path.basename(attachment_path)
        with open(attachment_path, 'rb') as attachment_file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment_file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="soce_cert.pdf"')
        message.attach(part)

    # Convert message to string
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

    # Send the email
    sent_message = service.users().messages().send(userId="me", body={'raw': raw_message}).execute()
    # Logging delivery status
    print(f'Sent message to {recipient}. Message Id: {sent_message["id"]}')

def main():
    # Authenticate request first
    authenticate()
    service = build('gmail', 'v1', credentials=creds)

    # Read attendance sheet to extract emails
    df = pd.read_excel("attendance.xlsx")
    
    # Iterate through attendance rows
    for indx, recipient in df.iterrows():
        try:
            num = indx + 1 # Start from index 1
            
            # define attachment path
            # Update with the path to your attachment
            attachment_path = f"certificates/{num}_PDFsam_شهادات برنامج أخصائي الامن السيبراني.pdf"  
            
            # Send email synchronously 
            send_email(service, recipient["email"], attachment_path)
        except HTTPError as error:
            print(f'An error occurred while sending email to {recipient}: {error}')

if __name__ == '__main__':
    main()