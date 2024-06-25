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
    message['to'] = recipient
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
    print(f'Sent message to {recipient}. Message Id: {sent_message["id"]}')

def main():
    # Attach the message body
    authenticate()
    service = build('gmail', 'v1', credentials=creds)

    df = pd.read_excel("attendance.xlsx")

    recipients = ['00mrxgames00@gmail.com']
    
    for indx, recipient in df.iterrows():
    # for indx, recipient in enumerate(recipients):
        try:
            num = indx + 1
            attachment_path = f"certificates/{num}_PDFsam_شهادات برنامج أخصائي الامن السيبراني.pdf"  # Update with the path to your attachment
            print("the index: ", num)
            print("The email: ", recipient['ar_name'])
            send_email(service, recipient["email"], attachment_path)
            # send_email(service, recipient, attachment_path)
        except HTTPError as error:
            print(f'An error occurred while sending email to {recipient}: {error}')

if __name__ == '__main__':
    main()