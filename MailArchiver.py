import base64
import os
import json
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

def read_json(json_path):
    """ Read a json file.
    Parameters:
        json_path: String. Path of the json file.
    Returns:
        d: dict. Content of the file.
    """
    with open(json_path, 'r') as f:
        d = json.load(f)
    return d

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def authenticate_gmail():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json')
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def get_messages_with_attachments(service, user_id='me', query='has:attachment'):
    try:
        response = service.users().messages().list(userId=user_id, q=query).execute()
        messages = response.get('messages', [])
        return messages
    except Exception as e:
        print(f"An error occurred: {e}")
        return []
    
def _get_message_details(service, user_id='me', msg_id=''):
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id).execute()

        # Extract subject
        subject = next(header['value'] for header in message['payload']['headers'] if header['name'] == 'Subject')

        # Extract body (assuming it's in the 'text/plain' format)
        body = base64.urlsafe_b64decode(message['payload']['parts'][0]['body']['data']).decode('utf-8')

        # Extract attachment names
        attachment_names = [part['filename'] for part in message['payload']['parts'] if 'filename' in part]

        return subject, body, attachment_names

    except Exception as e:
        print(f"An error occurred: {e}")
        return None, None, None


def download_attachment(service, msg_id, user_id='me', output_dir='attachments/'):
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id).execute()
        
        parts = message['payload']['parts']
        for part in parts:
            if 'filename' in part and 'mimeType' in part:
                if part['filename'] != '' and part['mimeType'] == 'application/pdf':
                    filename = part['filename']
                    attachment_id = part['body']['attachmentId']
                    attachment = service.users().messages().attachments().get(userId=user_id, messageId=msg_id, id=attachment_id).execute()
                    file_data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))
                    output_path = os.path.join(output_dir, filename)
                    with open(output_path, 'wb') as f:
                        f.write(file_data)
                    print(f"Attachment '{filename}' saved to {output_path}")
    except Exception as e:
        print(f"An error occurred: {e}")

def main():
    creds = authenticate_gmail()
    service = build('gmail', 'v1', credentials=creds)

    # Specify your search criteria (e.g., subject, sender, etc.)
    search_query = 'subject:book-n-drive-Rechnung has:attachment'
    messages = get_messages_with_attachments(service, query=search_query)

    for message in messages:
        msg_id = message['id']
        #subject, body, attachment_names = _get_message_details(service, msg_id=msg_id)
        #print("subject={}, attachment_names={}".format(subject, attachment_names))
        download_attachment(service, msg_id)

if __name__ == '__main__':
    main()
