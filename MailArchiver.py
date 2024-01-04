import base64
import os
import json
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

# CONSTANTS
## If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']


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


def authenticate_gmail():
    """ Perform the Gmail authentication. """
    creds = None
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


def download_attachment(service, msg_id, output_dir, user_id='me'):
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id).execute()
        
        parts = message['payload']['parts']
        for part in parts:
            if 'filename' in part and 'mimeType' in part:
                if part['filename'] != '' and part['mimeType'] == 'application/pdf':
                    filename = part['filename']
                    if not os.path.exists(os.path.join(output_dir, filename)):
                        attachment_id = part['body']['attachmentId']
                        attachment = service.users().messages().attachments().get(userId=user_id, messageId=msg_id, id=attachment_id).execute()
                        file_data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))
                        output_path = os.path.join(output_dir, filename)
                        with open(output_path, 'wb') as f:
                            f.write(file_data)
                        print(f"Attachment '{filename}' saved to {output_path}")
    except Exception as e:
        print(f"An error occurred: {e}")


def search_and_get_attachment(service, subject, destination_folder):
    """ Search for the desired messages and save the attachment file in the destination folder. 
    Paramters:
        service: Google auth service.
        subject: String.
        destination_folder: String. Path of the folder.
    """
    print("Searching for mesages: {}".format(subject))
    # Specify your search criteria (e.g., subject, sender, etc.)
    search_query = 'subject:{} has:attachment'.format(subject)
    messages = get_messages_with_attachments(service, query=search_query)
    print("\tFound {} messages...".format(len(messages)))

    for message in messages:
        msg_id = message['id']
        download_attachment(service, msg_id, destination_folder)


if __name__ == '__main__':
    # Load the configuration file
    settings = read_json('config.json')

    # Login in Gmail
    creds = authenticate_gmail()
    service = build('gmail', 'v1', credentials=creds)

    # perform the seach
    for i in list(settings.keys()):
        subject = settings[i]['subject']
        destination_folder = settings[i]['destination_folder']

        os.makedirs(destination_folder, exist_ok=True)
        # TODO: add time search, e.g. -w 3 in the last 3 weeks, or from the last time the script was run
        search_and_get_attachment(service, subject, destination_folder)
