import os
import pickle
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

import base64
import email
from email.mime.text import MIMEText

# Define the scopes
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def get_gmail_service():
    """Gets authenticated Gmail API service."""
    creds = None
    
    # Check if token.pickle file exists with stored credentials
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    # If credentials don't exist or are invalid, get new ones
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Load client secrets from downloaded file (rename to 'credentials.json')
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save credentials for future use
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    # Return Gmail API service
    return build('gmail', 'v1', credentials=creds)

def list_messages(service, user_id='me', max_results=10):
    """Lists messages in the user's account."""
    try:
        # Get list of messages
        response = service.users().messages().list(
            userId=user_id, maxResults=max_results).execute()
        messages = response.get('messages', [])
        
        if not messages:
            print('No messages found.')
            return []
            
        print(f'Found {len(messages)} messages.')
        return messages
        
    except Exception as error:
        print(f'An error occurred: {error}')
        return []

def list_msg_with_title(service, user_id='me', max_results=10, title= " Action Required: PO"):
    """Lists messages in the user's account."""
    try:
        # Get list of messages
        response = service.users().messages().list(
            userId=user_id, maxResults=max_results,  q=f'subject:"{title}"').execute()
        messages = response.get('messages', [])
        
        if not messages:
            print('No messages found.')
            return []
            
        print(f'Found {len(messages)} messages.')
        return messages
        
    except Exception as error:
        print(f'An error occurred: {error}')
        return []

def get_message_details(service, msg_id, user_id='me'):
    """Get message details for a specific message ID."""
    try:
        # Get full message details
        message = service.users().messages().get(
            userId=user_id, id=msg_id).execute()
        
        # Extract header data
        headers = message['payload']['headers']
        payload = message['payload']
        parts = payload.get('parts')
        subject = next((h['value'] for h in headers if h['name'].lower() == 'subject'), 'No Subject')
        sender = next((h['value'] for h in headers if h['name'].lower() == 'from'), 'Unknown')
        date = next((h['value'] for h in headers if h['name'].lower() == 'date'), 'Unknown')
        
        print(f'From: {sender}')
        print(f'Subject: {subject}')
        print(f'Date: {date}')
        print('-' * 40)
        
        # Extract and print body
        print("\nBody:")
        if parts:
            for part in parts:
                if part['mimeType'] == 'text/plain':
                    data = part['body']['data']
                    decoded_data = base64.urlsafe_b64decode(data).decode()
                    print(decoded_data)
                elif part['mimeType'] == 'text/html': #handle html emails
                    data = part['body']['data']
                    decoded_data = base64.urlsafe_b64decode(data).decode()
                    print(decoded_data)
        elif payload['mimeType'] == 'text/plain': #handles emails without parts
            data = payload['body']['data']
            decoded_data = base64.urlsafe_b64decode(data).decode()
            print(decoded_data)
        elif payload['mimeType'] == 'text/html':
            data = payload['body']['data']
            decoded_data = base64.urlsafe_b64decode(data).decode()
            print(decoded_data)
        
        return message
        
    except Exception as error:
        print(f'An error occurred: {error}')
        return None

def main():
    # Get Gmail API service
    service = get_gmail_service()
    
    # List recent messages
    messages = list_msg_with_title(service)
    
    # Get details for each message
    for message in messages[:5]:  # Limit to first 5 for demonstration
        print('\nRetrieving message...')
        get_message_details(service, message['id'])

if __name__ == '__main__':
    main()