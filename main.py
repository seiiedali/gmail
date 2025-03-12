import os
import pickle
import base64
import argparse

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

from extract import (
    setup_database,
    extract_data_from_html,
    save_to_database,
    export_to_excel,
)


EMAIL_DIR_NAME = "email_inventory"
WAYFAIR_TITLE = "Action Required: PO"

# Define the scopes
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]


def get_gmail_service():
    """Gets authenticated Gmail API service."""
    creds = None

    # Check if token.pickle file exists with stored credentials
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    # If credentials don't exist or are invalid, get new ones
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Load client secrets from downloaded file (rename to 'credentials.json')
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)

        # Save credentials for future use
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)

    # Return Gmail API service
    return build("gmail", "v1", credentials=creds)


def list_msg_with_title(service, user_id="me", max_results=10, title=" Action Required: PO"):
    """Lists messages in the user's account."""
    try:
        # Get list of messages
        response = (
            service.users()
            .messages()
            .list(userId=user_id, maxResults=max_results, q=f'subject:"{title}"')
            .execute()
        )
        messages = response.get("messages", [])

        if not messages:
            print("No messages found.")
            return []

        print(f"Found {len(messages)} messages.")
        return messages

    except Exception as error:
        print(f"An error occurred: {error}")
        return []


def save_message_html(service, msg_id, user_id="me"):
    """Get message details for a specific message ID."""
    try:
        # Get full message details
        message = service.users().messages().get(userId=user_id, id=msg_id).execute()

        # Extract header data
        headers = message["payload"]["headers"]
        parts = message["payload"].get("parts")
        date = next(
            (h["value"] for h in headers if h["name"].lower() == "date"), "Unknown"
        )

        decoded_data = []
        if parts:
            for part in parts:
                if part["mimeType"] == "text/plain":
                    continue
                elif part["mimeType"] == "text/html":  # handle html emails
                    data = part["body"]["data"]
                    decoded_data.append(base64.urlsafe_b64decode(data).decode())
                    sanitized_date = date.replace(":", "-").replace(",", "").replace(" ", "_")
                    with open(f"{EMAIL_DIR_NAME}/{sanitized_date}{msg_id}.html", "w", encoding="utf-8") as file:
                        file.write(base64.urlsafe_b64decode(data).decode())

    except Exception as error:
        print(f"An error occurred in retrieving message data: {error}")
        return None


def main():
    parser = argparse.ArgumentParser(description="Process email data.")
    parser.add_argument("--html", action="store_true", help="Update HTML inventory")
    parser.add_argument("--db", action="store_true", help="Update database")
    parser.add_argument("--excel", action="store_true", help="Update Excel file")
    args = parser.parse_args()

    # Get Gmail API service
    service = get_gmail_service()

    if args.html or not (args.html or args.db or args.excel):
        # List recent messages
        messages = list_msg_with_title(service, title=WAYFAIR_TITLE)

        # Create directory if it does not exist
        if not os.path.exists(EMAIL_DIR_NAME):
            os.makedirs(EMAIL_DIR_NAME)
        # Get details for each message
        for message in messages:
            save_message_html(service, message["id"])

    if args.db or not (args.html or args.db or args.excel):
        # List all .html files in the directory
        setup_database()
        html_files = [file for file in os.listdir(f"./{EMAIL_DIR_NAME}") if file.endswith(".html")]
        for html_file in html_files:
            with open(f"{EMAIL_DIR_NAME}/{html_file}", "r", encoding="utf-8") as file:
                html_content = file.read()
                customer, order, products, order_items = extract_data_from_html(html_content)
                save_to_database(customer, order, products, order_items)

    if args.excel or not (args.html or args.db or args.excel):
        export_to_excel()


if __name__ == "__main__":
    main()