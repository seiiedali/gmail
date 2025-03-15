import os
import pickle
import base64
import tkinter as tk
from tkinter import messagebox
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
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]


def get_gmail_service():
    """Gets authenticated Gmail API service."""
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        # if creds and creds.expired and creds.refresh_token:
        #     creds.refresh(Request())
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
        creds = flow.run_local_server(port=0)
    with open("token.pickle", "wb") as token:
        pickle.dump(creds, token)
    return build("gmail", "v1", credentials=creds)


def list_msg_with_title(
    service, user_id="me", max_results=10, title=" Action Required: PO"
):
    try:
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
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id).execute()
        headers = message["payload"]["headers"]
        parts = message["payload"].get("parts")
        date = next(
            (h["value"] for h in headers if h["name"].lower() == "date"), "Unknown"
        )
        if parts:
            for part in parts:
                if part["mimeType"] == "text/html":
                    data = part["body"]["data"]
                    sanitized_date = (
                        date.replace(":", "-").replace(",", "").replace(" ", "_")
                    )
                    with open(
                        f"{EMAIL_DIR_NAME}/{sanitized_date}{msg_id}.html",
                        "w",
                        encoding="utf-8",
                    ) as file:
                        file.write(base64.urlsafe_b64decode(data).decode())
    except Exception as error:
        print(f"An error occurred in retrieving message data: {error}")


def update_html():
    service = get_gmail_service()
    messages = list_msg_with_title(service, title=WAYFAIR_TITLE)
    if not os.path.exists(EMAIL_DIR_NAME):
        os.makedirs(EMAIL_DIR_NAME)
    for message in messages:
        save_message_html(service, message["id"])
    messagebox.showinfo("Success", "HTML inventory updated successfully!")


def update_db():
    setup_database()
    html_files = [
        file for file in os.listdir(f"./{EMAIL_DIR_NAME}") if file.endswith(".html")
    ]
    for html_file in html_files:
        with open(f"{EMAIL_DIR_NAME}/{html_file}", "r", encoding="utf-8") as file:
            html_content = file.read()
            customer, order, products, order_items = extract_data_from_html(
                html_content
            )
            save_to_database(customer, order, products, order_items)
    messagebox.showinfo("Success", "Database updated successfully!")


def update_excel():
    export_to_excel()
    messagebox.showinfo("Success", "Excel file updated successfully!")


def update_all():
    update_html()
    update_db()
    update_excel()


def main():
    root = tk.Tk()
    root.title("Email Data Processor")
    root.geometry("400x300")  # Set a fixed window size
    root.resizable(False, False)  # Disable resizing

    # Add a frame for better layout management
    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(expand=True)

    # Add a title label
    tk.Label(frame, text="Email Data Processor", font=("Arial", 16, "bold")).pack(
        pady=10
    )

    # Add buttons with side padding
    tk.Button(
        frame, text="Update HTML Inventory", command=update_html, width=25, height=2
    ).pack(pady=5, padx=10)
    tk.Button(
        frame, text="Update Database", command=update_db, width=25, height=2
    ).pack(pady=5, padx=10)
    tk.Button(
        frame, text="Update Excel File", command=update_excel, width=25, height=2
    ).pack(pady=5, padx=10)
    tk.Button(
        frame,
        text="Do All Tasks",
        command=update_all,
        width=25,
        height=2,
        bg="#4CAF50",
        fg="white",
    ).pack(pady=10, padx=10)

    # Start the main loop
    root.mainloop()


if __name__ == "__main__":
    main()
