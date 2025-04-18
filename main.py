import os
import pickle
import base64
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk  # Import ttk for the progress bar
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from extract import (
    setup_database,
    extract_data_from_html,
    save_to_database,
    export_to_excel,
)
import sqlite3
import pandas as pd

WAYFAIR_TITLE = "Action Required: PO"
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]


def get_gmail_service():
    """Gets authenticated Gmail API service."""
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
        creds = flow.run_local_server(port=0)
    with open("token.pickle", "wb") as token:
        pickle.dump(creds, token)
    return build("gmail", "v1", credentials=creds)


def list_msg_with_title(
    service, user_id="me", max_results=10, title="Action Required: PO"
):
    """List messages with the specified title, excluding already processed ones."""
    try:
        # Fetch already processed email IDs from the database
        conn = sqlite3.connect("orders.db")
        cursor = conn.cursor()
        cursor.execute("SELECT message_id FROM processed_emails")
        processed_ids = {row[0] for row in cursor.fetchall()}
        conn.close()

        # Fetch messages from Gmail
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

        # Filter out already processed messages
        new_messages = [msg for msg in messages if msg["id"] not in processed_ids]
        print(f"Found {len(new_messages)} new messages.")
        return new_messages
    except Exception as error:
        print(f"An error occurred: {error}")
        return []


def process_message_data(service, msg_id, user_id="me"):
    """Process the content of a Gmail message."""
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id).execute()
        headers = message["payload"]["headers"]
        parts = message["payload"].get("parts")
        if parts:
            for part in parts:
                if part["mimeType"] == "text/html":
                    data = part["body"]["data"]
                    html_content = base64.urlsafe_b64decode(data).decode()

                    # Extract and save data to the database
                    customer, order, products, order_items = extract_data_from_html(
                        html_content
                    )
                    save_to_database(customer, order, products, order_items)

                    # Save the processed email record
                    conn = sqlite3.connect("orders.db")
                    cursor = conn.cursor()
                    cursor.execute(
                        """INSERT INTO processed_emails (message_id, processed_at)
                           VALUES (?, ?)""",
                        (msg_id, pd.Timestamp.now().isoformat()),
                    )
                    conn.commit()
                    conn.close()
    except Exception as error:
        print(f"An error occurred in processing message data: {error}")


def extract_emails():
    """Extract emails from Gmail and save data to the database."""
    setup_database()  # Ensure the database is set up
    service = get_gmail_service()
    messages = list_msg_with_title(service, title=WAYFAIR_TITLE)

    if not messages:
        messagebox.showinfo("Info", "No new emails to process.")
        return

    # Configure the progress bar
    progress_bar["maximum"] = len(messages)
    progress_bar["value"] = 0

    # Display the total number of new messages
    total_new_messages = len(messages)
    status_label.config(text=f"Processing: 0/{total_new_messages} new messages")
    root.update_idletasks()

    for i, message in enumerate(messages):
        process_message_data(service, message["id"])
        progress_bar["value"] = i + 1  # Update progress bar
        status_label.config(
            text=f"Processing: {i + 1}/{total_new_messages} new messages"
        )
        root.update_idletasks()  # Refresh the GUI

    # Log the extraction call
    conn = sqlite3.connect("orders.db")
    cursor = conn.cursor()
    cursor.execute(
        """INSERT INTO extraction_logs (timestamp) VALUES (?)""",
        (pd.Timestamp.now().isoformat(),),
    )
    conn.commit()
    conn.close()

    messagebox.showinfo("Success", "Emails processed and data saved to the database!")
    update_status()


def update_status():
    """Update the status in the GUI."""
    setup_database()  # Ensure the database is set up
    conn = sqlite3.connect("orders.db")
    cursor = conn.cursor()

    # Get counts
    cursor.execute("SELECT COUNT(*) FROM processed_emails")
    processed_emails_count = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM customers")
    customers_count = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM orders")
    orders_count = cursor.fetchone()[0]

    # Get the last extraction timestamp
    cursor.execute("SELECT MAX(timestamp) FROM extraction_logs")
    last_updated = cursor.fetchone()[0]

    conn.close()

    # Update the status label
    status_label.config(
        text=f"Processed Emails: {processed_emails_count}\n"
        f"Customers: {customers_count}\n"
        f"Orders: {orders_count}\n"
        f"Last Updated: {last_updated}"
    )


def create_excel_file():
    """Create an Excel file from the database."""
    export_to_excel()
    messagebox.showinfo("Success", "Excel file created successfully!")


def main():
    global status_label, progress_bar, root  # Make widgets accessible globally

    root = tk.Tk()
    root.title("Email Data Processor")
    root.geometry("400x400")
    root.resizable(False, False)

    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(expand=True)

    tk.Label(frame, text="Email Data Processor", font=("Arial", 16, "bold")).pack(
        pady=10
    )

    # Buttons
    tk.Button(
        frame, text="Extract Emails", command=extract_emails, width=25, height=2
    ).pack(pady=10, padx=10)
    tk.Button(
        frame, text="Create Excel File", command=create_excel_file, width=25, height=2
    ).pack(pady=10, padx=10)

    # Progress bar
    progress_bar = ttk.Progressbar(
        frame, orient="horizontal", length=300, mode="determinate"
    )
    progress_bar.pack(pady=10)

    # Status label
    status_label = tk.Label(frame, text="", font=("Arial", 12), justify="left")
    status_label.pack(pady=20)

    # Initial status update
    update_status()

    root.mainloop()


if __name__ == "__main__":
    main()
