from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import pickle
import csv
from datetime import datetime, timezone
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
import os
from email.utils import parsedate_to_datetime
import pandas as pd
from datetime import datetime, timezone

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/spreadsheets"
]

EXCEL_PATH = "result/log2.xlsx"

def authenticate():
    flow = InstalledAppFlow.from_client_secrets_file(
        'credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)
    # Save token for reuse
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)
    return creds

import pickle

def load_credentials():
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
    return creds




def get_unread_question_emails(service, excel_path):
    # Load existing IDs from Excel
    try:
        df = pd.read_excel(excel_path)
        logged_ids = set(df['MessageID'].astype(str))
    except FileNotFoundError:
        logged_ids = set()
    
    # Retrieve unread messages
    results = service.users().messages().list(
        userId='me',
        labelIds=['INBOX'],
        q='is:unread'
    ).execute()
    messages = results.get('messages', [])
    question_emails = []

    for msg in messages:
        message_id = msg['id']

        # Skip if already logged
        if message_id in logged_ids:
            continue

        message = service.users().messages().get(
            userId='me',
            id=message_id,
            format='metadata',
            metadataHeaders=['From', 'Subject', 'Date']
        ).execute()

        headers = {h['name']: h['value'] for h in message['payload']['headers']}
        
        if '@Question' in headers.get('Subject', ''):
            # Attach message ID so you can save it later
            headers['MessageID'] = message_id
            question_emails.append(headers)

    return question_emails


def append_to_sheet(sheet_service, spreadsheet_id, data):
    body = {
        'values': [data]
    }
    sheet_service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range='Sheet1!A:E',
        valueInputOption='USER_ENTERED',
        body=body
    ).execute()



def append_to_csv(filename, records):
        """
        records: List of lists, e.g.,
        [
            ["someone@gmail.com", "2025-07-01", "12:00:00", "0:15:00", "Not started"],
            ...
        ]
        """
        # If file doesn't exist yet, write header
        try:
            with open(filename, 'x', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Email Sender", "Date", "Time Received", "Wait Time", "Status"])
                for row in records:
                    writer.writerow(row)
        except FileExistsError:
            # File exists, just append rows
            with open(filename, 'a', newline='') as csvfile:
                writer = csv.writer(csvfile)
                for row in records:
                    writer.writerow(row)




def append_to_excel(filename, records):
    if not os.path.exists(filename):
        # Create workbook and add headers
        wb = Workbook()
        ws = wb.active
        ws.title = "Activity Log"

        headers = ["Message_ID", "Email Sender", "Date", "Time Received", "Wait Time", "Status"]
        ws.append(headers)

        # Add dropdown validation to Status column
        dv = DataValidation(type="list", formula1='"Not started, In Progress, Done"', allow_blank=True)
        ws.add_data_validation(dv)

        # Apply validation to all rows in Status column
        max_row = 1048576  # Excel max rows
        status_col = get_column_letter(headers.index("Status")+1)  # E
        dv.add(f"{status_col}2:{status_col}{max_row}")

        wb.save(filename)
        print(f"✅ Created {filename} with headers and dropdowns.")

    # Load the workbook to append
    wb = load_workbook(filename)
    ws = wb.active

    # Append each record
    for row in records:
        ws.append(row)

    wb.save(filename)
    print(f"✅ Appended {len(records)} record(s) to {filename}.")

def format_wait_time(delta):
    """
    delta: datetime.timedelta
    Returns a string like "2 days, 3 hours, 15 minutes"
    """
    total_seconds = int(delta.total_seconds())
    days = total_seconds // 86400
    hours = (total_seconds % 86400) // 3600
    minutes = (total_seconds % 3600) // 60
    return f"{days} days, {hours} hours, {minutes} minutes"



def main():
    creds = load_credentials()
    
    gmail_service = build('gmail', 'v1', credentials=creds)

    emails = get_unread_question_emails(gmail_service, "result/log.xlsx")

    for email in emails:
        sender = email.get('From')
        date_str = email.get('Date')

        # Parse date correctly
        email_date = parsedate_to_datetime(date_str)
        if email_date.tzinfo is None:
            email_date = email_date.replace(tzinfo=timezone.utc)
        email_date = email_date.astimezone(timezone.utc)

        now = datetime.now(timezone.utc)
        delta = now - email_date
        wait_time = format_wait_time(delta)

        date_received = email_date.strftime('%Y-%m-%d')
        time_received = email_date.strftime('%H:%M:%S')

        message_id = email['MessageID'] 

        row = [
            message_id,
            sender,
            date_received,
            time_received,
            wait_time,
            "Not started"
        ]

        # check to see if result folder exists
        if not os.path.exists("result"):
            os.makedirs("result")

        append_to_excel(EXCEL_PATH, [row])

  

if __name__ == '__main__':
    main()