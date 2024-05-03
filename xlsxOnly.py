import pandas as pd
import os
from datetime import datetime
import sys
import json

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

SPREADSHEET_ID = "1Q6vPAovgHOffrr8uQDawW7M59wFNV14kI9eltW8P8RA"

def main():

    file_name = "sample_data.xlsx"

    if file_name.endswith('.csv'):
        df = pd.read_csv(file_name, encoding='latin1', dtype=str)
    elif file_name.endswith('.xlsx'):
        excel_data = pd.read_excel(file_name, engine='openpyxl', sheet_name=None)
        clear_all_sheets()
        for sheet_name, sheet_df in excel_data.items():
            create_sheet(sheet_name)
            upload_sheet_to_google(sheet_name, sheet_df)
    else:
        print("Unsupported file type")
        sys.exit(1)

def clear_all_sheets():
    credentials = None
    if os.path.exists("token.json"):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(credentials.to_json())

    service = build("sheets", "v4", credentials=credentials)
    sheets = service.spreadsheets()
    
    sheet_metadata = sheets.get(spreadsheetId=SPREADSHEET_ID).execute()
    sheet_list = sheet_metadata.get('sheets', [])

    if len(sheet_list) <= 1:
        print("There must be at least one sheet in the spreadsheet.")
        return
    
    for sheet in sheet_list[1:]:
        sheet_id = sheet['properties']['sheetId']
        sheets.batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": [{"deleteSheet": {"sheetId": sheet_id}}]}).execute()


def create_sheet(sheet_name):
    credentials = None
    if os.path.exists("token.json"):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(credentials.to_json())

    service = build("sheets", "v4", credentials=credentials)
    sheets = service.spreadsheets()
    
    body = {
        "requests": [
            {
                "addSheet": {
                    "properties": {
                        "title": sheet_name
                    }
                }
            }
        ]
    }
    sheets.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()

def upload_sheet_to_google(sheet_name, df):
    df.fillna('', inplace=True)

    credentials = None
    if os.path.exists("token.json"):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(credentials.to_json())

    try:
        service = build("sheets", "v4", credentials=credentials)
        sheets = service.spreadsheets()

        values = df.values.tolist()
        sheets.values().clear(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!A1:ZZ",
            body={}
        ).execute()
        for row in values:
            for i, value in enumerate(row):
                if isinstance(value, datetime):
                    row[i] = value.strftime('%Y-%m-%d %H:%M:%S')

        result = sheets.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!A1",
            valueInputOption="USER_ENTERED",
            body={"values": values}
        ).execute()

        print(f"Data from sheet '{sheet_name}' successfully uploaded to Google Spreadsheet.")

    except HttpError as error:
        print(error)

if __name__ == "__main__":
    main()
