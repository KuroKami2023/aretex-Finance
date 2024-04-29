import pandas as pd
import os

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

SPREADSHEET_ID = "1i5eUPdxRXW40Du8_DnnvqOMs0kSXcW_bdyRn0c0kLos"

# Excel file
df = pd.read_excel('sample_data.xlsx', sheet_name='Financials')

# Rows
def find_row_index(keyword):
    return df[df.iloc[:, 5].str.contains(keyword, na=False)].index[0]

# Indices
total_assets_index = find_row_index('Total Assets')
consolidated_assets_index = find_row_index('Total Assets')
total_liabilities_index = find_row_index('Total Liabilities')
total_equity_index = find_row_index('Total Equity')
month_index = find_row_index('MONTH')

# Date
month_row = df.iloc[month_index]
formatted_month = month_row.iloc[19].strftime('%Y-%m-%d')

# Assets
total_assets_row = df.iloc[total_assets_index]

# Liabilities
total_liabilities_row = df.iloc[total_liabilities_index]

# Equity
total_equity_row = df.iloc[total_equity_index]

def main():
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

        result = sheets.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range="Finances!A:A"
        ).execute()
        existing_dates = result.get('values', [])
        
        month = formatted_month

        if [month] not in existing_dates:
            total_assets = total_assets_row.iloc[11]
            total_liabilities = total_liabilities_row.iloc[11]
            total_equity = total_equity_row.iloc[11]

            values = [
                [month, "Total Assets", total_assets],
                [month, "Total Liabilities", total_liabilities],
                [month, "Total Equity", total_equity]
            ]

            body = {
                "values": values
            }

            last_row_index = len(existing_dates) + 1

            result = sheets.values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"Finances!A{last_row_index}",
                valueInputOption="USER_ENTERED",
                body=body
            ).execute()

            print("Data successfully appended to Google Spreadsheet.")
        else:
            print("Data for this month already exists in the spreadsheet.")

    except HttpError as error:
        print(error)

if __name__ == "__main__":
    main()
