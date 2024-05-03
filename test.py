import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

# Define the scope
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

# Add the credentials to the account
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)

# Authorize the clientsheet
client = gspread.authorize(creds)

# Open the Excel file
excel_file = pd.read_excel('sample_data.xlsx')

# Get the first sheet name
sheet_name = excel_file.sheet_names[0]

# Create a new Google Sheets workbook
workbook = client.create('New Google Sheets Workbook')

# Get the first sheet of the newly created workbook
sheet = workbook.get_worksheet(0)

# Convert the Excel file to a DataFrame
df = pd.DataFrame(excel_file.parse(sheet_name))

# Write the DataFrame to Google Sheets
client.import_dataframe(sheet.id, df)

print("Excel file imported to Google Sheets successfully.")
