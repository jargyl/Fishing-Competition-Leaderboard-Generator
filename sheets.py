import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('key.json', scope)
client = gspread.authorize(creds)

# Get the current date and time in European format
now = datetime.now()
date_string = now.strftime("%d/%m/%Y")

# Create a new Google Sheets file with the desired sheet names and the current date in the name
base_name = 'De Zwaan Scorebord'
file_name = f"{base_name} - {date_string}"
sheet_names = ['Group 1', 'Group 2', 'Group 3', 'Group 4', 'Global Leaderboard']
spreadsheet = client.create(file_name)
for sheet_name in sheet_names:
    worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="20")

# Set the sharing permissions for the file to be viewable by anyone with the link
spreadsheet.share(None, perm_type='anyone', role='reader')

# Delete the default Sheet1 worksheet
default_worksheet = spreadsheet.worksheet('Sheet1')
spreadsheet.del_worksheet(default_worksheet)

# Read all sheets of the local Excel file into a dictionary of DataFrames
dfs = pd.read_excel('leaderboards.xlsx', sheet_name=None)

# Loop through the sheets and write each one to the corresponding sheet in the new Google Sheets file
for sheet_name, df in dfs.items():
    # Convert non-numeric columns to string
    df = df.fillna('')
    # Write the pandas dataframe to the new Google Sheets file
    worksheet = spreadsheet.worksheet(sheet_name)
    worksheet.update([df.columns.values.tolist()] + df.values.tolist())

# Print the URL of the new Google Sheets file
print(f"New Google Sheets file URL: {spreadsheet.url}")
