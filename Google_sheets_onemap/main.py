from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
import gspread
import requests
import time

# Define the credentials JSON path
CREDS_JSON_PATH = ''

# Google Sheets scopes
SCOPES = ['https://spreadsheets.google.com/feeds',
          'https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive']

# Authenticate with gspread using google-auth
gc = gspread.service_account(filename=CREDS_JSON_PATH)
sheet_id = ""
reference_sheet_name = "Reference (Active Sheet for Address)"

# Function to initialize Google Sheets connection
def initialize_sheet():
    try:
        reference_sheet = gc.open_by_key(sheet_id).worksheet(reference_sheet_name)
        dynamic_sheet_name = reference_sheet.cell(1, 2).value
        print(f"Working on sheet: {dynamic_sheet_name}")
        return gc.open_by_key(sheet_id).worksheet(dynamic_sheet_name)
    except Exception as e:
        print(f"Error initializing sheet: {e}")
        return None

# Set up the Drive API client
creds = Credentials.from_service_account_file(CREDS_JSON_PATH, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=creds)

def update_address_from_postal_code(sheet):
    try:
        # Get all values from the sheet
        all_values = sheet.get_all_values()
        
        # Get the indices of the 'Postal' and 'Location' columns
        postal_col_idx = all_values[1].index('Postal')
        location_col_idx = all_values[1].index('Location')

        # Get the column values for 'Postal' and 'Location'
        postal_values = [row[postal_col_idx] for row in all_values[2:]]
        location_values = [row[location_col_idx] for row in all_values[2:]]

        for index, (postal_code, address) in enumerate(zip(postal_values, location_values)):
            if postal_code and not address:
                # Fetch address from OneMap API
                response = requests.get(f"https://www.onemap.gov.sg/api/common/elastic/search?searchVal={postal_code}&returnGeom=N&getAddrDetails=Y&pageNum=1")
                new_address = response.json()['results'][0]['ADDRESS']
                
                # Remove the last word from the address string
                new_address = " ".join(new_address.split()[:-2])

                # Update the address in Google Sheet
                sheet.update_cell(index + 3, location_col_idx + 1, new_address)
                print('Updated row')
    except Exception as e:
        print(f"Error processing row {index + 2}: {e}")

FREQUENCY_SCRIPT_SEC = 10

# Main loop
while True:
    try:
        sheet = initialize_sheet()  # Moved this inside the loop to reinitialize each iteration
        
        if sheet:  # Only proceed if sheet is initialized properly
            update_address_from_postal_code(sheet)  # Passed the sheet as an argument
            time.sleep(FREQUENCY_SCRIPT_SEC)
            print(f"Pending Changes")
        else:  # If sheet is None (due to error), wait before trying again
            time.sleep(FREQUENCY_SCRIPT_SEC)
    except Exception as e:
        print(f"Error in main loop: {e}")
        time.sleep(FREQUENCY_SCRIPT_SEC)
