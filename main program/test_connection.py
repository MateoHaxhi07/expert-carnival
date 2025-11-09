"""
SIMPLE CONNECTION TEST
======================
Run this FIRST to verify your Google Sheets credentials work.

This is a quick test - if this works, the main program will work too!
"""

import gspread
from oauth2client.service_account import ServiceAccountCredentials

print("\n" + "="*60)
print("🔍 TESTING GOOGLE SHEETS CONNECTION")
print("="*60)

# Your Google Sheet URL
SHEET_URL = "https://docs.google.com/spreadsheets/d/13XPpCghDGXQzaciU3DR4eFz6TW3IXYmcq7p9300Gpu4/edit?gid=1476407013#gid=1476407013"

try:
    print("\n1️⃣ Loading credentials.json...")
    
    # Define scope
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]
    
    # Load credentials
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    print("   ✅ Credentials loaded successfully!")
    
    print("\n2️⃣ Connecting to Google...")
    client = gspread.authorize(creds)
    print("   ✅ Authenticated with Google!")
    
    print("\n3️⃣ Opening your spreadsheet...")
    sheet = client.open_by_url(SHEET_URL)
    print(f"   ✅ Opened spreadsheet: '{sheet.title}'")
    
    print("\n4️⃣ Reading data...")
    worksheet = sheet.get_worksheet(0)
    print(f"   ✅ Found worksheet: '{worksheet.title}'")
    
    # Get row count
    data = worksheet.get_all_records()
    print(f"   ✅ Loaded {len(data)} rows of data!")
    
    # Show first few column names
    if data:
        columns = list(data[0].keys())
        print(f"\n📋 Columns found: {', '.join(columns[:5])}...")
    
    print("\n" + "="*60)
    print("✅ SUCCESS! YOUR CONNECTION WORKS!")
    print("="*60)
    print("\n💡 Next step: Run the main program:")
    print("   python restaurant_report_generator.py")
    print("\n")
    
except FileNotFoundError:
    print("\n❌ ERROR: credentials.json not found!")
    print("\n💡 Make sure:")
    print("   1. You downloaded the JSON key from Google Cloud")
    print("   2. You renamed it to 'credentials.json'")
    print("   3. You put it in the same folder as this script")
    print("\n")
    
except Exception as e:
    print(f"\n❌ ERROR: {str(e)}")
    print("\n💡 Common causes:")
    print("   1. Haven't shared the sheet with service account email")
    print("   2. Wrong spreadsheet URL")
    print("   3. Haven't enabled Google Sheets API")
    print("\n💡 Solution:")
    print("   - Check YOUR_QUICK_START.md STEP 3")
    print("   - Make sure you shared the sheet with the service account")
    print("\n")