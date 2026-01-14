"""
Configuration file for Gmail to Sheets automation.
"""
import os

# Google API Configuration
SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/spreadsheets'
]

# Paths
CREDENTIALS_DIR = os.path.join(os.path.dirname(__file__), 'credentials')
CREDENTIALS_FILE = os.path.join(CREDENTIALS_DIR, 'credentials.json')
TOKEN_FILE = os.path.join(CREDENTIALS_DIR, 'token.json')
STATE_FILE = os.path.join(CREDENTIALS_DIR, 'state.json')

# Gmail Configuration
GMAIL_QUERY = 'is:unread in:inbox'  # Unread emails in inbox

# Subject-based filtering (optional)
# Set to a keyword string to filter emails by subject (case-insensitive)
# Example: SUBJECT_FILTER = "Invoice" will only process emails with "Invoice" in subject
# Set to None to disable filtering
SUBJECT_FILTER = os.getenv('SUBJECT_FILTER', None)  # Can be set via environment variable

# Google Sheets Configuration
# Can be set via environment variables or hardcoded here
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID', 'spreadsheet's id')
SHEET_NAME = os.getenv('SHEET_NAME', 'sheet name')

# Email Processing Configuration
MAX_RESULTS = int(os.getenv('MAX_RESULTS', '50'))  # Maximum number of emails to process per run
