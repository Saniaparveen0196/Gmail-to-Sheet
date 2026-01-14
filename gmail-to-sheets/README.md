Gmail to Google Sheets Automation

A Python automation script that reads unread emails from Gmail and logs them into a Google Sheet using Google APIs. The project uses OAuth 2.0 authentication, prevents duplicate entries, and maintains state across runs.

1️⃣High-Level Architecture
                ┌─────────────────────┐
                │     Gmail API        │
                │ (Unread Emails)      │
                └──────────┬──────────┘
                           │
                           ▼
                ┌─────────────────────┐
                │   Gmail Service      │
                │ - OAuth Auth         │
                │ - Fetch Emails       │
                │ - Mark as Read       │
                └──────────┬──────────┘
                           │
                           ▼
                ┌─────────────────────┐
                │   Email Parser       │
                │ - Extract headers    │
                │ - Parse body & date  │
                └──────────┬──────────┘
                           │
                           ▼
                ┌─────────────────────┐
                │  Sheets Service      │
                │ - OAuth Auth         │
                │ - Duplicate Check   │
                │ - Append Rows       │
                └──────────┬──────────┘
                           │
                           ▼
                ┌─────────────────────┐
                │  Google Sheets      │
                │  (Email Log)        │
                └─────────────────────┘
                           ▲
                           │
                ┌─────────────────────┐
                │  State Persistence  │
                │  (state.json)       │
                └─────────────────────┘
2️⃣ Setup Instructions (Step-by-Step)
Prerequisites

Python 3.7+

Gmail account

Google Cloud account

Google Sheet

Step 1: Google Cloud Setup

Create a project in Google Cloud Console

Enable:

Gmail API

Google Sheets API

Create OAuth Client ID

Type: Desktop App

Download credentials

Save file as:

credentials/credentials.json

Step 2: Google Sheet Setup

Create a Google Sheet

Copy Spreadsheet ID from URL

Add it to config.py

Step 3: Install Dependencies
pip install -r requirements.txt

Step 4: Run the Script
python src/main.py


Browser opens for OAuth (first run)

Permissions granted

Emails are logged to the sheet

3️⃣ OAuth Flow Used

Uses OAuth 2.0 Authorization Code Flow

User grants Gmail & Sheets permissions

Access and refresh tokens are stored locally

Tokens are reused on subsequent runs

Why OAuth?

Gmail API requires user-based access

Secure and revocable by the user

4️⃣ Duplicate Prevention Logic

Duplicates are prevented using two methods:

Gmail filter

is:unread in:inbox


Only unread emails are fetched.

State file tracking

File: credentials/state.json

Stores processed Gmail message IDs

Emails already processed are skipped

This ensures the script can be run multiple times safely.

5️⃣ State Persistence Method

Uses a JSON file (state.json)

Stores:

Processed email IDs

Last run timestamp

State is loaded at startup and saved after execution

Why JSON?

Simple

Lightweight

No database required

6️⃣ Challenge Faced & Solution
Challenge: Preventing Duplicate Entries on Multiple Runs

Problem:
When the script was executed multiple times, the same emails could be fetched again, leading to duplicate rows being added to the Google Sheet.

Solution:

Implemented a state persistence mechanism using a local JSON file (state.json)

Each processed Gmail message ID is stored after successful insertion

Before processing an email, the script checks whether the message ID already exists

Additionally, emails are marked as read after processing

Result:

Running the script multiple times does not create duplicate entries

Only new, unread emails are processed

The automation is safe to rerun at any time
7️⃣ Limitations

Works for a single Gmail account

Uses file-based state (not suitable for very large scale)

Basic email body parsing

No rollback if script stops midway

📌 Conclusion

This project demonstrates:

OAuth-based API integration

Gmail automation

Google Sheets logging

Duplicate prevention

Persistent state handling

Author: Sania Parveen
Tech Stack: Python, Gmail API, Google Sheets API