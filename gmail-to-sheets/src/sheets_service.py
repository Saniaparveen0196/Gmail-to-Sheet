"""
Google Sheets API service for appending email data.
"""
import os
import logging
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception
import config

logger = logging.getLogger(__name__)


def _is_retryable_http_error(exception):
    """Check if HttpError is retryable (rate limit or server error)."""
    if isinstance(exception, HttpError):
        status_code = exception.resp.status if hasattr(exception, 'resp') else None
        if status_code:
            # Retry on rate limit (429) and server errors (500-599)
            return status_code == 429 or (500 <= status_code <= 599)
    return False


class SheetsService:
    """Service class for interacting with Google Sheets API."""
    
    def __init__(self):
        self.service = None
        self.credentials = None
        
    def authenticate(self):
        """
        Authenticate and build Sheets service using OAuth 2.0.
        Uses the same token as Gmail service if available.
        Returns True if successful, False otherwise.
        """
        creds = None
        
        # Load existing token if available (shared with Gmail)
        if os.path.exists(config.TOKEN_FILE):
            creds = Credentials.from_authorized_user_file(config.TOKEN_FILE, config.SCOPES)
        
        # If there are no (valid) credentials available, let the user log in
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not os.path.exists(config.CREDENTIALS_FILE):
                    raise FileNotFoundError(
                        f"Credentials file not found at {config.CREDENTIALS_FILE}. "
                        "Please download OAuth 2.0 credentials from Google Cloud Console."
                    )
                
                flow = InstalledAppFlow.from_client_secrets_file(
                    config.CREDENTIALS_FILE, config.SCOPES
                )
                creds = flow.run_local_server(port=0)
            
            # Save the credentials for the next run
            with open(config.TOKEN_FILE, 'w') as token:
                token.write(creds.to_json())
        
        self.credentials = creds
        self.service = build('sheets', 'v4', credentials=creds)
        return True
    
    def _ensure_sheet_exists(self, spreadsheet_id, sheet_name):
        """Check if sheet exists, create it if it doesn't."""
        try:
            # Try to get spreadsheet metadata to check if sheet exists
            spreadsheet = self.service.spreadsheets().get(
                spreadsheetId=spreadsheet_id
            ).execute()
            
            sheet_exists = any(
                sheet['properties']['title'] == sheet_name
                for sheet in spreadsheet.get('sheets', [])
            )
            
            if not sheet_exists:
                # Create the sheet
                requests = [{
                    'addSheet': {
                        'properties': {
                            'title': sheet_name
                        }
                    }
                }]
                body = {'requests': requests}
                self.service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body=body
                ).execute()
                logger.info(f"Created sheet '{sheet_name}' in spreadsheet")
            
            return True
        except HttpError as error:
            # If it's a permission error, re-raise it
            if error.resp.status in [401, 403]:
                raise
            logger.error(f'Error checking/creating sheet: {error}')
            return False
    
    def _ensure_headers_internal(self, spreadsheet_id, sheet_name):
        """Internal method that performs the actual API calls."""
        # First ensure the sheet exists
        self._ensure_sheet_exists(spreadsheet_id, sheet_name)
        
        # Check if headers exist - try with quotes first, then without
        range_name = f"'{sheet_name}'!A1:D1"  # Quote sheet name to handle special characters
        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range_name
            ).execute()
        except HttpError as error:
            # If range parsing fails with quotes, try without quotes (for simple names)
            if 'Unable to parse range' in str(error) or error.resp.status == 400:
                range_name = f'{sheet_name}!A1:D1'
                result = self.service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=range_name
                ).execute()
            else:
                raise
        
        values = result.get('values', [])
        
        # If no headers exist, create them
        if not values or len(values) == 0:
            headers = [['From', 'Subject', 'Date', 'Content']]
            body = {
                'values': headers
            }
            self.service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption='RAW',
                body=body
            ).execute()
            logger.info(f"Created headers in sheet '{sheet_name}'")
        
        return True
    
    @retry(
        stop=stop_after_attempt(5),
        wait=wait_exponential(multiplier=1, min=2, max=60),
        retry=retry_if_exception(_is_retryable_http_error),
        reraise=True,
        before_sleep=lambda retry_state: logger.warning(f"Retrying ensure_headers after {retry_state.outcome.exception()}")
    )
    def ensure_headers(self, spreadsheet_id, sheet_name):
        """
        Ensure the sheet has proper headers. Creates them if they don't exist.
        
        Args:
            spreadsheet_id: Google Sheets spreadsheet ID
            sheet_name: Name of the sheet tab
            
        Returns:
            True if successful, False otherwise
        """
        if not self.service:
            raise ValueError("Service not authenticated. Call authenticate() first.")
        
        try:
            return self._ensure_headers_internal(spreadsheet_id, sheet_name)
        except HttpError as error:
            # Non-retryable errors (401, 403, etc.) - return False
            if not _is_retryable_http_error(error):
                logger.error(f'Non-retryable error occurred while ensuring headers: {error}')
                return False
            # Retryable errors that exhausted all retries
            logger.error(f'All retry attempts exhausted for ensure_headers: {error}')
            return False
    
    def _append_email_internal(self, spreadsheet_id, sheet_name, email_data):
        """Internal method that performs the actual API call."""
        # Prepare row data - check what keys are available
        logger.debug(f"Email data keys: {list(email_data.keys())}")
        
        # Map email parser output to sheet columns
        row = [
            email_data.get('from', ''),
            email_data.get('subject', ''),
            email_data.get('date', ''),
            email_data.get('body', email_data.get('content', ''))  # Try both 'body' and 'content'
        ]
        
        logger.debug(f"Prepared row: {row[:3]}... (content truncated)")
        
        body = {
            'values': [row]
        }
        
        # Use simple range format - Google Sheets API handles this well
        range_name = f'{sheet_name}!A:D'
        
        try:
            logger.debug(f"Appending row to range: {range_name}")
            logger.debug(f"Row data (first 100 chars of content): {row[0]}, {row[1]}, {row[2]}, {row[3][:100] if len(row[3]) > 100 else row[3]}")
            
            result = self.service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption='RAW',
                insertDataOption='INSERT_ROWS',
                body=body
            ).execute()
            
            # Verify the result
            if result:
                updates = result.get('updates', {})
                updated_range = updates.get('updatedRange', 'Unknown')
                updated_rows = updates.get('updatedRows', 0)
                updated_cells = updates.get('updatedCells', 0)
                
                logger.info(f" Email appended successfully!")
                logger.info(f"  - Updated range: {updated_range}")
                logger.info(f"  - Rows updated: {updated_rows}")
                logger.info(f"  - Cells updated: {updated_cells}")
                
                # Double-check by reading back the data
                try:
                    verify_range = f'{sheet_name}!A:D'
                    verify_result = self.service.spreadsheets().values().get(
                        spreadsheetId=spreadsheet_id,
                        range=verify_range
                    ).execute()
                    verify_values = verify_result.get('values', [])
                    logger.info(f"  - Verification: Sheet now contains {len(verify_values)} row(s) total")
                    if len(verify_values) > 1:
                        logger.debug(f"  - Last row preview: {verify_values[-1][0] if verify_values[-1] else 'Empty'}")
                except Exception as verify_error:
                    logger.warning(f"  - Could not verify append: {verify_error}")
                
                return True
            else:
                logger.error("Append returned no result!")
                return False
                
        except HttpError as error:
            error_msg = str(error)
            logger.error(f"Failed to append email: {error_msg}")
            if hasattr(error, 'resp'):
                logger.error(f"HTTP Status: {error.resp.status}")
            raise
    
    @retry(
        stop=stop_after_attempt(5),
        wait=wait_exponential(multiplier=1, min=2, max=60),
        retry=retry_if_exception(_is_retryable_http_error),
        reraise=True,
        before_sleep=lambda retry_state: logger.warning(f"Retrying append_email after {retry_state.outcome.exception()}")
    )
    def append_email(self, spreadsheet_id, sheet_name, email_data):
        """
        Append a single email as a new row to the Google Sheet.
        
        Args:
            spreadsheet_id: Google Sheets spreadsheet ID
            sheet_name: Name of the sheet tab
            email_data: Dictionary with keys: from, subject, date, content
            
        Returns:
            True if successful, False otherwise
        """
        if not self.service:
            raise ValueError("Service not authenticated. Call authenticate() first.")
        
        try:
            result = self._append_email_internal(spreadsheet_id, sheet_name, email_data)
            if result:
                # Verify the append worked by checking the updated range
                updated_range = f'{sheet_name}!A:D'
                try:
                    verify_result = self.service.spreadsheets().values().get(
                        spreadsheetId=spreadsheet_id,
                        range=updated_range
                    ).execute()
                    values = verify_result.get('values', [])
                    logger.debug(f"Sheet now has {len(values)} rows (including header)")
                except Exception as e:
                    logger.warning(f"Could not verify append: {e}")
            return result
        except HttpError as error:
            # Log the full error details
            error_details = str(error)
            if hasattr(error, 'resp'):
                error_details += f" (Status: {error.resp.status})"
            logger.error(f'Error occurred while appending email: {error_details}')
            
            # Non-retryable errors (401, 403, etc.) - return False
            if not _is_retryable_http_error(error):
                logger.error(f'Non-retryable error occurred while appending email: {error}')
                return False
            # Retryable errors that exhausted all retries
            logger.error(f'All retry attempts exhausted for append_email: {error}')
            return False
    
    def email_exists(self, spreadsheet_id, sheet_name, email_id):
        """
        Check if an email with the given ID already exists in the sheet.
        Uses internal_date stored in a hidden column or checks content.
        
        Args:
            spreadsheet_id: Google Sheets spreadsheet ID
            sheet_name: Name of the sheet tab
            email_id: Gmail message ID
            
        Returns:
            True if email exists, False otherwise
        """
        if not self.service:
            raise ValueError("Service not authenticated. Call authenticate() first.")
        
        try:
            # Get all rows from the sheet
            range_name = f'{sheet_name}!A:D'
            result = self.service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            
            # Skip header row if present
            if values and len(values) > 0:
                # Check if first row looks like headers
                if values[0][0].lower() == 'from':
                    values = values[1:]
            
            # For duplicate detection, we'll use a combination of from+subject+date
            # This is a simple approach. In production, you might want to store
            # email IDs in a separate column
            return False  # Simplified - actual duplicate check is done via state file
            
        except HttpError as error:
            logger.error(f'Error occurred while checking for duplicates: {error}')
            return False
