"""
Gmail API service for fetching and managing emails.
"""
import os
import base64
import json
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


class GmailService:
    """Service class for interacting with Gmail API."""
    
    def __init__(self):
        self.service = None
        self.credentials = None
        
    def authenticate(self):
        """
        Authenticate and build Gmail service using OAuth 2.0.
        Returns True if successful, False otherwise.
        """
        creds = None
        
        # Load existing token if available
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
        self.service = build('gmail', 'v1', credentials=creds)
        return True
    
    def _get_unread_emails_internal(self, max_results=None):
        """Internal method that performs the actual API call."""
        query = config.GMAIL_QUERY
        
        # Add subject filter if configured
        if config.SUBJECT_FILTER:
            subject_filter_query = f'subject:"{config.SUBJECT_FILTER}"'
            # Combine with existing query
            query = f'{query} {subject_filter_query}'
        
        max_results = max_results or config.MAX_RESULTS
        
        results = self.service.users().messages().list(
            userId='me',
            q=query,
            maxResults=max_results
        ).execute()
        
        return results.get('messages', [])
    
    @retry(
        stop=stop_after_attempt(5),
        wait=wait_exponential(multiplier=1, min=2, max=60),
        retry=retry_if_exception(_is_retryable_http_error),
        reraise=True,
        before_sleep=lambda retry_state: logger.warning(f"Retrying get_unread_emails after {retry_state.outcome.exception()}")
    )
    def get_unread_emails(self, max_results=None):
        """
        Fetch unread emails from inbox.
        
        Args:
            max_results: Maximum number of emails to fetch
            
        Returns:
            List of email message objects
        """
        if not self.service:
            raise ValueError("Service not authenticated. Call authenticate() first.")
        
        try:
            return self._get_unread_emails_internal(max_results)
        except HttpError as error:
            # Non-retryable errors (401, 403, etc.) - return empty list
            if not _is_retryable_http_error(error):
                logger.error(f'Non-retryable error occurred: {error}')
                return []
            # Retryable errors that exhausted all retries
            logger.error(f'All retry attempts exhausted for get_unread_emails: {error}')
            return []
    
    def _get_email_details_internal(self, message_id):
        """Internal method that performs the actual API call."""
        message = self.service.users().messages().get(
            userId='me',
            id=message_id,
            format='full'
        ).execute()
        
        headers = message['payload'].get('headers', [])
        
        # Extract headers
        email_data = {}
        for header in headers:
            name = header['name'].lower()
            if name == 'from':
                email_data['from'] = header['value']
            elif name == 'subject':
                email_data['subject'] = header['value']
            elif name == 'date':
                email_data['date'] = header['value']
        
        # Extract body
        email_data['body'] = self._extract_body(message['payload'])
        email_data['id'] = message_id
        email_data['internal_date'] = message.get('internalDate', '')
        
        return email_data
    
    @retry(
        stop=stop_after_attempt(5),
        wait=wait_exponential(multiplier=1, min=2, max=60),
        retry=retry_if_exception(_is_retryable_http_error),

        #retry=retry_if(_is_retryable_http_error),
        reraise=True,
        before_sleep=lambda retry_state: logger.warning(f"Retrying get_email_details after {retry_state.outcome.exception()}")
    )
    def get_email_details(self, message_id):
        """
        Get detailed information about a specific email.
        
        Args:
            message_id: Gmail message ID
            
        Returns:
            Dictionary containing email details (from, subject, date, body)
        """
        if not self.service:
            raise ValueError("Service not authenticated. Call authenticate() first.")
        
        try:
            return self._get_email_details_internal(message_id)
        except HttpError as error:
            # Non-retryable errors (401, 403, etc.) - return None
            if not _is_retryable_http_error(error):
                logger.error(f'Non-retryable error occurred while fetching email {message_id}: {error}')
                return None
            # Retryable errors that exhausted all retries
            logger.error(f'All retry attempts exhausted for get_email_details: {error}')
            return None
    
    def _extract_body(self, payload):
        """
        Extract plain text body from email payload.
        
        Args:
            payload: Email payload object
            
        Returns:
            Plain text body content
        """
        body = ""
        
        if 'parts' in payload:
            # Multipart message
            for part in payload['parts']:
                if part['mimeType'] == 'text/plain':
                    data = part['body'].get('data')
                    if data:
                        body = base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
                        break
                elif part['mimeType'] == 'text/html' and not body:
                    # Fallback to HTML if no plain text
                    data = part['body'].get('data')
                    if data:
                        html_body = base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
                        # Simple HTML tag removal (basic conversion)
                        import re
                        body = re.sub('<[^<]+?>', '', html_body)
        else:
            # Single part message
            if payload['mimeType'] == 'text/plain':
                data = payload['body'].get('data')
                if data:
                    body = base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
            elif payload['mimeType'] == 'text/html':
                data = payload['body'].get('data')
                if data:
                    html_body = base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
                    # Simple HTML tag removal
                    import re
                    body = re.sub('<[^<]+?>', '', html_body)
        
        return body.strip()
    
    def _mark_as_read_internal(self, message_id):
        """Internal method that performs the actual API call."""
        self.service.users().messages().modify(
            userId='me',
            id=message_id,
            body={'removeLabelIds': ['UNREAD']}
        ).execute()
        return True
    
    @retry(
        stop=stop_after_attempt(5),
        wait=wait_exponential(multiplier=1, min=2, max=60),
        retry=retry_if_exception(_is_retryable_http_error),

        #retry=retry_if(_is_retryable_http_error),
        reraise=True,
        before_sleep=lambda retry_state: logger.warning(f"Retrying mark_as_read after {retry_state.outcome.exception()}")
    )
    def mark_as_read(self, message_id):
        """
        Mark an email as read by removing the UNREAD label.
        
        Args:
            message_id: Gmail message ID
            
        Returns:
            True if successful, False otherwise
        """
        if not self.service:
            raise ValueError("Service not authenticated. Call authenticate() first.")
        
        try:
            return self._mark_as_read_internal(message_id)
        except HttpError as error:
            # Non-retryable errors (401, 403, etc.) - return False
            if not _is_retryable_http_error(error):
                logger.error(f'Non-retryable error occurred while marking email {message_id} as read: {error}')
                return False
            # Retryable errors that exhausted all retries
            logger.error(f'All retry attempts exhausted for mark_as_read: {error}')
            return False
