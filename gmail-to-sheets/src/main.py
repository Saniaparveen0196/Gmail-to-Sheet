"""
Main script for Gmail to Google Sheets automation.
"""
import os
import sys
import json
import logging
from datetime import datetime

# Add project root to path
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

import config
from src.gmail_service import GmailService
from src.sheets_service import SheetsService
from src.email_parser import EmailParser


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('gmail_to_sheets.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class StateManager:
    """Manages state to prevent duplicate processing."""
    
    def __init__(self, state_file):
        self.state_file = state_file
        self.state = self._load_state()
    
    def _load_state(self):
        """Load state from file."""
        if os.path.exists(self.state_file):
            try:
                with open(self.state_file, 'r') as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError) as e:
                logger.warning(f"Error loading state file: {e}. Starting fresh.")
                return {'processed_ids': [], 'last_run': None}
        return {'processed_ids': [], 'last_run': None}
    
    def save_state(self):
        """Save state to file."""
        try:
            os.makedirs(os.path.dirname(self.state_file), exist_ok=True)
            with open(self.state_file, 'w') as f:
                json.dump(self.state, f, indent=2)
            logger.info(f"State saved to {self.state_file}")
        except IOError as e:
            logger.error(f"Error saving state file: {e}")
    
    def is_processed(self, email_id):
        """Check if email has been processed."""
        return email_id in self.state.get('processed_ids', [])
    
    def mark_processed(self, email_id):
        """Mark email as processed."""
        if 'processed_ids' not in self.state:
            self.state['processed_ids'] = []
        self.state['processed_ids'].append(email_id)
        # Keep only last 1000 IDs to prevent file from growing too large
        if len(self.state['processed_ids']) > 1000:
            self.state['processed_ids'] = self.state['processed_ids'][-1000:]
    
    def update_last_run(self):
        """Update last run timestamp."""
        self.state['last_run'] = datetime.now().isoformat()


def main():
    """Main execution function."""
    logger.info("=" * 60)
    logger.info("Gmail to Google Sheets Automation - Starting")
    logger.info("=" * 60)
    
    # Validate configuration
    if not config.SPREADSHEET_ID:
        logger.error("SPREADSHEET_ID not set. Please set it in config.py or as environment variable.")
        return
    
    # Initialize state manager
    state_manager = StateManager(config.STATE_FILE)
    logger.info(f"Loaded state: {len(state_manager.state.get('processed_ids', []))} processed emails")
    
    # Initialize services
    gmail_service = GmailService()
    sheets_service = SheetsService()
    
    try:
        # Authenticate services
        logger.info("Authenticating Gmail service...")
        gmail_service.authenticate()
        logger.info("Gmail authentication successful")
        
        logger.info("Authenticating Sheets service...")
        sheets_service.authenticate()
        logger.info("Sheets authentication successful")
        
        # Ensure sheet headers exist
        logger.info(f"Ensuring headers exist in sheet '{config.SHEET_NAME}'...")
        sheets_service.ensure_headers(config.SPREADSHEET_ID, config.SHEET_NAME)
        
        # Fetch unread emails
        if config.SUBJECT_FILTER:
            logger.info(f"Fetching unread emails with subject filter: '{config.SUBJECT_FILTER}'...")
        else:
            logger.info("Fetching unread emails...")
        messages = gmail_service.get_unread_emails()
        logger.info(f"Found {len(messages)} unread email(s)")
        
        if not messages:
            logger.info("No new emails to process.")
            state_manager.update_last_run()
            state_manager.save_state()
            return
        
        # Process emails
        processed_count = 0
        skipped_count = 0
        
        for message in messages:
            message_id = message['id']
            
            # Check if already processed
            if state_manager.is_processed(message_id):
                logger.debug(f"Email {message_id} already processed, skipping...")
                skipped_count += 1
                continue
            
            # Get email details
            logger.info(f"Processing email {message_id}...")
            email_data = gmail_service.get_email_details(message_id)
            
            if not email_data:
                logger.warning(f"Failed to fetch details for email {message_id}")
                continue
            
            # Parse email
            parsed_email = EmailParser.parse_email(email_data)
            logger.debug(f"Parsed email keys: {list(parsed_email.keys())}")
            logger.debug(f"Parsed email from: {parsed_email.get('from', 'N/A')}, subject: {parsed_email.get('subject', 'N/A')[:50]}")
            
            # Append to sheet
            logger.info(f"Appending email to sheet '{config.SHEET_NAME}'...")
            success = sheets_service.append_email(
                config.SPREADSHEET_ID,
                config.SHEET_NAME,
                parsed_email
            )
            
            if success:
                # Mark as read
                gmail_service.mark_as_read(message_id)
                
                # Mark as processed in state
                state_manager.mark_processed(message_id)
                
                logger.info(f"Successfully processed email from: {parsed_email.get('from', 'Unknown')}")
                processed_count += 1
            else:
                logger.error(f"Failed to append email {message_id} to sheet")
                logger.error(f"Email data was: from={parsed_email.get('from')}, subject={parsed_email.get('subject')}")
        
        # Update and save state
        state_manager.update_last_run()
        state_manager.save_state()
        
        # Summary
        logger.info("=" * 60)
        logger.info(f"Processing complete!")
        logger.info(f"  - Processed: {processed_count} email(s)")
        logger.info(f"  - Skipped (duplicates): {skipped_count} email(s)")
        logger.info("=" * 60)
        
    except FileNotFoundError as e:
        logger.error(f"Configuration error: {e}")
        logger.error("Please ensure credentials.json is in the credentials/ directory")
    except Exception as e:
        logger.error(f"An error occurred: {e}", exc_info=True)
        raise


if __name__ == '__main__':
    main()
