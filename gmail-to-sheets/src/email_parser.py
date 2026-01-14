"""
Email parser for extracting and formatting email data.
"""
from email.utils import parsedate_to_datetime
from datetime import datetime


class EmailParser:
    """Parser for email data extraction and formatting."""
    
    @staticmethod
    def parse_email(email_data):
        """
        Parse email data into structured format for Google Sheets.
        
        Args:
            email_data: Dictionary containing raw email data from Gmail API
            
        Returns:
            Dictionary with keys: from, subject, date, content
        """
        parsed = {
            'from': email_data.get('from', ''),
            'subject': email_data.get('subject', ''),
            'date': EmailParser._format_date(email_data.get('date', '')),
            'content': email_data.get('body', ''),
            'id': email_data.get('id', ''),
            'internal_date': email_data.get('internal_date', '')
        }
        
        return parsed
    
    @staticmethod
    def _format_date(date_string):
        """
        Format date string to a readable format.
        
        Args:
            date_string: RFC 2822 date string
            
        Returns:
            Formatted date string (YYYY-MM-DD HH:MM:SS)
        """
        if not date_string:
            return ''
        
        try:
            # Parse RFC 2822 date
            dt = parsedate_to_datetime(date_string)
            return dt.strftime('%Y-%m-%d %H:%M:%S')
        except (ValueError, TypeError):
            # Fallback to original string if parsing fails
            return date_string
    
    @staticmethod
    def extract_sender_email(from_field):
        """
        Extract email address from 'From' field.
        
        Args:
            from_field: Full 'From' field (e.g., "John Doe <john@example.com>")
            
        Returns:
            Email address only
        """
        if not from_field:
            return ''
        
        # Try to extract email from angle brackets
        import re
        match = re.search(r'<(.+?)>', from_field)
        if match:
            return match.group(1)
        
        # If no angle brackets, assume the whole string is the email
        return from_field.strip()
