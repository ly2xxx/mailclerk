import win32com.client as win32
from datetime import datetime, timedelta
import os
import re
import logging
from typing import List, Dict, Optional, Tuple
import json
import hashlib
from pathlib import Path, PurePosixPath

logger = logging.getLogger(__name__)

class SecurityError(Exception):
    """Custom exception for security-related errors"""
    pass

class OutlookEmailHandler:
    """Handle Outlook email operations using win32com.client with enhanced security"""
    
    # Security constants
    MAX_FILENAME_LENGTH = 255
    MAX_PATH_LENGTH = 260
    MAX_EMAIL_SIZE_MB = 100
    MAX_ATTACHMENT_SIZE_MB = 50
    
    # Allowed and blocked file extensions
    BLOCKED_EXTENSIONS = {
        '.exe', '.bat', '.cmd', '.com', '.scr', '.pif', '.vbs', '.vbe',
        '.js', '.jse', '.wsf', '.wsh', '.msi', '.msp', '.hta', '.jar',
        '.ps1', '.psm1', '.psd1', '.ps1xml', '.pssc', '.psrc', '.cdxml'
    }
    
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.inbox = None
        
    def connect(self) -> bool:
        """Connect to Outlook application with security checks"""
        try:
            # Basic security check - ensure we're on Windows
            import platform
            if platform.system() != 'Windows':
                logger.error("Outlook connection only supported on Windows")
                return False
                
            self.outlook = win32.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            
            # Verify we can access the default folder
            self.inbox = self.namespace.GetDefaultFolder(6)  # 6 = Inbox
            
            # Test connection by attempting to get folder name
            folder_name = self.inbox.Name
            logger.info(f"Successfully connected to Outlook folder: {folder_name}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to connect to Outlook: Secure connection error")
            # Don't log the actual exception details for security
            return False
    
    def _validate_date_range(self, start_date: datetime.date, end_date: datetime.date) -> bool:
        """Validate date range for security"""
        if not isinstance(start_date, datetime.date) or not isinstance(end_date, datetime.date):
            raise SecurityError("Invalid date format")
        
        if start_date > end_date:
            raise SecurityError("Start date cannot be after end date")
        
        # Prevent fetching emails from too far in the past (performance/security)
        max_days_back = 365  # 1 year
        if (datetime.now().date() - start_date).days > max_days_back:
            raise SecurityError(f"Date range cannot exceed {max_days_back} days")
        
        return True
    
    def _validate_rules(self, rules: Dict) -> Dict:
        """Validate and sanitize filtering rules"""
        if not rules:
            return {}
        
        validated_rules = {}
        
        for rule_type in ['sender_keywords', 'subject_keywords', 'high_priority_senders', 'exclude_keywords']:
            if rule_type in rules:
                keywords = rules[rule_type]
                if isinstance(keywords, list):
                    # Validate each keyword
                    validated_keywords = []
                    for keyword in keywords[:20]:  # Limit to 20 keywords per type
                        if isinstance(keyword, str) and len(keyword.strip()) <= 100:
                            # Basic sanitization - remove control characters
                            sanitized = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', keyword.strip())
                            if sanitized:
                                validated_keywords.append(sanitized.lower())
                    validated_rules[rule_type] = validated_keywords
        
        return validated_rules
    
    def fetch_emails(self, start_date: datetime.date, end_date: datetime.date, 
                    rules: Dict = None) -> List[Dict]:
        """Fetch emails based on date range and rules with security validation"""
        if not self.inbox:
            raise SecurityError("Not connected to Outlook")
        
        try:
            # Validate inputs
            self._validate_date_range(start_date, end_date)
            validated_rules = self._validate_rules(rules)
            
            # Convert dates to datetime for filtering
            start_datetime = datetime.combine(start_date, datetime.min.time())
            end_datetime = datetime.combine(end_date, datetime.max.time())
            
            # Get messages from inbox with security limits
            messages = self.inbox.Items
            messages.Sort("[ReceivedTime]", True)  # Sort by received time, descending
            
            # Filter by date - use secure date formatting
            filter_criteria = f"[ReceivedTime] >= '{start_datetime.strftime('%m/%d/%Y %H:%M %p')}' AND [ReceivedTime] <= '{end_datetime.strftime('%m/%d/%Y %H:%M %p')}'"
            filtered_messages = messages.Restrict(filter_criteria)
            
            emails = []
            max_emails = 1000  # Security limit
            
            for i, message in enumerate(filtered_messages):
                if i >= max_emails:
                    logger.warning(f"Reached maximum email limit ({max_emails})")
                    break
                    
                try:
                    email_data = self._extract_email_data(message, validated_rules)
                    if email_data and self._passes_rules(email_data, validated_rules):
                        emails.append(email_data)
                except Exception:
                    # Log without exposing details
                    logger.warning(f"Error processing email {i+1}")
                    continue
            
            logger.info(f"Successfully fetched {len(emails)} emails")
            return emails
            
        except SecurityError:
            raise
        except Exception as e:
            logger.error("Error fetching emails")
            raise SecurityError("Failed to fetch emails securely")
    
    def _sanitize_filename(self, filename: str) -> str:
        """Securely sanitize filename to prevent path traversal and invalid characters"""
        if not filename or not isinstance(filename, str):
            return "unnamed_file"
        
        # Remove path components and normalize
        filename = os.path.basename(filename)
        
        # Remove or replace dangerous characters
        # Keep only alphanumeric, spaces, dots, dashes, underscores
        sanitized = re.sub(r'[^\w\s.-]', '_', filename)
        
        # Remove multiple consecutive dots (prevent ../ attacks)
        sanitized = re.sub(r'\.{2,}', '.', sanitized)
        
        # Remove leading/trailing dots and spaces
        sanitized = sanitized.strip('. ')
        
        # Ensure filename is not empty and not too long
        if not sanitized:
            sanitized = "unnamed_file"
        
        if len(sanitized) > self.MAX_FILENAME_LENGTH:
            name, ext = os.path.splitext(sanitized)
            sanitized = name[:self.MAX_FILENAME_LENGTH-len(ext)-10] + "_truncated" + ext
        
        return sanitized
    
    def _validate_download_path(self, base_path: str) -> str:
        """Validate and secure the download path"""
        if not base_path or not isinstance(base_path, str):
            raise SecurityError("Invalid download path")
        
        try:
            # Normalize and resolve the path
            path = Path(base_path).resolve()
            
            # Ensure path length is reasonable
            if len(str(path)) > self.MAX_PATH_LENGTH:
                raise SecurityError("Download path too long")
            
            # Ensure the path is within expected boundaries (not root or system dirs)
            path_str = str(path).lower()
            forbidden_paths = ['c:\\windows', 'c:\\system32', 'c:\\program files']
            
            for forbidden in forbidden_paths:
                if path_str.startswith(forbidden):
                    raise SecurityError("Cannot download to system directories")
            
            # Create directory if it doesn't exist (with proper error handling)
            path.mkdir(parents=True, exist_ok=True)
            
            # Test write permissions
            test_file = path / "test_write_permission.tmp"
            try:
                test_file.touch()
                test_file.unlink()
            except Exception:
                raise SecurityError("No write permission to download path")
            
            return str(path)
            
        except SecurityError:
            raise
        except Exception:
            raise SecurityError("Invalid download path")
    
    def _validate_file_extension(self, filename: str) -> bool:
        """Validate file extension for security"""
        if not filename:
            return True
        
        ext = os.path.splitext(filename.lower())[1]
        return ext not in self.BLOCKED_EXTENSIONS
    
    def _extract_email_data(self, message, rules: Dict = None) -> Optional[Dict]:
        """Extract relevant data from an email message with security validation"""
        try:
            # Check email size
            email_size = getattr(message, 'Size', 0)
            if email_size > self.MAX_EMAIL_SIZE_MB * 1024 * 1024:
                logger.warning(f"Email too large: {email_size} bytes")
                return None
            
            # Safely extract basic information
            subject = getattr(message, 'Subject', 'No Subject') or 'No Subject'
            sender_name = getattr(message, 'SenderName', 'Unknown Sender') or 'Unknown Sender'
            sender_email = getattr(message, 'SenderEmailAddress', '') or ''
            
            # Sanitize extracted data
            subject = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', subject)[:200]
            sender_name = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', sender_name)[:100]
            sender_email = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', sender_email)[:100]
            
            email_data = {
                'subject': subject,
                'sender': sender_name,
                'sender_email': sender_email,
                'received_date': self._format_date(getattr(message, 'ReceivedTime', None)),
                'size': email_size,
                'has_attachments': self._check_safe_attachments(message),
                'message_id': hashlib.sha256(str(getattr(message, 'EntryID', '')).encode()).hexdigest()[:16],
                'conversation_id': hashlib.sha256(str(getattr(message, 'ConversationID', '')).encode()).hexdigest()[:16],
                'importance': min(max(getattr(message, 'Importance', 1), 0), 2),
                'preview': self._get_safe_email_preview(message),
                'category': re.sub(r'[\x00-\x1f\x7f-\x9f]', '', getattr(message, 'Categories', ''))[:50],
                'message_object': message
            }
            
            # Calculate chain timestamp
            email_data['chain_timestamp'] = self._get_chain_timestamp(message)
            
            # Determine if high priority based on rules
            email_data['is_high_priority'] = self._is_high_priority(email_data, rules)
            
            return email_data
            
        except Exception:
            logger.warning("Error extracting email data")
            return None
    
    def _check_safe_attachments(self, message) -> bool:
        """Check if email has safe attachments"""
        try:
            attachments = getattr(message, 'Attachments', None)
            if not attachments or len(attachments) == 0:
                return False
            
            # Check each attachment for safety
            for attachment in attachments:
                filename = getattr(attachment, 'FileName', '')
                if filename and not self._validate_file_extension(filename):
                    logger.warning(f"Blocked unsafe attachment: {filename}")
                    return False
                
                # Check attachment size
                size = getattr(attachment, 'Size', 0)
                if size > self.MAX_ATTACHMENT_SIZE_MB * 1024 * 1024:
                    logger.warning(f"Attachment too large: {size} bytes")
                    return False
            
            return True
            
        except Exception:
            return False
    
    def _get_safe_email_preview(self, message) -> str:
        """Get email body preview safely"""
        try:
            body = getattr(message, 'Body', '')
            if body:
                # Remove control characters and limit length
                clean_body = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', body)
                preview = re.sub(r'\s+', ' ', clean_body.strip())
                return preview[:200] if len(preview) > 200 else preview
            return 'No preview available'
        except Exception:
            return 'Preview unavailable'
    
    def _format_date(self, date_obj) -> str:
        """Format date object to string safely"""
        if date_obj:
            try:
                if hasattr(date_obj, 'strftime'):
                    return date_obj.strftime('%Y-%m-%d %H:%M:%S')
                else:
                    return str(date_obj)[:19]  # Limit length
            except Exception:
                pass
        return 'Unknown Date'
    
    def _get_chain_timestamp(self, message) -> str:
        """Get the latest timestamp from the email conversation chain"""
        try:
            received_time = getattr(message, 'ReceivedTime', None)
            if received_time:
                return self._format_date(received_time)
            return 'Unknown'
        except Exception:
            return 'Unknown'
    
    def _is_high_priority(self, email_data: Dict, rules: Dict = None) -> bool:
        """Determine if email is high priority based on rules"""
        try:
            if not rules:
                return email_data.get('importance', 1) == 2
            
            # Check if sender is in high priority list
            high_priority_senders = rules.get('high_priority_senders', [])
            sender_email = email_data.get('sender_email', '').lower()
            sender_name = email_data.get('sender', '').lower()
            
            for priority_sender in high_priority_senders:
                if priority_sender and (priority_sender in sender_email or priority_sender in sender_name):
                    return True
            
            # Check Outlook importance flag
            return email_data.get('importance', 1) == 2
            
        except Exception:
            return False
    
    def _passes_rules(self, email_data: Dict, rules: Dict = None) -> bool:
        """Check if email passes the filtering rules"""
        try:
            if not rules:
                return True
            
            subject = email_data.get('subject', '').lower()
            sender = email_data.get('sender', '').lower()
            
            # Check exclude keywords
            exclude_keywords = rules.get('exclude_keywords', [])
            for exclude_keyword in exclude_keywords:
                if exclude_keyword and (exclude_keyword in subject or exclude_keyword in sender):
                    return False
            
            # Check sender keywords (if specified, email must match at least one)
            sender_keywords = rules.get('sender_keywords', [])
            if sender_keywords:
                sender_match = any(keyword and keyword in sender for keyword in sender_keywords)
                if not sender_match:
                    return False
            
            # Check subject keywords (if specified, email must match at least one)
            subject_keywords = rules.get('subject_keywords', [])
            if subject_keywords:
                subject_match = any(keyword and keyword in subject for keyword in subject_keywords)
                if not subject_match:
                    return False
            
            return True
            
        except Exception:
            return False
    
    def download_email(self, email_data: Dict, download_path: str) -> bool:
        """Download email and its attachments to local storage with enhanced security"""
        try:
            # Validate inputs
            validated_path = self._validate_download_path(download_path)
            
            message = email_data.get('message_object')
            if not message:
                logger.error("No message object found for download")
                return False
            
            # Create safe filename
            subject = email_data.get('subject', 'No_Subject')
            safe_subject = self._sanitize_filename(subject)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            # Create email-specific folder with length limit
            folder_name = f"{timestamp}_{safe_subject[:50]}"
            email_folder = os.path.join(validated_path, folder_name)
            
            # Additional path validation
            if len(email_folder) > self.MAX_PATH_LENGTH:
                email_folder = os.path.join(validated_path, f"{timestamp}_email")
            
            os.makedirs(email_folder, exist_ok=True)
            
            # Save email metadata securely
            metadata = {
                'subject': email_data.get('subject', '')[:200],
                'sender': email_data.get('sender', '')[:100],
                'sender_email': email_data.get('sender_email', '')[:100],
                'received_date': email_data.get('received_date', ''),
                'size': email_data.get('size', 0),
                'has_attachments': email_data.get('has_attachments', False),
                'importance': email_data.get('importance', 1),
                'category': email_data.get('category', '')[:50],
                'chain_timestamp': email_data.get('chain_timestamp', ''),
                'download_timestamp': datetime.now().isoformat()
            }
            
            metadata_file = os.path.join(email_folder, 'metadata.json')
            with open(metadata_file, 'w', encoding='utf-8') as f:
                json.dump(metadata, f, indent=2, ensure_ascii=False)
            
            # Save email body securely
            self._save_email_body(message, email_folder)
            
            # Download attachments securely
            if email_data.get('has_attachments', False):
                self._download_attachments_safely(message, email_folder)
            
            logger.info(f"Successfully downloaded email to: {email_folder}")
            return True
            
        except SecurityError as e:
            logger.error(f"Security error during download: {str(e)}")
            return False
        except Exception:
            logger.error("Error downloading email")
            return False
    
    def _save_email_body(self, message, email_folder: str):
        """Save email body safely"""
        try:
            # Save plain text body
            body = getattr(message, 'Body', '')
            if body:
                # Clean and limit body content
                clean_body = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', body)
                if len(clean_body) > 1024 * 1024:  # 1MB limit
                    clean_body = clean_body[:1024*1024] + "\n[Content truncated for security]"
                
                body_file = os.path.join(email_folder, 'email_body.txt')
                with open(body_file, 'w', encoding='utf-8') as f:
                    f.write(clean_body)
        except Exception:
            logger.warning("Could not save email body")
        
        try:
            # Save HTML body if available
            html_body = getattr(message, 'HTMLBody', '')
            if html_body:
                # Clean HTML and limit size
                if len(html_body) > 2 * 1024 * 1024:  # 2MB limit
                    html_body = html_body[:2*1024*1024] + "\n<!-- Content truncated for security -->"
                
                html_file = os.path.join(email_folder, 'email_body.html')
                with open(html_file, 'w', encoding='utf-8') as f:
                    f.write(html_body)
        except Exception:
            logger.warning("Could not save HTML body")
    
    def _download_attachments_safely(self, message, email_folder: str):
        """Download attachments with security validation"""
        try:
            attachments_folder = os.path.join(email_folder, 'attachments')
            os.makedirs(attachments_folder, exist_ok=True)
            
            for i, attachment in enumerate(message.Attachments):
                try:
                    filename = getattr(attachment, 'FileName', f'attachment_{i}')
                    
                    # Validate file extension
                    if not self._validate_file_extension(filename):
                        logger.warning(f"Skipped unsafe attachment: {filename}")
                        continue
                    
                    # Validate attachment size
                    size = getattr(attachment, 'Size', 0)
                    if size > self.MAX_ATTACHMENT_SIZE_MB * 1024 * 1024:
                        logger.warning(f"Skipped large attachment: {filename} ({size} bytes)")
                        continue
                    
                    # Sanitize filename
                    safe_filename = self._sanitize_filename(filename)
                    if not safe_filename:
                        safe_filename = f"attachment_{i}"
                    
                    attachment_path = os.path.join(attachments_folder, safe_filename)
                    
                    # Ensure path is safe
                    if len(attachment_path) > self.MAX_PATH_LENGTH:
                        logger.warning(f"Attachment path too long: {filename}")
                        continue
                    
                    # Save attachment
                    attachment.SaveAsFile(attachment_path)
                    logger.info(f"Downloaded attachment: {safe_filename}")
                    
                except Exception:
                    logger.warning(f"Could not download attachment {i}")
                    
        except Exception:
            logger.warning("Error downloading attachments")
    
    def get_folder_list(self) -> List[str]:
        """Get list of available Outlook folders safely"""
        try:
            if not self.namespace:
                return []
            
            folders = []
            for folder in self.namespace.Folders:
                folder_name = getattr(folder, 'Name', 'Unknown')
                # Sanitize folder name
                safe_name = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', folder_name)[:100]
                if safe_name:
                    folders.append(safe_name)
            
            return folders[:20]  # Limit number of folders returned
            
        except Exception:
            logger.error("Error getting folder list")
            return []
    
    def close_connection(self):
        """Close the Outlook connection safely"""
        try:
            if self.outlook:
                # Clear references
                self.outlook = None
                self.namespace = None
                self.inbox = None
                logger.info("Outlook connection closed")
        except Exception:
            logger.error("Error closing Outlook connection")