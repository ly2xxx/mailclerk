import win32com.client as win32
from datetime import datetime, timedelta
import os
import re
import logging
from typing import List, Dict, Optional, Tuple
import json

logger = logging.getLogger(__name__)

class OutlookEmailHandler:
    """Handle Outlook email operations using win32com.client"""
    
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.inbox = None
        
    def connect(self) -> bool:
        """Connect to Outlook application"""
        try:
            self.outlook = win32.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.inbox = self.namespace.GetDefaultFolder(6)  # 6 = Inbox
            logger.info("Successfully connected to Outlook")
            return True
        except Exception as e:
            logger.error(f"Failed to connect to Outlook: {str(e)}")
            return False
    
    def fetch_emails(self, start_date: datetime.date, end_date: datetime.date, 
                    rules: Dict = None) -> List[Dict]:
        """Fetch emails based on date range and rules"""
        if not self.inbox:
            raise Exception("Not connected to Outlook")
        
        try:
            # Convert dates to datetime for filtering
            start_datetime = datetime.combine(start_date, datetime.min.time())
            end_datetime = datetime.combine(end_date, datetime.max.time())
            
            # Get messages from inbox
            messages = self.inbox.Items
            messages.Sort("[ReceivedTime]", True)  # Sort by received time, descending
            
            # Filter by date
            filter_criteria = f"[ReceivedTime] >= '{start_datetime.strftime('%m/%d/%Y %H:%M %p')}' AND [ReceivedTime] <= '{end_datetime.strftime('%m/%d/%Y %H:%M %p')}'"
            filtered_messages = messages.Restrict(filter_criteria)
            
            emails = []
            for message in filtered_messages:
                try:
                    email_data = self._extract_email_data(message, rules)
                    if self._passes_rules(email_data, rules):
                        emails.append(email_data)
                except Exception as e:
                    logger.warning(f"Error processing email: {str(e)}")
                    continue
            
            logger.info(f"Fetched {len(emails)} emails")
            return emails
            
        except Exception as e:
            logger.error(f"Error fetching emails: {str(e)}")
            raise
    
    def _extract_email_data(self, message, rules: Dict = None) -> Dict:
        """Extract relevant data from an email message"""
        try:
            # Basic email information
            email_data = {
                'subject': getattr(message, 'Subject', 'No Subject'),
                'sender': getattr(message, 'SenderName', 'Unknown Sender'),
                'sender_email': getattr(message, 'SenderEmailAddress', ''),
                'received_date': self._format_date(getattr(message, 'ReceivedTime', None)),
                'size': getattr(message, 'Size', 0),
                'has_attachments': getattr(message, 'Attachments', None) and len(message.Attachments) > 0,
                'message_id': getattr(message, 'EntryID', ''),
                'conversation_id': getattr(message, 'ConversationID', ''),
                'importance': getattr(message, 'Importance', 1),  # 0=Low, 1=Normal, 2=High
                'preview': self._get_email_preview(message),
                'category': getattr(message, 'Categories', ''),
                'message_object': message  # Store reference for downloading
            }
            
            # Calculate chain timestamp (latest email in conversation)
            email_data['chain_timestamp'] = self._get_chain_timestamp(message)
            
            # Determine if high priority based on rules
            email_data['is_high_priority'] = self._is_high_priority(email_data, rules)
            
            return email_data
            
        except Exception as e:
            logger.error(f"Error extracting email data: {str(e)}")
            raise
    
    def _format_date(self, date_obj) -> str:
        """Format date object to string"""
        if date_obj:
            try:
                if hasattr(date_obj, 'strftime'):
                    return date_obj.strftime('%Y-%m-%d %H:%M:%S')
                else:
                    return str(date_obj)
            except:
                pass
        return 'Unknown Date'
    
    def _get_email_preview(self, message) -> str:
        """Get email body preview (first 200 characters)"""
        try:
            body = getattr(message, 'Body', '')
            if body:
                # Clean up the preview text
                preview = re.sub(r'\s+', ' ', body.strip())
                return preview[:200] if len(preview) > 200 else preview
            return 'No preview available'
        except:
            return 'Preview unavailable'
    
    def _get_chain_timestamp(self, message) -> str:
        """Get the latest timestamp from the email conversation chain"""
        try:
            # For now, return the received time of this message
            # In a more advanced implementation, you could fetch all emails
            # in the conversation and find the latest one
            received_time = getattr(message, 'ReceivedTime', None)
            if received_time:
                return self._format_date(received_time)
            return 'Unknown'
        except:
            return 'Unknown'
    
    def _is_high_priority(self, email_data: Dict, rules: Dict = None) -> bool:
        """Determine if email is high priority based on rules"""
        if not rules:
            return email_data['importance'] == 2  # Outlook high importance
        
        # Check if sender is in high priority list
        high_priority_senders = rules.get('high_priority_senders', [])
        sender_email = email_data.get('sender_email', '').lower()
        sender_name = email_data.get('sender', '').lower()
        
        for priority_sender in high_priority_senders:
            if priority_sender.lower() in sender_email or priority_sender.lower() in sender_name:
                return True
        
        # Check Outlook importance flag
        return email_data['importance'] == 2
    
    def _passes_rules(self, email_data: Dict, rules: Dict = None) -> bool:
        """Check if email passes the filtering rules"""
        if not rules:
            return True
        
        # Check exclude keywords
        exclude_keywords = rules.get('exclude_keywords', [])
        subject = email_data.get('subject', '').lower()
        sender = email_data.get('sender', '').lower()
        
        for exclude_keyword in exclude_keywords:
            if exclude_keyword.lower() in subject or exclude_keyword.lower() in sender:
                return False
        
        # Check sender keywords (if specified, email must match at least one)
        sender_keywords = rules.get('sender_keywords', [])
        if sender_keywords:
            sender_match = any(keyword.lower() in sender for keyword in sender_keywords)
            if not sender_match:
                return False
        
        # Check subject keywords (if specified, email must match at least one)
        subject_keywords = rules.get('subject_keywords', [])
        if subject_keywords:
            subject_match = any(keyword.lower() in subject for keyword in subject_keywords)
            if not subject_match:
                return False
        
        return True
    
    def download_email(self, email_data: Dict, download_path: str) -> bool:
        """Download email and its attachments to local storage"""
        try:
            message = email_data.get('message_object')
            if not message:
                logger.error("No message object found for download")
                return False
            
            # Create safe filename
            safe_subject = re.sub(r'[<>:"/\\|?*]', '_', email_data.get('subject', 'No_Subject'))
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            # Create email-specific folder
            email_folder = os.path.join(download_path, f"{timestamp}_{safe_subject[:50]}")
            os.makedirs(email_folder, exist_ok=True)
            
            # Save email metadata
            metadata = {
                'subject': email_data.get('subject', ''),
                'sender': email_data.get('sender', ''),
                'sender_email': email_data.get('sender_email', ''),
                'received_date': email_data.get('received_date', ''),
                'size': email_data.get('size', 0),
                'has_attachments': email_data.get('has_attachments', False),
                'importance': email_data.get('importance', 1),
                'category': email_data.get('category', ''),
                'chain_timestamp': email_data.get('chain_timestamp', '')
            }
            
            metadata_file = os.path.join(email_folder, 'metadata.json')
            with open(metadata_file, 'w', encoding='utf-8') as f:
                json.dump(metadata, f, indent=2, ensure_ascii=False)
            
            # Save email body
            try:
                body = getattr(message, 'Body', '')
                if body:
                    body_file = os.path.join(email_folder, 'email_body.txt')
                    with open(body_file, 'w', encoding='utf-8') as f:
                        f.write(body)
            except Exception as e:
                logger.warning(f"Could not save email body: {str(e)}")
            
            # Save HTML body if available
            try:
                html_body = getattr(message, 'HTMLBody', '')
                if html_body:
                    html_file = os.path.join(email_folder, 'email_body.html')
                    with open(html_file, 'w', encoding='utf-8') as f:
                        f.write(html_body)
            except Exception as e:
                logger.warning(f"Could not save HTML body: {str(e)}")
            
            # Download attachments
            if email_data.get('has_attachments', False):
                attachments_folder = os.path.join(email_folder, 'attachments')
                os.makedirs(attachments_folder, exist_ok=True)
                
                try:
                    for i, attachment in enumerate(message.Attachments):
                        attachment_name = getattr(attachment, 'FileName', f'attachment_{i}')
                        # Clean filename
                        safe_attachment_name = re.sub(r'[<>:"/\\|?*]', '_', attachment_name)
                        attachment_path = os.path.join(attachments_folder, safe_attachment_name)
                        
                        try:
                            attachment.SaveAsFile(attachment_path)
                            logger.info(f"Downloaded attachment: {attachment_name}")
                        except Exception as e:
                            logger.warning(f"Could not download attachment {attachment_name}: {str(e)}")
                            
                except Exception as e:
                    logger.warning(f"Error downloading attachments: {str(e)}")
            
            logger.info(f"Successfully downloaded email to: {email_folder}")
            return True
            
        except Exception as e:
            logger.error(f"Error downloading email: {str(e)}")
            return False
    
    def get_folder_list(self) -> List[str]:
        """Get list of available Outlook folders"""
        try:
            if not self.namespace:
                return []
            
            folders = []
            for folder in self.namespace.Folders:
                folders.append(folder.Name)
                # You could recursively get subfolders here if needed
            
            return folders
        except Exception as e:
            logger.error(f"Error getting folder list: {str(e)}")
            return []
    
    def close_connection(self):
        """Close the Outlook connection"""
        try:
            if self.outlook:
                # Note: Usually we don't quit Outlook as user might be using it
                # self.outlook.Quit()
                self.outlook = None
                self.namespace = None
                self.inbox = None
                logger.info("Outlook connection closed")
        except Exception as e:
            logger.error(f"Error closing Outlook connection: {str(e)}")