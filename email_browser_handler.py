import requests
import json
import os
import re
import logging
import hashlib
from datetime import datetime, timedelta
from typing import List, Dict, Optional
from pathlib import Path
import base64
import webbrowser
from urllib.parse import urlparse, parse_qs

from email_handler_base import EmailHandlerBase

logger = logging.getLogger(__name__)

class BrowserEmailHandler(EmailHandlerBase):
    """Handle Outlook emails using Microsoft Graph API (browser-based authentication)"""
    
    # Microsoft Graph API endpoints
    GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
    AUTH_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
    TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    
    # Security constants
    MAX_FILENAME_LENGTH = 255
    MAX_PATH_LENGTH = 260
    MAX_EMAIL_SIZE_MB = 100
    MAX_ATTACHMENT_SIZE_MB = 50
    
    # Required scopes for email access
    SCOPES = [
        "https://graph.microsoft.com/Mail.Read",
        "https://graph.microsoft.com/Mail.ReadWrite", 
        "https://graph.microsoft.com/User.Read"
    ]
    
    def __init__(self, client_id: str = None):
        super().__init__()
        # You'll need to register an app in Azure AD to get this
        # For now, using a public client ID that works for personal accounts
        self.client_id = client_id or "04b07795-8ddb-461a-bbee-02f9e1bf7b46"  # Microsoft Graph Explorer public client
        self.access_token = None
        self.token_file = "outlook_token.json"
        
    def connect(self) -> bool:
        """Connect using OAuth 2.0 authentication"""
        try:
            # Try to load existing token
            if self._load_token():
                if self._validate_token():
                    logger.info("Using existing valid token")
                    self.connected = True
                    return True
                else:
                    logger.info("Existing token expired, refreshing...")
                    if self._refresh_token():
                        self.connected = True
                        return True
            
            # Need new authentication
            logger.info("Starting OAuth authentication flow...")
            if self._authenticate():
                self.connected = True
                return True
            
            return False
            
        except Exception as e:
            logger.error(f"Authentication failed: {type(e).__name__}")
            return False
    
    def _authenticate(self) -> bool:
        """Perform OAuth authentication"""
        try:
            # Generate authorization URL
            auth_params = {
                'client_id': self.client_id,
                'response_type': 'code',
                'redirect_uri': 'http://localhost:8080',
                'scope': ' '.join(self.SCOPES),
                'response_mode': 'query'
            }
            
            auth_url = self.AUTH_URL + '?' + '&'.join([f"{k}={v}" for k, v in auth_params.items()])
            
            print("\n" + "="*60)
            print("OUTLOOK AUTHENTICATION REQUIRED")
            print("="*60)
            print("1. A browser window will open")
            print("2. Sign in to your Outlook.com account")
            print("3. Copy the FULL URL from the browser after authorization")
            print("4. Paste it here")
            print("="*60)
            
            # Open browser
            webbrowser.open(auth_url)
            
            # Get authorization code from user
            redirect_url = input("\nPaste the full redirect URL here: ").strip()
            
            # Extract authorization code
            parsed_url = urlparse(redirect_url)
            query_params = parse_qs(parsed_url.query)
            
            if 'code' not in query_params:
                logger.error("No authorization code found in URL")
                return False
            
            auth_code = query_params['code'][0]
            
            # Exchange code for token
            token_data = {
                'client_id': self.client_id,
                'grant_type': 'authorization_code',
                'code': auth_code,
                'redirect_uri': 'http://localhost:8080',
                'scope': ' '.join(self.SCOPES)
            }
            
            response = requests.post(self.TOKEN_URL, data=token_data)
            
            if response.status_code == 200:
                token_info = response.json()
                self.access_token = token_info['access_token']
                
                # Save token info
                self._save_token(token_info)
                
                logger.info("✅ Authentication successful!")
                return True
            else:
                logger.error(f"Token exchange failed: {response.status_code}")
                logger.error(response.text)
                return False
                
        except Exception as e:
            logger.error(f"Authentication error: {e}")
            return False
    
    def _load_token(self) -> bool:
        """Load saved token"""
        try:
            if os.path.exists(self.token_file):
                with open(self.token_file, 'r') as f:
                    token_info = json.load(f)
                    self.access_token = token_info.get('access_token')
                    return self.access_token is not None
        except Exception as e:
            logger.warning(f"Could not load token: {e}")
        return False
    
    def _save_token(self, token_info: Dict):
        """Save token info"""
        try:
            # Add expiration time
            if 'expires_in' in token_info:
                expires_at = datetime.now() + timedelta(seconds=token_info['expires_in'])
                token_info['expires_at'] = expires_at.isoformat()
            
            with open(self.token_file, 'w') as f:
                json.dump(token_info, f, indent=2)
            
            # Set restrictive permissions
            os.chmod(self.token_file, 0o600)
            
        except Exception as e:
            logger.warning(f"Could not save token: {e}")
    
    def _validate_token(self) -> bool:
        """Validate current token"""
        if not self.access_token:
            return False
        
        try:
            # Test token by making a simple API call
            headers = {'Authorization': f'Bearer {self.access_token}'}
            response = requests.get(f"{self.GRAPH_BASE_URL}/me", headers=headers, timeout=10)
            
            if response.status_code == 200:
                user_info = response.json()
                logger.info(f"Authenticated as: {user_info.get('displayName', 'Unknown')}")
                return True
            else:
                logger.warning("Token validation failed")
                return False
                
        except Exception as e:
            logger.warning(f"Token validation error: {e}")
            return False
    
    def _refresh_token(self) -> bool:
        """Refresh access token"""
        try:
            if os.path.exists(self.token_file):
                with open(self.token_file, 'r') as f:
                    token_info = json.load(f)
                
                refresh_token = token_info.get('refresh_token')
                if not refresh_token:
                    return False
                
                # Refresh token request
                refresh_data = {
                    'client_id': self.client_id,
                    'grant_type': 'refresh_token',
                    'refresh_token': refresh_token,
                    'scope': ' '.join(self.SCOPES)
                }
                
                response = requests.post(self.TOKEN_URL, data=refresh_data)
                
                if response.status_code == 200:
                    new_token_info = response.json()
                    self.access_token = new_token_info['access_token']
                    self._save_token(new_token_info)
                    logger.info("Token refreshed successfully")
                    return True
                    
        except Exception as e:
            logger.warning(f"Token refresh failed: {e}")
        
        return False
    
    def fetch_emails(self, start_date: datetime.date, end_date: datetime.date, 
                    rules: Dict = None) -> List[Dict]:
        """Fetch emails using Microsoft Graph API"""
        if not self.connected or not self.access_token:
            raise ConnectionError("Not connected to Outlook")
        
        try:
            # Validate inputs
            self._validate_date_range(start_date, end_date)
            validated_rules = self._validate_rules(rules)
            
            # Build Graph API query
            start_datetime = datetime.combine(start_date, datetime.min.time())
            end_datetime = datetime.combine(end_date, datetime.max.time())
            
            # Format dates for Graph API (ISO 8601)
            start_iso = start_datetime.isoformat() + 'Z'
            end_iso = end_datetime.isoformat() + 'Z'
            
            # Build filter query
            filter_query = f"receivedDateTime ge {start_iso} and receivedDateTime le {end_iso}"
            
            # API parameters
            params = {
                '$filter': filter_query,
                '$orderby': 'receivedDateTime desc',
                '$top': 1000,  # Max emails per request
                '$select': 'id,subject,sender,receivedDateTime,bodyPreview,hasAttachments,importance,conversationId,body'
            }
            
            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Content-Type': 'application/json'
            }
            
            # Make API request
            response = requests.get(
                f"{self.GRAPH_BASE_URL}/me/messages",
                headers=headers,
                params=params,
                timeout=30
            )
            
            if response.status_code == 200:
                data = response.json()
                emails = []
                
                for message in data.get('value', []):
                    email_data = self._convert_graph_message(message, validated_rules)
                    if email_data and self._passes_rules(email_data, validated_rules):
                        emails.append(email_data)
                
                logger.info(f"Successfully fetched {len(emails)} emails")
                return emails
            
            elif response.status_code == 401:
                logger.error("Token expired, please reconnect")
                self.connected = False
                return []
            else:
                logger.error(f"API request failed: {response.status_code}")
                logger.error(response.text)
                return []
                
        except Exception as e:
            logger.error(f"Error fetching emails: {e}")
            return []
    
    def _convert_graph_message(self, message: Dict, rules: Dict = None) -> Optional[Dict]:
        """Convert Graph API message to our email data format"""
        try:
            # Extract sender info
            sender_info = message.get('sender', {})
            sender_email_addr = sender_info.get('emailAddress', {})
            sender_name = sender_email_addr.get('name', 'Unknown Sender')
            sender_email = sender_email_addr.get('address', '')
            
            # Parse received date
            received_str = message.get('receivedDateTime', '')
            try:
                received_dt = datetime.fromisoformat(received_str.replace('Z', '+00:00'))
                received_date = received_dt.strftime('%Y-%m-%d %H:%M:%S')
            except:
                received_date = received_str
            
            # Extract and sanitize data
            subject = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', message.get('subject', 'No Subject'))[:200]
            sender_name = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', sender_name)[:100]
            sender_email = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', sender_email)[:100]
            
            # Get body preview
            preview = message.get('bodyPreview', '')[:200]
            preview = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', preview)
            
            email_data = {
                'subject': subject,
                'sender': sender_name,
                'sender_email': sender_email,
                'received_date': received_date,
                'chain_timestamp': received_date,
                'size': len(str(message)),  # Approximate size
                'has_attachments': message.get('hasAttachments', False),
                'message_id': hashlib.sha256(message.get('id', '').encode()).hexdigest()[:16],
                'conversation_id': hashlib.sha256(message.get('conversationId', '').encode()).hexdigest()[:16],
                'importance': self._convert_importance(message.get('importance', 'normal')),
                'preview': preview,
                'category': '',  # Graph API doesn't provide categories in the same way
                'graph_message_id': message.get('id'),  # Store Graph ID for downloads
                'message_object': message  # Store full message for download
            }
            
            # Determine if high priority
            email_data['is_high_priority'] = self._is_high_priority(email_data, rules)
            
            return email_data
            
        except Exception as e:
            logger.warning(f"Error converting message: {e}")
            return None
    
    def _convert_importance(self, importance: str) -> int:
        """Convert Graph API importance to numeric value"""
        importance_map = {
            'low': 0,
            'normal': 1,
            'high': 2
        }
        return importance_map.get(importance.lower(), 1)
    
    def download_email(self, email_data: Dict, download_path: str) -> bool:
        """Download email and its attachments"""
        if not self.connected or not self.access_token:
            logger.error("Not connected to Outlook")
            return False
        
        try:
            # Validate download path
            validated_path = self._validate_download_path(download_path)
            
            graph_message_id = email_data.get('graph_message_id')
            if not graph_message_id:
                logger.error("No Graph message ID found for download")
                return False
            
            # Create safe filename
            subject = email_data.get('subject', 'No_Subject')
            safe_subject = self._sanitize_filename(subject)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            # Create email-specific folder
            folder_name = f"{timestamp}_{safe_subject[:50]}"
            email_folder = os.path.join(validated_path, folder_name)
            
            if len(email_folder) > self.MAX_PATH_LENGTH:
                email_folder = os.path.join(validated_path, f"{timestamp}_email")
            
            os.makedirs(email_folder, exist_ok=True)
            
            # Save email metadata
            metadata = {
                'subject': email_data.get('subject', '')[:200],
                'sender': email_data.get('sender', '')[:100],
                'sender_email': email_data.get('sender_email', '')[:100],
                'received_date': email_data.get('received_date', ''),
                'size': email_data.get('size', 0),
                'has_attachments': email_data.get('has_attachments', False),
                'importance': email_data.get('importance', 1),
                'chain_timestamp': email_data.get('chain_timestamp', ''),
                'download_timestamp': datetime.now().isoformat()
            }
            
            metadata_file = os.path.join(email_folder, 'metadata.json')
            with open(metadata_file, 'w', encoding='utf-8') as f:
                json.dump(metadata, f, indent=2, ensure_ascii=False)
            
            # Download full email content
            self._download_email_content(graph_message_id, email_folder)
            
            # Download attachments if present
            if email_data.get('has_attachments', False):
                self._download_attachments(graph_message_id, email_folder)
            
            logger.info(f"Successfully downloaded email to: {email_folder}")
            return True
            
        except Exception as e:
            logger.error(f"Error downloading email: {e}")
            return False
    
    def _download_email_content(self, message_id: str, email_folder: str):
        """Download email body content"""
        try:
            headers = {'Authorization': f'Bearer {self.access_token}'}
            
            # Get full message
            response = requests.get(
                f"{self.GRAPH_BASE_URL}/me/messages/{message_id}",
                headers=headers,
                params={'$select': 'body'},
                timeout=30
            )
            
            if response.status_code == 200:
                message = response.json()
                body_data = message.get('body', {})
                
                # Save HTML body
                html_content = body_data.get('content', '')
                if html_content:
                    html_file = os.path.join(email_folder, 'email_body.html')
                    with open(html_file, 'w', encoding='utf-8') as f:
                        f.write(html_content)
                
                # Convert to text and save
                text_content = self._html_to_text(html_content)
                if text_content:
                    text_file = os.path.join(email_folder, 'email_body.txt')
                    with open(text_file, 'w', encoding='utf-8') as f:
                        f.write(text_content)
                        
        except Exception as e:
            logger.warning(f"Could not download email content: {e}")
    
    def _download_attachments(self, message_id: str, email_folder: str):
        """Download email attachments"""
        try:
            headers = {'Authorization': f'Bearer {self.access_token}'}
            
            # Get attachments list
            response = requests.get(
                f"{self.GRAPH_BASE_URL}/me/messages/{message_id}/attachments",
                headers=headers,
                timeout=30
            )
            
            if response.status_code == 200:
                attachments_data = response.json()
                attachments = attachments_data.get('value', [])
                
                if attachments:
                    attachments_folder = os.path.join(email_folder, 'attachments')
                    os.makedirs(attachments_folder, exist_ok=True)
                    
                    for attachment in attachments:
                        self._download_single_attachment(attachment, attachments_folder)
                        
        except Exception as e:
            logger.warning(f"Could not download attachments: {e}")
    
    def _download_single_attachment(self, attachment: Dict, attachments_folder: str):
        """Download a single attachment"""
        try:
            filename = attachment.get('name', 'unnamed_attachment')
            
            # Validate file extension and size
            if not self._validate_file_extension(filename):
                logger.warning(f"Skipped unsafe attachment: {filename}")
                return
            
            size = attachment.get('size', 0)
            if size > self.MAX_ATTACHMENT_SIZE_MB * 1024 * 1024:
                logger.warning(f"Skipped large attachment: {filename} ({size} bytes)")
                return
            
            # Sanitize filename
            safe_filename = self._sanitize_filename(filename)
            attachment_path = os.path.join(attachments_folder, safe_filename)
            
            # Get attachment content
            content_bytes = attachment.get('contentBytes')
            if content_bytes:
                # Decode base64 content
                content = base64.b64decode(content_bytes)
                
                with open(attachment_path, 'wb') as f:
                    f.write(content)
                
                logger.info(f"Downloaded attachment: {safe_filename}")
                
        except Exception as e:
            logger.warning(f"Could not download attachment: {e}")
    
    def _html_to_text(self, html_content: str) -> str:
        """Convert HTML to plain text"""
        try:
            # Simple HTML to text conversion
            import re
            
            # Remove HTML tags
            text = re.sub(r'<[^>]+>', '', html_content)
            
            # Decode HTML entities
            text = text.replace('&nbsp;', ' ')
            text = text.replace('&amp;', '&')
            text = text.replace('&lt;', '<')
            text = text.replace('&gt;', '>')
            text = text.replace('&quot;', '"')
            
            # Clean up whitespace
            text = re.sub(r'\s+', ' ', text.strip())
            
            return text
            
        except Exception:
            return html_content
    
    def _validate_download_path(self, base_path: str) -> str:
        """Validate and secure the download path"""
        if not base_path or not isinstance(base_path, str):
            raise ValueError("Invalid download path")
        
        try:
            path = Path(base_path).resolve()
            
            if len(str(path)) > self.MAX_PATH_LENGTH:
                raise ValueError("Download path too long")
            
            path.mkdir(parents=True, exist_ok=True)
            
            return str(path)
            
        except Exception as e:
            raise ValueError(f"Invalid download path: {e}")
    
    def _sanitize_filename(self, filename: str) -> str:
        """Securely sanitize filename"""
        if not filename or not isinstance(filename, str):
            return "unnamed_file"
        
        filename = os.path.basename(filename)
        sanitized = re.sub(r'[^\w\s.-]', '_', filename)
        sanitized = re.sub(r'\.{2,}', '.', sanitized)
        sanitized = sanitized.strip('. ')
        
        if not sanitized:
            sanitized = "unnamed_file"
        
        if len(sanitized) > self.MAX_FILENAME_LENGTH:
            name, ext = os.path.splitext(sanitized)
            sanitized = name[:self.MAX_FILENAME_LENGTH-len(ext)-10] + "_truncated" + ext
        
        return sanitized
    
    def _validate_file_extension(self, filename: str) -> bool:
        """Validate file extension for security"""
        if not filename:
            return True
        
        blocked_extensions = {
            '.exe', '.bat', '.cmd', '.com', '.scr', '.pif', '.vbs', '.vbe',
            '.js', '.jse', '.wsf', '.wsh', '.msi', '.msp', '.hta', '.jar',
            '.ps1', '.psm1', '.psd1', '.ps1xml', '.pssc', '.psrc', '.cdxml'
        }
        
        ext = os.path.splitext(filename.lower())[1]
        return ext not in blocked_extensions
    
    def get_folder_list(self) -> List[str]:
        """Get list of available email folders"""
        if not self.connected or not self.access_token:
            return []
        
        try:
            headers = {'Authorization': f'Bearer {self.access_token}'}
            
            response = requests.get(
                f"{self.GRAPH_BASE_URL}/me/mailFolders",
                headers=headers,
                timeout=30
            )
            
            if response.status_code == 200:
                data = response.json()
                folders = []
                
                for folder in data.get('value', []):
                    folder_name = folder.get('displayName', 'Unknown')
                    folders.append(folder_name)
                
                return folders[:20]  # Limit number of folders
            
        except Exception as e:
            logger.error(f"Error getting folder list: {e}")
        
        return []
    
    def close_connection(self):
        """Close the connection"""
        self.connected = False
        self.access_token = None
        logger.info("Browser email connection closed")