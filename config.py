import json
import os
from typing import Dict, List
import logging

logger = logging.getLogger(__name__)

class EmailConfig:
    """Configuration management for email sorting rules and settings"""
    
    CONFIG_FILE = 'mailclerk_config.json'
    
    @staticmethod
    def get_default_rules() -> Dict:
        """Get default sorting rules"""
        return {
            'sender_keywords': [
                'support',
                'notification',
                'admin',
                'team'
            ],
            'subject_keywords': [
                'urgent',
                'important',
                'action required',
                'deadline',
                'meeting',
                'project'
            ],
            'high_priority_senders': [
                'manager',
                'boss',
                'ceo',
                'director',
                'urgent'
            ],
            'exclude_keywords': [
                'spam',
                'newsletter',
                'unsubscribe',
                'marketing',
                'advertisement',
                'promotional'
            ]
        }
    
    @staticmethod
    def load_rules() -> Dict:
        """Load sorting rules from config file"""
        try:
            if os.path.exists(EmailConfig.CONFIG_FILE):
                with open(EmailConfig.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    return config.get('sorting_rules', EmailConfig.get_default_rules())
            else:
                logger.info("No config file found, using default rules")
                return EmailConfig.get_default_rules()
        except Exception as e:
            logger.error(f"Error loading config: {str(e)}")
            return EmailConfig.get_default_rules()
    
    @staticmethod
    def save_rules(rules: Dict) -> bool:
        """Save sorting rules to config file"""
        try:
            config = {}
            
            # Load existing config if it exists
            if os.path.exists(EmailConfig.CONFIG_FILE):
                try:
                    with open(EmailConfig.CONFIG_FILE, 'r', encoding='utf-8') as f:
                        config = json.load(f)
                except:
                    config = {}
            
            # Update sorting rules
            config['sorting_rules'] = rules
            config['last_updated'] = str(datetime.now())
            
            # Save to file
            with open(EmailConfig.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            
            logger.info("Successfully saved sorting rules")
            return True
            
        except Exception as e:
            logger.error(f"Error saving config: {str(e)}")
            return False
    
    @staticmethod
    def get_download_settings() -> Dict:
        """Get download settings from config"""
        try:
            if os.path.exists(EmailConfig.CONFIG_FILE):
                with open(EmailConfig.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    return config.get('download_settings', EmailConfig.get_default_download_settings())
            else:
                return EmailConfig.get_default_download_settings()
        except Exception as e:
            logger.error(f"Error loading download settings: {str(e)}")
            return EmailConfig.get_default_download_settings()
    
    @staticmethod
    def get_default_download_settings() -> Dict:
        """Get default download settings"""
        return {
            'download_path': os.path.expanduser("~/Downloads/MailClerk"),
            'create_subfolders': True,
            'include_attachments': True,
            'save_html_body': True,
            'save_text_body': True,
            'save_metadata': True,
            'max_email_size_mb': 50,  # Maximum email size to download
            'allowed_attachment_types': [  # Empty list means all types allowed
                # '.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx',
                # '.txt', '.jpg', '.jpeg', '.png', '.gif', '.zip'
            ],
            'exclude_attachment_types': [
                '.exe', '.bat', '.cmd', '.scr', '.com', '.pif', '.vbs'
            ]
        }
    
    @staticmethod
    def save_download_settings(settings: Dict) -> bool:
        """Save download settings to config file"""
        try:
            config = {}
            
            # Load existing config if it exists
            if os.path.exists(EmailConfig.CONFIG_FILE):
                try:
                    with open(EmailConfig.CONFIG_FILE, 'r', encoding='utf-8') as f:
                        config = json.load(f)
                except:
                    config = {}
            
            # Update download settings
            config['download_settings'] = settings
            config['last_updated'] = str(datetime.now())
            
            # Save to file
            with open(EmailConfig.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            
            logger.info("Successfully saved download settings")
            return True
            
        except Exception as e:
            logger.error(f"Error saving download settings: {str(e)}")
            return False
    
    @staticmethod
    def get_outlook_settings() -> Dict:
        """Get Outlook connection settings"""
        return {
            'connection_timeout': 30,  # seconds
            'retry_attempts': 3,
            'default_folder': 'Inbox',
            'max_emails_per_fetch': 1000,
            'enable_conversation_grouping': True
        }
    
    @staticmethod
    def validate_rules(rules: Dict) -> List[str]:
        """Validate sorting rules and return list of errors"""
        errors = []
        
        required_keys = ['sender_keywords', 'subject_keywords', 'high_priority_senders', 'exclude_keywords']
        
        for key in required_keys:
            if key not in rules:
                errors.append(f"Missing required key: {key}")
            elif not isinstance(rules[key], list):
                errors.append(f"Key '{key}' must be a list")
        
        return errors
    
    @staticmethod
    def export_config(filepath: str) -> bool:
        """Export current configuration to a file"""
        try:
            config = {}
            
            if os.path.exists(EmailConfig.CONFIG_FILE):
                with open(EmailConfig.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            
            logger.info(f"Configuration exported to: {filepath}")
            return True
            
        except Exception as e:
            logger.error(f"Error exporting config: {str(e)}")
            return False
    
    @staticmethod
    def import_config(filepath: str) -> bool:
        """Import configuration from a file"""
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # Validate the imported config
            if 'sorting_rules' in config:
                errors = EmailConfig.validate_rules(config['sorting_rules'])
                if errors:
                    logger.error(f"Invalid configuration: {errors}")
                    return False
            
            # Save the imported config
            with open(EmailConfig.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            
            logger.info(f"Configuration imported from: {filepath}")
            return True
            
        except Exception as e:
            logger.error(f"Error importing config: {str(e)}")
            return False

# Import datetime here to avoid circular import
from datetime import datetime