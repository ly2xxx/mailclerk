import json
import os
from typing import Dict, List
import logging
from pathlib import Path
import re
import stat

logger = logging.getLogger(__name__)

class ConfigSecurityError(Exception):
    """Custom exception for configuration security errors"""
    pass

class EmailConfig:
    """Configuration management for email sorting rules and settings with enhanced security"""
    
    CONFIG_FILE = 'mailclerk_config.json'
    MAX_CONFIG_SIZE = 1024 * 1024  # 1MB max config file size
    MAX_KEYWORD_LENGTH = 100
    MAX_KEYWORDS_PER_TYPE = 50
    MAX_PATH_LENGTH = 260
    
    @staticmethod
    def _validate_path_security(path: str) -> str:
        """Validate path for security issues"""
        if not path or not isinstance(path, str):
            raise ConfigSecurityError("Invalid path provided")
        
        # Normalize path and check length
        try:
            normalized_path = os.path.normpath(path)
            if len(normalized_path) > EmailConfig.MAX_PATH_LENGTH:
                raise ConfigSecurityError("Path too long")
            
            # Check for path traversal attempts
            if '..' in normalized_path or normalized_path.startswith('/'):
                if not os.path.isabs(normalized_path):
                    raise ConfigSecurityError("Path traversal detected")
            
            # Check for dangerous paths
            dangerous_paths = [
                'c:\\windows', 'c:\\system32', 'c:\\program files',
                '/etc', '/bin', '/sbin', '/usr/bin', '/sys', '/proc'
            ]
            
            path_lower = normalized_path.lower()
            for dangerous in dangerous_paths:
                if path_lower.startswith(dangerous):
                    raise ConfigSecurityError("Cannot access system directories")
            
            return normalized_path
            
        except ConfigSecurityError:
            raise
        except Exception:
            raise ConfigSecurityError("Invalid path format")
    
    @staticmethod
    def _sanitize_keyword(keyword: str) -> str:
        """Sanitize and validate keyword input"""
        if not isinstance(keyword, str):
            return ""
        
        # Remove control characters and limit length
        sanitized = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', keyword.strip())
        
        if len(sanitized) > EmailConfig.MAX_KEYWORD_LENGTH:
            sanitized = sanitized[:EmailConfig.MAX_KEYWORD_LENGTH]
        
        return sanitized
    
    @staticmethod
    def _validate_keywords_list(keywords: List[str]) -> List[str]:
        """Validate and sanitize a list of keywords"""
        if not isinstance(keywords, list):
            return []
        
        validated = []
        for keyword in keywords[:EmailConfig.MAX_KEYWORDS_PER_TYPE]:
            sanitized = EmailConfig._sanitize_keyword(keyword)
            if sanitized and len(sanitized) >= 2:  # Minimum length check
                validated.append(sanitized.lower())
        
        return validated
    
    @staticmethod
    def _secure_file_operation(filepath: str, mode: str, operation_func):
        """Perform file operations securely"""
        try:
            # Validate file path
            safe_path = EmailConfig._validate_path_security(filepath)
            
            # Check if we're trying to read and file doesn't exist
            if 'r' in mode and not os.path.exists(safe_path):
                return None
            
            # Check file size for reading operations
            if 'r' in mode and os.path.exists(safe_path):
                file_size = os.path.getsize(safe_path)
                if file_size > EmailConfig.MAX_CONFIG_SIZE:
                    raise ConfigSecurityError("Configuration file too large")
            
            # Perform operation with proper error handling
            try:
                with open(safe_path, mode, encoding='utf-8') as f:
                    return operation_func(f)
            except PermissionError:
                raise ConfigSecurityError("No permission to access configuration file")
            except json.JSONDecodeError:
                raise ConfigSecurityError("Invalid JSON in configuration file")
                
        except ConfigSecurityError:
            raise
        except Exception as e:
            raise ConfigSecurityError(f"File operation failed: {type(e).__name__}")
    
    @staticmethod
    def get_default_rules() -> Dict:
        """Get default sorting rules with validation"""
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
                'supervisor',
                'director'
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
        """Load sorting rules from config file securely"""
        try:
            def read_operation(f):
                config = json.load(f)
                rules = config.get('sorting_rules', {})
                return EmailConfig._validate_rules_structure(rules)
            
            rules = EmailConfig._secure_file_operation(
                EmailConfig.CONFIG_FILE, 
                'r', 
                read_operation
            )
            
            if rules is None:
                logger.info("No config file found, using default rules")
                return EmailConfig.get_default_rules()
            
            return rules
            
        except ConfigSecurityError as e:
            logger.error(f"Security error loading config: {str(e)}")
            return EmailConfig.get_default_rules()
        except Exception:
            logger.error("Error loading configuration")
            return EmailConfig.get_default_rules()
    
    @staticmethod
    def _validate_rules_structure(rules: Dict) -> Dict:
        """Validate the structure and content of rules"""
        if not isinstance(rules, dict):
            return EmailConfig.get_default_rules()
        
        validated_rules = {}
        default_rules = EmailConfig.get_default_rules()
        
        for rule_type in ['sender_keywords', 'subject_keywords', 'high_priority_senders', 'exclude_keywords']:
            if rule_type in rules:
                validated_rules[rule_type] = EmailConfig._validate_keywords_list(rules[rule_type])
            else:
                validated_rules[rule_type] = default_rules[rule_type]
        
        return validated_rules
    
    @staticmethod
    def save_rules(rules: Dict) -> bool:
        """Save sorting rules to config file securely"""
        try:
            # Validate rules before saving
            validated_rules = EmailConfig._validate_rules_structure(rules)
            
            def write_operation(f):
                # Load existing config safely
                config = {}
                if os.path.exists(EmailConfig.CONFIG_FILE):
                    try:
                        f.seek(0)
                        config = json.load(f) if f.read() else {}
                        f.seek(0)
                        f.truncate()
                    except json.JSONDecodeError:
                        config = {}
                
                # Update with validated rules
                config['sorting_rules'] = validated_rules
                config['last_updated'] = datetime.now().isoformat()
                config['version'] = '1.0'
                
                # Write atomically
                json.dump(config, f, indent=2, ensure_ascii=False)
                f.flush()
                os.fsync(f.fileno())
                
                return True
            
            # First, read existing config if it exists
            existing_config = {}
            if os.path.exists(EmailConfig.CONFIG_FILE):
                def read_operation(f):
                    return json.load(f)
                
                try:
                    existing_config = EmailConfig._secure_file_operation(
                        EmailConfig.CONFIG_FILE, 'r', read_operation
                    ) or {}
                except Exception:
                    existing_config = {}
            
            # Now write the updated config
            def write_operation(f):
                existing_config['sorting_rules'] = validated_rules
                existing_config['last_updated'] = datetime.now().isoformat()
                existing_config['version'] = '1.0'
                
                json.dump(existing_config, f, indent=2, ensure_ascii=False)
                f.flush()
                os.fsync(f.fileno())
                return True
            
            result = EmailConfig._secure_file_operation(
                EmailConfig.CONFIG_FILE, 'w', write_operation
            )
            
            if result:
                # Set appropriate file permissions (owner read/write only)
                try:
                    os.chmod(EmailConfig.CONFIG_FILE, stat.S_IRUSR | stat.S_IWUSR)
                except Exception:
                    pass  # Permission setting is best effort
                
                logger.info("Successfully saved sorting rules")
                return True
            
            return False
            
        except ConfigSecurityError as e:
            logger.error(f"Security error saving config: {str(e)}")
            return False
        except Exception:
            logger.error("Error saving configuration")
            return False
    
    @staticmethod
    def get_download_settings() -> Dict:
        """Get download settings from config securely"""
        try:
            def read_operation(f):
                config = json.load(f)
                settings = config.get('download_settings', {})
                return EmailConfig._validate_download_settings(settings)
            
            settings = EmailConfig._secure_file_operation(
                EmailConfig.CONFIG_FILE, 
                'r', 
                read_operation
            )
            
            if settings is None:
                return EmailConfig.get_default_download_settings()
            
            return settings
            
        except Exception:
            logger.error("Error loading download settings")
            return EmailConfig.get_default_download_settings()
    
    @staticmethod
    def _validate_download_settings(settings: Dict) -> Dict:
        """Validate download settings structure and values"""
        if not isinstance(settings, dict):
            return EmailConfig.get_default_download_settings()
        
        default_settings = EmailConfig.get_default_download_settings()
        validated = {}
        
        # Validate download path
        if 'download_path' in settings:
            try:
                validated['download_path'] = EmailConfig._validate_path_security(
                    settings['download_path']
                )
            except ConfigSecurityError:
                validated['download_path'] = default_settings['download_path']
        else:
            validated['download_path'] = default_settings['download_path']
        
        # Validate boolean settings
        bool_settings = ['create_subfolders', 'include_attachments', 'save_html_body', 
                        'save_text_body', 'save_metadata']
        for setting in bool_settings:
            validated[setting] = bool(settings.get(setting, default_settings[setting]))
        
        # Validate numeric settings with bounds
        max_size = min(max(int(settings.get('max_email_size_mb', default_settings['max_email_size_mb'])), 1), 500)
        validated['max_email_size_mb'] = max_size
        
        # Validate file extension lists
        if 'allowed_attachment_types' in settings:
            validated['allowed_attachment_types'] = EmailConfig._validate_extensions_list(
                settings['allowed_attachment_types']
            )
        else:
            validated['allowed_attachment_types'] = default_settings['allowed_attachment_types']
        
        if 'exclude_attachment_types' in settings:
            validated['exclude_attachment_types'] = EmailConfig._validate_extensions_list(
                settings['exclude_attachment_types']
            )
        else:
            validated['exclude_attachment_types'] = default_settings['exclude_attachment_types']
        
        return validated
    
    @staticmethod
    def _validate_extensions_list(extensions: List[str]) -> List[str]:
        """Validate file extensions list"""
        if not isinstance(extensions, list):
            return []
        
        validated = []
        for ext in extensions[:50]:  # Limit number of extensions
            if isinstance(ext, str):
                clean_ext = re.sub(r'[^a-zA-Z0-9.]', '', ext.lower().strip())
                if clean_ext.startswith('.') and len(clean_ext) <= 10:
                    validated.append(clean_ext)
        
        return validated
    
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
            'max_email_size_mb': 50,
            'allowed_attachment_types': [],  # Empty list means all allowed (except excluded)
            'exclude_attachment_types': [
                '.exe', '.bat', '.cmd', '.scr', '.com', '.pif', '.vbs'
            ]
        }
    
    @staticmethod
    def save_download_settings(settings: Dict) -> bool:
        """Save download settings to config file securely"""
        try:
            validated_settings = EmailConfig._validate_download_settings(settings)
            
            # Load existing config
            existing_config = {}
            if os.path.exists(EmailConfig.CONFIG_FILE):
                def read_operation(f):
                    return json.load(f)
                
                try:
                    existing_config = EmailConfig._secure_file_operation(
                        EmailConfig.CONFIG_FILE, 'r', read_operation
                    ) or {}
                except Exception:
                    existing_config = {}
            
            # Update with validated settings
            def write_operation(f):
                existing_config['download_settings'] = validated_settings
                existing_config['last_updated'] = datetime.now().isoformat()
                
                json.dump(existing_config, f, indent=2, ensure_ascii=False)
                f.flush()
                os.fsync(f.fileno())
                return True
            
            result = EmailConfig._secure_file_operation(
                EmailConfig.CONFIG_FILE, 'w', write_operation
            )
            
            if result:
                logger.info("Successfully saved download settings")
                return True
            
            return False
            
        except Exception:
            logger.error("Error saving download settings")
            return False
    
    @staticmethod
    def get_outlook_settings() -> Dict:
        """Get Outlook connection settings with security defaults"""
        return {
            'connection_timeout': 30,
            'retry_attempts': 3,
            'default_folder': 'Inbox',
            'max_emails_per_fetch': 1000,
            'enable_conversation_grouping': True
        }
    
    @staticmethod
    def validate_rules(rules: Dict) -> List[str]:
        """Validate sorting rules and return list of errors"""
        errors = []
        
        if not isinstance(rules, dict):
            errors.append("Rules must be a dictionary")
            return errors
        
        required_keys = ['sender_keywords', 'subject_keywords', 'high_priority_senders', 'exclude_keywords']
        
        for key in required_keys:
            if key not in rules:
                errors.append(f"Missing required key: {key}")
            elif not isinstance(rules[key], list):
                errors.append(f"Key '{key}' must be a list")
            elif len(rules[key]) > EmailConfig.MAX_KEYWORDS_PER_TYPE:
                errors.append(f"Too many keywords in '{key}' (max: {EmailConfig.MAX_KEYWORDS_PER_TYPE})")
            else:
                # Validate individual keywords
                for i, keyword in enumerate(rules[key]):
                    if not isinstance(keyword, str):
                        errors.append(f"Keyword {i} in '{key}' must be a string")
                    elif len(keyword) > EmailConfig.MAX_KEYWORD_LENGTH:
                        errors.append(f"Keyword {i} in '{key}' too long (max: {EmailConfig.MAX_KEYWORD_LENGTH})")
        
        return errors
    
    @staticmethod
    def export_config(filepath: str) -> bool:
        """Export current configuration to a file securely"""
        try:
            # Validate export path
            safe_path = EmailConfig._validate_path_security(filepath)
            
            # Read current config
            def read_operation(f):
                return json.load(f)
            
            config = {}
            if os.path.exists(EmailConfig.CONFIG_FILE):
                config = EmailConfig._secure_file_operation(
                    EmailConfig.CONFIG_FILE, 'r', read_operation
                ) or {}
            
            # Write to export file
            def write_operation(f):
                # Add export metadata
                export_data = config.copy()
                export_data['export_timestamp'] = datetime.now().isoformat()
                export_data['export_version'] = '1.0'
                
                json.dump(export_data, f, indent=2, ensure_ascii=False)
                f.flush()
                os.fsync(f.fileno())
                return True
            
            result = EmailConfig._secure_file_operation(safe_path, 'w', write_operation)
            
            if result:
                logger.info(f"Configuration exported to: {safe_path}")
                return True
            
            return False
            
        except ConfigSecurityError as e:
            logger.error(f"Security error exporting config: {str(e)}")
            return False
        except Exception:
            logger.error("Error exporting configuration")
            return False
    
    @staticmethod
    def import_config(filepath: str) -> bool:
        """Import configuration from a file securely"""
        try:
            # Validate import path
            safe_path = EmailConfig._validate_path_security(filepath)
            
            # Read import file
            def read_operation(f):
                config = json.load(f)
                
                # Validate imported config structure
                if not isinstance(config, dict):
                    raise ConfigSecurityError("Invalid config format")
                
                # Validate sorting rules if present
                if 'sorting_rules' in config:
                    errors = EmailConfig.validate_rules(config['sorting_rules'])
                    if errors:
                        raise ConfigSecurityError(f"Invalid rules: {', '.join(errors)}")
                
                # Validate download settings if present
                if 'download_settings' in config:
                    config['download_settings'] = EmailConfig._validate_download_settings(
                        config['download_settings']
                    )
                
                return config
            
            imported_config = EmailConfig._secure_file_operation(
                safe_path, 'r', read_operation
            )
            
            if imported_config is None:
                return False
            
            # Save the imported config
            def write_operation(f):
                imported_config['import_timestamp'] = datetime.now().isoformat()
                json.dump(imported_config, f, indent=2, ensure_ascii=False)
                f.flush()
                os.fsync(f.fileno())
                return True
            
            result = EmailConfig._secure_file_operation(
                EmailConfig.CONFIG_FILE, 'w', write_operation
            )
            
            if result:
                logger.info(f"Configuration imported from: {safe_path}")
                return True
            
            return False
            
        except ConfigSecurityError as e:
            logger.error(f"Security error importing config: {str(e)}")
            return False
        except Exception:
            logger.error("Error importing configuration")
            return False

# Import datetime here to avoid circular import
from datetime import datetime