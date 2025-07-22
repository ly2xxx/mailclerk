from abc import ABC, abstractmethod
from typing import List, Dict, Optional
from datetime import datetime

class EmailHandlerBase(ABC):
    """Abstract base class for email handlers"""
    
    def __init__(self):
        self.connected = False
    
    @abstractmethod
    def connect(self) -> bool:
        """Connect to email service"""
        pass
    
    @abstractmethod
    def fetch_emails(self, start_date: datetime.date, end_date: datetime.date, 
                    rules: Dict = None) -> List[Dict]:
        """Fetch emails based on date range and rules"""
        pass
    
    @abstractmethod
    def download_email(self, email_data: Dict, download_path: str) -> bool:
        """Download email and its attachments to local storage"""
        pass
    
    @abstractmethod
    def get_folder_list(self) -> List[str]:
        """Get list of available email folders"""
        pass
    
    @abstractmethod
    def close_connection(self):
        """Close the email service connection"""
        pass
    
    # Common helper methods that can be shared
    def _validate_date_range(self, start_date: datetime.date, end_date: datetime.date) -> bool:
        """Validate date range for security"""
        if not isinstance(start_date, datetime.date) or not isinstance(end_date, datetime.date):
            raise ValueError("Invalid date format")
        
        if start_date > end_date:
            raise ValueError("Start date cannot be after end date")
        
        # Prevent fetching emails from too far in the past (performance/security)
        max_days_back = 365  # 1 year
        if (datetime.now().date() - start_date).days > max_days_back:
            raise ValueError(f"Date range cannot exceed {max_days_back} days")
        
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
                            import re
                            sanitized = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', keyword.strip())
                            if sanitized:
                                validated_keywords.append(sanitized.lower())
                    validated_rules[rule_type] = validated_keywords
        
        return validated_rules
    
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
            
            # Check importance flag
            return email_data.get('importance', 1) == 2
            
        except Exception:
            return False