import pytest
import os
import tempfile
import json
from datetime import datetime, timedelta
from unittest.mock import Mock, patch, MagicMock

from email_handler import OutlookEmailHandler, SecurityError
from config import EmailConfig, ConfigSecurityError
from app import (
    init_session_state, sanitize_user_input, validate_download_path,
    get_date_range
)

class TestEndToEndWorkflow:
    """Test complete end-to-end workflows"""
    
    @pytest.fixture
    def mock_streamlit_session(self):
        """Mock Streamlit session state for testing"""
        session_mock = {}
        with patch('streamlit.session_state', session_mock):
            yield session_mock
    
    @pytest.fixture
    def mock_outlook_system(self):
        """Mock complete Outlook system for integration testing"""
        with patch('email_handler.win32') as mock_win32:
            # Mock Outlook application
            mock_outlook = Mock()
            mock_namespace = Mock()
            mock_inbox = Mock()
            mock_inbox.Name = "Inbox"
            
            # Mock messages
            mock_messages = Mock()
            mock_message1 = Mock()
            mock_message1.Subject = "Important Meeting"
            mock_message1.SenderName = "Boss"
            mock_message1.SenderEmailAddress = "boss@company.com"
            mock_message1.ReceivedTime = datetime.now()
            mock_message1.Size = 2048
            mock_message1.Attachments = []
            mock_message1.EntryID = "msg1_id"
            mock_message1.ConversationID = "conv1_id"
            mock_message1.Importance = 2  # High importance
            mock_message1.Categories = "Work"
            mock_message1.Body = "Please review the quarterly reports."
            
            mock_message2 = Mock()
            mock_message2.Subject = "Spam Newsletter"
            mock_message2.SenderName = "Marketing"
            mock_message2.SenderEmailAddress = "marketing@spam.com"
            mock_message2.ReceivedTime = datetime.now()
            mock_message2.Size = 1024
            mock_message2.Attachments = []
            mock_message2.EntryID = "msg2_id"
            mock_message2.ConversationID = "conv2_id"
            mock_message2.Importance = 1
            mock_message2.Categories = ""
            mock_message2.Body = "Buy our products now!"
            
            mock_messages.Sort = Mock()
            mock_messages.Restrict = Mock(return_value=[mock_message1, mock_message2])
            mock_inbox.Items = mock_messages
            
            # Wire up the mocks
            mock_win32.Dispatch.return_value = mock_outlook
            mock_outlook.GetNamespace.return_value = mock_namespace
            mock_namespace.GetDefaultFolder.return_value = mock_inbox
            
            yield {
                'outlook': mock_outlook,
                'namespace': mock_namespace,
                'inbox': mock_inbox,
                'messages': [mock_message1, mock_message2]
            }
    
    def test_complete_email_workflow(self, mock_outlook_system, mock_streamlit_session):
        """Test complete email processing workflow"""
        # Step 1: Initialize session state
        with patch('app.EmailConfig') as mock_config:
            mock_config.load_rules.return_value = {
                'sender_keywords': ['boss'],
                'subject_keywords': ['important'],
                'high_priority_senders': ['boss@company.com'],
                'exclude_keywords': ['spam', 'marketing']
            }
            
            init_session_state()
            
            # Verify session state initialization
            assert 'email_handler' in mock_streamlit_session
            assert 'sorting_rules' in mock_streamlit_session
        
        # Step 2: Connect to Outlook
        with patch('platform.system', return_value='Windows'):
            handler = OutlookEmailHandler()
            connected = handler.connect()
            assert connected is True
        
        # Step 3: Fetch emails with filtering
        start_date = datetime.now().date()
        end_date = start_date
        
        emails = handler.fetch_emails(
            start_date=start_date,
            end_date=end_date,
            rules=mock_streamlit_session.get('sorting_rules', {})
        )
        
        # Should filter out spam but include important email
        assert len(emails) == 1  # Only the important email should pass filters
        assert emails[0]['subject'] == "Important Meeting"
        assert emails[0]['is_high_priority'] is True
        
        # Step 4: Download selected email
        with tempfile.TemporaryDirectory() as temp_dir:
            download_success = handler.download_email(emails[0], temp_dir)
            assert download_success is True
            
            # Verify email folder was created
            email_folders = [d for d in os.listdir(temp_dir) if os.path.isdir(os.path.join(temp_dir, d))]
            assert len(email_folders) == 1
            
            # Verify metadata file exists
            email_folder = os.path.join(temp_dir, email_folders[0])
            metadata_file = os.path.join(email_folder, 'metadata.json')
            assert os.path.exists(metadata_file)
            
            # Verify metadata content
            with open(metadata_file, 'r') as f:
                metadata = json.load(f)
                assert metadata['subject'] == "Important Meeting"
                assert metadata['sender'] == "Boss"
    
    def test_configuration_workflow(self):
        """Test complete configuration management workflow"""
        # Step 1: Load default rules
        default_rules = EmailConfig.get_default_rules()
        assert 'sender_keywords' in default_rules
        
        # Step 2: Modify rules
        new_rules = {
            'sender_keywords': ['test_sender'],
            'subject_keywords': ['test_subject'],
            'high_priority_senders': ['test_priority'],
            'exclude_keywords': ['test_exclude']
        }
        
        # Step 3: Validate rules
        validation_errors = EmailConfig.validate_rules(new_rules)
        assert validation_errors == []  # Should be valid
        
        # Step 4: Save rules
        with tempfile.NamedTemporaryFile(suffix='.json', delete=False) as temp_file:
            try:
                with patch.object(EmailConfig, 'CONFIG_FILE', temp_file.name):
                    save_success = EmailConfig.save_rules(new_rules)
                    assert save_success is True
                    
                    # Step 5: Load saved rules
                    loaded_rules = EmailConfig.load_rules()
                    assert loaded_rules['sender_keywords'] == ['test_sender']
                    assert loaded_rules['subject_keywords'] == ['test_subject']
            finally:
                os.unlink(temp_file.name)
    
    def test_error_recovery_workflow(self, mock_outlook_system):
        """Test error recovery scenarios"""
        handler = OutlookEmailHandler()
        
        # Test 1: Connection failure recovery
        with patch('email_handler.win32.Dispatch', side_effect=Exception("Connection failed")):
            connected = handler.connect()
            assert connected is False
            
            # Handler should remain in safe state
            assert handler.outlook is None
            assert handler.namespace is None
            assert handler.inbox is None
        
        # Test 2: Email fetch with invalid date range
        handler.inbox = mock_outlook_system['inbox']
        
        invalid_start = datetime.now().date()
        invalid_end = invalid_start - timedelta(days=1)  # End before start
        
        with pytest.raises(SecurityError):
            handler.fetch_emails(invalid_start, invalid_end)
        
        # Test 3: Download to invalid path
        email_data = {
            'subject': 'Test',
            'message_object': Mock()
        }
        
        download_success = handler.download_email(email_data, "../../../invalid/path")
        assert download_success is False
        
        # Test 4: Configuration error recovery
        with patch.object(EmailConfig, '_secure_file_operation', side_effect=ConfigSecurityError("Config error")):
            # Should fall back to defaults
            rules = EmailConfig.load_rules()
            default_rules = EmailConfig.get_default_rules()
            assert rules == default_rules
    
    def test_security_integration(self, mock_outlook_system):
        """Test security measures across integrated components"""
        handler = OutlookEmailHandler()
        handler.inbox = mock_outlook_system['inbox']
        
        # Test 1: Malicious rules injection
        malicious_rules = {
            'sender_keywords': ['../../../etc/passwd', '<script>alert("xss")</script>'],
            'subject_keywords': [''; DROP TABLE emails; --'],
            'high_priority_senders': ['$(rm -rf /)'],
            'exclude_keywords': ['\x00\x1f\x7f']
        }
        
        # Rules should be sanitized during validation
        validated_rules = EmailConfig._validate_rules_structure(malicious_rules)
        
        # Verify dangerous content is removed/sanitized
        all_keywords = []
        for keyword_list in validated_rules.values():
            all_keywords.extend(keyword_list)
        
        for keyword in all_keywords:
            assert '../' not in keyword
            assert '<script>' not in keyword
            assert 'DROP TABLE' not in keyword.upper()
            assert '$(rm' not in keyword
            assert '\x00' not in keyword
        
        # Test 2: Path traversal in download
        email_data = {
            'subject': '../../../evil',
            'sender': 'test@example.com',
            'message_object': Mock()
        }
        
        with tempfile.TemporaryDirectory() as temp_dir:
            # Download should sanitize the path
            download_success = handler.download_email(email_data, temp_dir)
            
            if download_success:
                # Verify no files were created outside temp_dir
                created_files = []
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        created_files.append(os.path.join(root, file))
                
                # All files should be within temp_dir
                for file_path in created_files:
                    assert temp_dir in os.path.abspath(file_path)
    
    def test_performance_integration(self, mock_outlook_system):
        """Test performance with realistic data volumes"""
        handler = OutlookEmailHandler()
        handler.inbox = mock_outlook_system['inbox']
        
        # Mock large number of emails
        large_message_list = []
        for i in range(500):  # Realistic inbox size
            mock_message = Mock()
            mock_message.Subject = f"Email {i}"
            mock_message.SenderName = f"Sender {i}"
            mock_message.SenderEmailAddress = f"sender{i}@example.com"
            mock_message.ReceivedTime = datetime.now() - timedelta(hours=i)
            mock_message.Size = 1024 + i
            mock_message.Attachments = []
            mock_message.EntryID = f"msg{i}_id"
            mock_message.ConversationID = f"conv{i}_id"
            mock_message.Importance = 1
            mock_message.Categories = ""
            mock_message.Body = f"Email body content {i}"
            large_message_list.append(mock_message)
        
        mock_outlook_system['inbox'].Items.Restrict.return_value = large_message_list
        
        # Fetch should handle large volumes efficiently
        start_date = datetime.now().date() - timedelta(days=1)
        end_date = datetime.now().date()
        
        import time
        start_time = time.time()
        
        emails = handler.fetch_emails(start_date, end_date, {})
        
        end_time = time.time()
        processing_time = end_time - start_time
        
        # Should complete within reasonable time and respect limits
        assert processing_time < 10.0  # Should process within 10 seconds
        assert len(emails) <= 1000  # Should respect max email limit
    
    def test_concurrent_safety(self):
        """Test thread safety and concurrent access"""
        # Create multiple handlers to simulate concurrent usage
        handlers = [OutlookEmailHandler() for _ in range(3)]
        
        # Each should maintain independent state
        for i, handler in enumerate(handlers):
            assert handler.outlook is None
            assert handler.namespace is None
            assert handler.inbox is None
        
        # Mock connection for one handler
        with patch('email_handler.win32') as mock_win32:
            mock_outlook = Mock()
            mock_win32.Dispatch.return_value = mock_outlook
            
            handlers[0].outlook = mock_outlook
            
            # Other handlers should not be affected
            assert handlers[1].outlook is None
            assert handlers[2].outlook is None
    
    def test_memory_management(self, mock_outlook_system):
        """Test memory management and resource cleanup"""
        handler = OutlookEmailHandler()
        
        # Test 1: Connection cleanup
        handler.outlook = Mock()
        handler.namespace = Mock()
        handler.inbox = Mock()
        
        handler.close_connection()
        
        # All references should be cleared
        assert handler.outlook is None
        assert handler.namespace is None
        assert handler.inbox is None
        
        # Test 2: Large data handling
        handler.inbox = mock_outlook_system['inbox']
        
        # Create emails with large content
        large_message = Mock()
        large_message.Subject = "Large Email"
        large_message.SenderName = "Test Sender"
        large_message.SenderEmailAddress = "test@example.com"
        large_message.ReceivedTime = datetime.now()
        large_message.Size = 50 * 1024 * 1024  # 50MB
        large_message.Attachments = []
        large_message.EntryID = "large_msg_id"
        large_message.ConversationID = "large_conv_id"
        large_message.Importance = 1
        large_message.Categories = ""
        large_message.Body = "A" * (10 * 1024 * 1024)  # 10MB body
        
        # Should handle large emails safely
        email_data = handler._extract_email_data(large_message)
        # Should reject emails that are too large
        assert email_data is None or email_data.get('size', 0) <= handler.MAX_EMAIL_SIZE_MB * 1024 * 1024

class TestDataValidation:
    """Test data validation across components"""
    
    def test_email_data_consistency(self):
        """Test consistency of email data across processing stages"""
        handler = OutlookEmailHandler()
        
        # Mock message with all fields
        mock_message = Mock()
        mock_message.Subject = "Test Subject"
        mock_message.SenderName = "Test Sender"
        mock_message.SenderEmailAddress = "test@example.com"
        mock_message.ReceivedTime = datetime.now()
        mock_message.Size = 1024
        mock_message.Attachments = []
        mock_message.EntryID = "test_id"
        mock_message.ConversationID = "test_conv"
        mock_message.Importance = 1
        mock_message.Categories = "Work"
        mock_message.Body = "Test body content"
        
        # Extract data
        email_data = handler._extract_email_data(mock_message)
        
        # Verify all required fields are present
        required_fields = [
            'subject', 'sender', 'sender_email', 'received_date',
            'size', 'has_attachments', 'message_id', 'conversation_id',
            'importance', 'preview', 'category', 'message_object',
            'chain_timestamp', 'is_high_priority'
        ]
        
        for field in required_fields:
            assert field in email_data
        
        # Verify data types
        assert isinstance(email_data['subject'], str)
        assert isinstance(email_data['sender'], str)
        assert isinstance(email_data['size'], int)
        assert isinstance(email_data['has_attachments'], bool)
        assert isinstance(email_data['is_high_priority'], bool)
    
    def test_configuration_data_integrity(self):
        """Test configuration data integrity across save/load cycles"""
        original_rules = {
            'sender_keywords': ['sender1', 'sender2'],
            'subject_keywords': ['urgent', 'important'],
            'high_priority_senders': ['boss@company.com'],
            'exclude_keywords': ['spam', 'marketing']
        }
        
        with tempfile.NamedTemporaryFile(suffix='.json', delete=False) as temp_file:
            try:
                with patch.object(EmailConfig, 'CONFIG_FILE', temp_file.name):
                    # Save rules
                    save_success = EmailConfig.save_rules(original_rules)
                    assert save_success is True
                    
                    # Load rules
                    loaded_rules = EmailConfig.load_rules()
                    
                    # Verify integrity
                    assert loaded_rules == original_rules
                    
                    # Verify structure
                    for key in original_rules:
                        assert key in loaded_rules
                        assert isinstance(loaded_rules[key], list)
                        assert loaded_rules[key] == original_rules[key]
            finally:
                os.unlink(temp_file.name)
    
    def test_cross_component_validation(self):
        """Test validation consistency across different components"""
        # Test keyword validation consistency
        test_keywords = [
            "valid_keyword",
            "a",  # Too short
            "a" * 200,  # Too long
            "keyword\x00with\x1fcontrol\x7fchars",  # Control characters
            123,  # Wrong type
            None,  # Null
            "",  # Empty
        ]
        
        # Validate through EmailConfig
        config_validated = EmailConfig._validate_keywords_list(test_keywords)
        
        # Validate through OutlookEmailHandler
        handler = OutlookEmailHandler()
        handler_rules = {
            'sender_keywords': test_keywords,
            'subject_keywords': [],
            'high_priority_senders': [],
            'exclude_keywords': []
        }
        handler_validated = handler._validate_rules(handler_rules)
        
        # Results should be consistent
        assert config_validated == handler_validated['sender_keywords']
        
        # Both should only contain valid keywords
        for keyword in config_validated:
            assert isinstance(keyword, str)
            assert len(keyword) >= 2
            assert len(keyword) <= EmailConfig.MAX_KEYWORD_LENGTH
            assert all(ord(c) >= 32 for c in keyword)  # No control characters

class TestRegressionTests:
    """Test for regression issues and edge cases"""
    
    def test_empty_email_handling(self):
        """Test handling of emails with missing or empty fields"""
        handler = OutlookEmailHandler()
        
        # Mock message with minimal/missing fields
        mock_message = Mock()
        mock_message.Size = 1024
        # Don't set other attributes to test defaults
        
        email_data = handler._extract_email_data(mock_message)
        
        # Should handle missing fields gracefully
        assert email_data is not None
        assert email_data['subject'] == "No Subject"
        assert email_data['sender'] == "Unknown Sender"
        assert email_data['sender_email'] == ""
    
    def test_unicode_handling(self):
        """Test proper handling of unicode content"""
        handler = OutlookEmailHandler()
        
        # Test unicode in various fields
        unicode_content = {
            'subject': "测试邮件 📧 Subject",
            'sender': "José María García",
            'body': "Здравствуй мир! 🌍"
        }
        
        mock_message = Mock()
        mock_message.Subject = unicode_content['subject']
        mock_message.SenderName = unicode_content['sender']
        mock_message.Body = unicode_content['body']
        mock_message.Size = 1024
        mock_message.Attachments = []
        mock_message.EntryID = "unicode_test"
        mock_message.ConversationID = "unicode_conv"
        mock_message.Importance = 1
        mock_message.Categories = ""
        mock_message.ReceivedTime = datetime.now()
        
        email_data = handler._extract_email_data(mock_message)
        
        # Should preserve unicode content
        assert email_data is not None
        assert "测试邮件" in email_data['subject']
        assert "José" in email_data['sender']
        assert "Здравствуй" in email_data['preview']
    
    def test_timezone_handling(self):
        """Test handling of different timezone scenarios"""
        handler = OutlookEmailHandler()
        
        # Test with different date formats
        test_dates = [
            datetime.now(),  # Current time
            datetime(2024, 1, 1, 12, 0, 0),  # Specific date
            None,  # Missing date
        ]
        
        for test_date in test_dates:
            formatted_date = handler._format_date(test_date)
            
            if test_date:
                assert isinstance(formatted_date, str)
                assert len(formatted_date) > 0
            else:
                assert formatted_date == "Unknown Date"
    
    def test_attachment_edge_cases(self):
        """Test edge cases in attachment handling"""
        handler = OutlookEmailHandler()
        
        # Test with various attachment scenarios
        test_cases = [
            # No attachments
            [],
            # Single safe attachment
            [Mock(FileName="document.pdf", Size=1024)],
            # Multiple attachments
            [Mock(FileName="doc1.pdf", Size=1024), Mock(FileName="doc2.txt", Size=512)],
            # Attachment without filename
            [Mock(Size=1024)],
            # Attachment with dangerous name
            [Mock(FileName="virus.exe", Size=1024)],
        ]
        
        for attachments in test_cases:
            mock_message = Mock()
            mock_message.Attachments = attachments
            
            # Set FileName attribute for attachments that don't have it
            for i, attachment in enumerate(attachments):
                if not hasattr(attachment, 'FileName'):
                    attachment.FileName = f"attachment_{i}"
            
            result = handler._check_safe_attachments(mock_message)
            
            # Should handle all cases without crashing
            assert isinstance(result, bool)

if __name__ == "__main__":
    pytest.main([__file__])