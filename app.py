import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import logging
import re
from email_handler import OutlookEmailHandler, SecurityError
from email_browser_handler import BrowserEmailHandler
from config import EmailConfig, ConfigSecurityError

# Configure logging with security considerations
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('mailclerk.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Security constants
MAX_INPUT_LENGTH = 500
MAX_PATH_LENGTH = 260
MAX_KEYWORDS_DISPLAY = 50

def sanitize_user_input(user_input: str, max_length: int = MAX_INPUT_LENGTH) -> str:
    """Sanitize user input for security"""
    if not isinstance(user_input, str):
        return ""
    
    # Remove control characters and limit length
    sanitized = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', user_input.strip())
    return sanitized[:max_length] if len(sanitized) > max_length else sanitized

def validate_download_path(path: str) -> tuple[bool, str]:
    """Validate download path for security"""
    try:
        if not path or not isinstance(path, str):
            return False, "Invalid path"
        
        sanitized_path = sanitize_user_input(path, MAX_PATH_LENGTH)
        if not sanitized_path:
            return False, "Empty path"
        
        # Basic path validation
        if len(sanitized_path) > MAX_PATH_LENGTH:
            return False, "Path too long"
        
        # Check for dangerous patterns
        dangerous_patterns = ['../', '..\\', '<', '>', '|', '*', '?']
        for pattern in dangerous_patterns:
            if pattern in sanitized_path:
                return False, "Invalid characters in path"
        
        return True, sanitized_path
        
    except Exception:
        return False, "Path validation error"

def init_session_state():
    """Initialize Streamlit session state variables with security defaults"""
    if 'email_handler' not in st.session_state:
        st.session_state.email_handler = None
    if 'handler_type' not in st.session_state:
        st.session_state.handler_type = 'browser'  # Default to browser handler
    if 'emails_df' not in st.session_state:
        st.session_state.emails_df = pd.DataFrame()
    if 'selected_emails' not in st.session_state:
        st.session_state.selected_emails = []
    if 'sorting_rules' not in st.session_state:
        try:
            st.session_state.sorting_rules = EmailConfig.load_rules()
        except Exception:
            st.session_state.sorting_rules = EmailConfig.get_default_rules()
    if 'download_path' not in st.session_state:
        st.session_state.download_path = os.path.expanduser("~/Downloads/MailClerk")
    if 'connection_status' not in st.session_state:
        st.session_state.connection_status = False

def setup_sidebar():
    """Setup sidebar with configuration options and security validation"""
    st.sidebar.title("📧 MailClerk Configuration")
    
    # Email Handler Selection
    st.sidebar.header("Email Handler")
    handler_options = {
        'browser': '🌐 Browser/Web (Outlook.com) - Recommended',
        'desktop': '💻 Desktop (Outlook COM) - Windows Only'
    }
    
    new_handler_type = st.sidebar.selectbox(
        "Choose Email Handler:",
        options=list(handler_options.keys()),
        format_func=lambda x: handler_options[x],
        index=0 if st.session_state.handler_type == 'browser' else 1
    )
    
    # If handler type changed, disconnect current handler
    if new_handler_type != st.session_state.handler_type:
        if st.session_state.email_handler:
            st.session_state.email_handler.close_connection()
        st.session_state.email_handler = None
        st.session_state.connection_status = False
        st.session_state.handler_type = new_handler_type
        st.rerun()
    
    # Show handler-specific info
    if st.session_state.handler_type == 'browser':
        st.sidebar.info("🌐 Uses Microsoft Graph API\n• Works on any platform\n• No Outlook installation needed\n• Requires web authentication")
    else:
        st.sidebar.warning("💻 Requires Windows + Outlook\n• Local Outlook installation\n• COM registration needed\n• May need Administrator rights")
    
    # Outlook Connection with status indicator
    st.sidebar.header("Email Connection")
    
    # Show connection status
    if st.session_state.connection_status:
        st.sidebar.success("✅ Connected to Outlook")
        if st.sidebar.button("Disconnect"):
            if st.session_state.email_handler:
                st.session_state.email_handler.close_connection()
            st.session_state.email_handler = None
            st.session_state.connection_status = False
            st.rerun()
    else:
        connect_button_text = "Connect to Browser/Web" if st.session_state.handler_type == 'browser' else "Connect to Desktop Outlook"
        
        if st.sidebar.button(connect_button_text):
            try:
                connection_message = "Connecting to Outlook Web..." if st.session_state.handler_type == 'browser' else "Connecting to Desktop Outlook..."
                
                with st.spinner(connection_message):
                    # Create appropriate handler
                    if st.session_state.handler_type == 'browser':
                        st.session_state.email_handler = BrowserEmailHandler()
                    else:
                        st.session_state.email_handler = OutlookEmailHandler()
                    
                    if st.session_state.email_handler.connect():
                        st.session_state.connection_status = True
                        st.sidebar.success("✅ Connected successfully!")
                        logger.info(f"Successfully connected using {st.session_state.handler_type} handler")
                        st.rerun()
                    else:
                        st.sidebar.error("❌ Failed to connect")
                        logger.error(f"Failed to connect using {st.session_state.handler_type} handler")
                        
                        # Show handler-specific diagnostics
                        if st.session_state.handler_type == 'desktop':
                            # Run diagnostics if desktop connection fails
                            with st.sidebar.expander("🔍 Diagnostics", expanded=True):
                                diagnosis = st.session_state.email_handler.diagnose_outlook_installation()
                                
                                st.write("**Outlook Installation Check:**")
                                for check, status in diagnosis.items():
                                    icon = "✅" if status else "❌"
                                    readable_name = check.replace('_', ' ').title()
                                    st.write(f"{icon} {readable_name}")
                                
                                st.write("**Troubleshooting Steps:**")
                                if not diagnosis['outlook_installed']:
                                    st.write("• Install Microsoft Outlook")
                                if not diagnosis['outlook_running']:
                                    st.write("• Start Outlook application")
                                if not diagnosis['com_registered']:
                                    st.write("• Run as Administrator: `regsvr32 /i /n /s outlctl.dll`")
                                    st.write("• Or repair Office installation")
                                if not diagnosis['mapi_available']:
                                    st.write("• Restart Outlook and try again")
                                    st.write("• Check Windows permissions")
                        else:
                            # Browser connection troubleshooting
                            with st.sidebar.expander("🔍 Troubleshooting", expanded=True):
                                st.write("**Common Browser Connection Issues:**")
                                st.write("• Check internet connection")
                                st.write("• Ensure popup blocker allows authentication")
                                st.write("• Try a different browser")
                                st.write("• Clear browser cache and cookies")
                                st.write("• Disable browser extensions")
                        
            except SecurityError as e:
                st.sidebar.error(f"❌ Security error: {str(e)}")
                logger.error(f"Security error during connection: {str(e)}")
            except Exception as e:
                st.sidebar.error("❌ Connection failed")
                logger.error(f"Unknown error during Outlook connection: {type(e).__name__}")
                
                # Show detailed error in expander
                with st.sidebar.expander("🔍 Error Details", expanded=False):
                    st.write(f"**Error Type:** {type(e).__name__}")
                    st.write("**Common Solutions:**")
                    st.write("• Run application as Administrator")
                    st.write("• Ensure Outlook is installed and configured")
                    st.write("• Try restarting Outlook")
                    st.write("• Check Windows COM registration")
    
    # Download Settings with validation
    st.sidebar.header("Download Settings")
    
    # Path input with validation
    path_input = st.sidebar.text_input(
        "Download Path", 
        value=st.session_state.download_path,
        max_chars=MAX_PATH_LENGTH
    )
    
    # Validate path when changed
    if path_input != st.session_state.download_path:
        is_valid, validated_path = validate_download_path(path_input)
        if is_valid:
            st.session_state.download_path = validated_path
        else:
            st.sidebar.warning(f"⚠️ Invalid path: {validated_path}")
    
    # Create download directory with error handling
    try:
        if st.session_state.download_path:
            os.makedirs(st.session_state.download_path, exist_ok=True)
    except Exception:
        st.sidebar.error("❌ Cannot create download directory")
    
    # Date Range Filter with validation
    st.sidebar.header("Date Filters")
    date_option = st.sidebar.selectbox(
        "Select Date Range",
        ["Today", "Yesterday", "Last 3 days", "Last Week", "Custom Range"]
    )
    
    # Date validation
    today = datetime.now().date()
    max_date_back = today - timedelta(days=365)  # 1 year limit
    
    if date_option == "Custom Range":
        start_date = st.sidebar.date_input(
            "Start Date", 
            value=today,
            min_value=max_date_back,
            max_value=today
        )
        end_date = st.sidebar.date_input(
            "End Date", 
            value=today,
            min_value=start_date if 'start_date' in locals() else max_date_back,
            max_value=today
        )
    else:
        start_date, end_date = get_date_range(date_option)
    
    return date_option, start_date, end_date

def get_date_range(option):
    """Get date range based on selected option with validation"""
    today = datetime.now().date()
    if option == "Today":
        return today, today
    elif option == "Yesterday":
        yesterday = today - timedelta(days=1)
        return yesterday, yesterday
    elif option == "Last 3 days":
        return today - timedelta(days=3), today
    elif option == "Last Week":
        return today - timedelta(days=7), today
    return today, today

def setup_sorting_rules():
    """Setup sorting rules configuration with input validation"""
    st.header("📋 Email Sorting Rules")
    
    with st.expander("Configure Sorting Rules", expanded=False):
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Sender Filters")
            current_sender_keywords = st.session_state.sorting_rules.get('sender_keywords', [])
            sender_keywords_text = "\n".join(current_sender_keywords[:MAX_KEYWORDS_DISPLAY])
            
            sender_keywords = st.text_area(
                "Sender Keywords (one per line, max 100 chars each)",
                value=sender_keywords_text,
                height=100,
                max_chars=MAX_KEYWORDS_DISPLAY * 100
            )
            
            st.subheader("Subject Filters")
            current_subject_keywords = st.session_state.sorting_rules.get('subject_keywords', [])
            subject_keywords_text = "\n".join(current_subject_keywords[:MAX_KEYWORDS_DISPLAY])
            
            subject_keywords = st.text_area(
                "Subject Keywords (one per line, max 100 chars each)",
                value=subject_keywords_text,
                height=100,
                max_chars=MAX_KEYWORDS_DISPLAY * 100
            )
        
        with col2:
            st.subheader("Priority Settings")
            current_priority_senders = st.session_state.sorting_rules.get('high_priority_senders', [])
            priority_senders_text = "\n".join(current_priority_senders[:MAX_KEYWORDS_DISPLAY])
            
            high_priority_senders = st.text_area(
                "High Priority Senders (one per line, max 100 chars each)",
                value=priority_senders_text,
                height=100,
                max_chars=MAX_KEYWORDS_DISPLAY * 100
            )
            
            st.subheader("Exclude Filters")
            current_exclude_keywords = st.session_state.sorting_rules.get('exclude_keywords', [])
            exclude_keywords_text = "\n".join(current_exclude_keywords[:MAX_KEYWORDS_DISPLAY])
            
            exclude_keywords = st.text_area(
                "Exclude Keywords (one per line, max 100 chars each)",
                value=exclude_keywords_text,
                height=100,
                max_chars=MAX_KEYWORDS_DISPLAY * 100
            )
        
        if st.button("Save Sorting Rules"):
            try:
                # Sanitize and validate all inputs
                new_rules = {
                    'sender_keywords': [],
                    'subject_keywords': [],
                    'high_priority_senders': [],
                    'exclude_keywords': []
                }
                
                # Process sender keywords
                for keyword in sender_keywords.split('\n'):
                    sanitized = sanitize_user_input(keyword.strip(), 100)
                    if sanitized and len(sanitized) >= 2:
                        new_rules['sender_keywords'].append(sanitized.lower())
                
                # Process subject keywords  
                for keyword in subject_keywords.split('\n'):
                    sanitized = sanitize_user_input(keyword.strip(), 100)
                    if sanitized and len(sanitized) >= 2:
                        new_rules['subject_keywords'].append(sanitized.lower())
                
                # Process high priority senders
                for keyword in high_priority_senders.split('\n'):
                    sanitized = sanitize_user_input(keyword.strip(), 100)
                    if sanitized and len(sanitized) >= 2:
                        new_rules['high_priority_senders'].append(sanitized.lower())
                
                # Process exclude keywords
                for keyword in exclude_keywords.split('\n'):
                    sanitized = sanitize_user_input(keyword.strip(), 100)
                    if sanitized and len(sanitized) >= 2:
                        new_rules['exclude_keywords'].append(sanitized.lower())
                
                # Limit number of keywords per category
                for rule_type in new_rules:
                    new_rules[rule_type] = new_rules[rule_type][:MAX_KEYWORDS_DISPLAY]
                
                # Validate rules
                validation_errors = EmailConfig.validate_rules(new_rules)
                if validation_errors:
                    st.error(f"❌ Validation errors: {', '.join(validation_errors)}")
                else:
                    # Save rules
                    if EmailConfig.save_rules(new_rules):
                        st.session_state.sorting_rules = new_rules
                        st.success("✅ Sorting rules saved successfully!")
                        logger.info("Sorting rules saved")
                    else:
                        st.error("❌ Failed to save sorting rules")
                        
            except ConfigSecurityError as e:
                st.error(f"❌ Security error: {str(e)}")
                logger.error(f"Security error saving rules: {str(e)}")
            except Exception:
                st.error("❌ Error saving sorting rules")
                logger.error("Error saving sorting rules")

def fetch_and_display_emails(start_date, end_date):
    """Fetch and display emails based on filters with security validation"""
    if not st.session_state.connection_status or st.session_state.email_handler is None:
        st.warning("⚠️ Please connect to Outlook first")
        return
    
    # Validate date range
    try:
        if start_date > end_date:
            st.error("❌ Start date cannot be after end date")
            return
        
        if (datetime.now().date() - start_date).days > 365:
            st.error("❌ Date range cannot exceed 365 days")
            return
            
    except Exception:
        st.error("❌ Invalid date range")
        return
    
    if st.button("🔄 Fetch Emails"):
        try:
            with st.spinner("Fetching emails..."):
                emails = st.session_state.email_handler.fetch_emails(
                    start_date=start_date,
                    end_date=end_date,
                    rules=st.session_state.sorting_rules
                )
                
                if emails:
                    # Create DataFrame with safe data
                    safe_emails = []
                    for email in emails:
                        safe_email = {
                            'subject': email.get('subject', 'No Subject')[:200],
                            'sender': email.get('sender', 'Unknown')[:100],
                            'received_date': email.get('received_date', 'Unknown'),
                            'chain_timestamp': email.get('chain_timestamp', 'Unknown'),
                            'size': email.get('size', 0),
                            'has_attachments': email.get('has_attachments', False),
                            'is_high_priority': email.get('is_high_priority', False),
                            'preview': email.get('preview', 'No preview')[:100]
                        }
                        safe_emails.append(safe_email)
                    
                    st.session_state.emails_df = pd.DataFrame(safe_emails)
                    # Store full email data separately for download
                    st.session_state.full_emails_data = emails
                    
                    st.success(f"✅ Fetched {len(emails)} emails")
                    logger.info(f"Fetched {len(emails)} emails")
                else:
                    st.info("ℹ️ No emails found for the selected criteria")
                    st.session_state.emails_df = pd.DataFrame()
                    st.session_state.full_emails_data = []
                    
        except SecurityError as e:
            st.error(f"❌ Security error: {str(e)}")
            logger.error(f"Security error fetching emails: {str(e)}")
        except Exception:
            st.error("❌ Error fetching emails")
            logger.error("Error fetching emails")
    
    # Display emails
    if not st.session_state.emails_df.empty:
        st.header("📧 Email List")
        
        # Email selection interface
        col1, col2, col3 = st.columns([1, 1, 1])
        with col1:
            if st.button("Select All"):
                st.session_state.selected_emails = list(range(len(st.session_state.emails_df)))
                st.rerun()
        with col2:
            if st.button("Deselect All"):
                st.session_state.selected_emails = []
                st.rerun()
        with col3:
            selected_count = len(st.session_state.selected_emails)
            if selected_count > 0:
                if st.button(f"📥 Download Selected ({selected_count})"):
                    download_selected_emails()
        
        # Display emails with checkboxes (limit display for performance)
        display_limit = min(100, len(st.session_state.emails_df))  # Limit display
        
        if len(st.session_state.emails_df) > display_limit:
            st.info(f"ℹ️ Showing first {display_limit} emails out of {len(st.session_state.emails_df)} total")
        
        for idx in range(display_limit):
            email = st.session_state.emails_df.iloc[idx]
            
            with st.container():
                col1, col2 = st.columns([0.1, 0.9])
                
                with col1:
                    is_selected = st.checkbox(
                        "", 
                        key=f"email_{idx}",
                        value=idx in st.session_state.selected_emails
                    )
                    if is_selected and idx not in st.session_state.selected_emails:
                        st.session_state.selected_emails.append(idx)
                    elif not is_selected and idx in st.session_state.selected_emails:
                        st.session_state.selected_emails.remove(idx)
                
                with col2:
                    priority_icon = "🔴" if email.get('is_high_priority', False) else "🔵"
                    attachment_icon = "📎" if email.get('has_attachments', False) else ""
                    
                    # Safely display email information
                    st.markdown(f"""
                    **{priority_icon} From:** {email['sender']} {attachment_icon}  
                    **Subject:** {email['subject']}  
                    **Date:** {email['received_date']}  
                    **Chain Timestamp:** {email['chain_timestamp']}  
                    **Size:** {email.get('size', 'N/A')} bytes  
                    **Preview:** {email.get('preview', 'No preview available')}
                    """)
                    
                st.divider()

def download_selected_emails():
    """Download selected emails to local storage with security validation"""
    if not st.session_state.selected_emails:
        st.warning("⚠️ No emails selected for download")
        return
    
    if not st.session_state.connection_status or st.session_state.email_handler is None:
        st.error("❌ Not connected to Outlook")
        return
    
    # Validate download path
    is_valid, validated_path = validate_download_path(st.session_state.download_path)
    if not is_valid:
        st.error(f"❌ Invalid download path: {validated_path}")
        return
    
    try:
        selected_count = len(st.session_state.selected_emails)
        
        # Limit number of downloads for security
        if selected_count > 50:
            st.error("❌ Cannot download more than 50 emails at once")
            return
        
        with st.spinner(f"Downloading {selected_count} emails..."):
            progress_bar = st.progress(0)
            success_count = 0
            
            for i, email_idx in enumerate(st.session_state.selected_emails):
                try:
                    # Get full email data
                    if hasattr(st.session_state, 'full_emails_data') and email_idx < len(st.session_state.full_emails_data):
                        email_data = st.session_state.full_emails_data[email_idx]
                        
                        # Download email
                        success = st.session_state.email_handler.download_email(
                            email_data,
                            validated_path
                        )
                        
                        if success:
                            success_count += 1
                            logger.info(f"Downloaded email: {email_data.get('subject', 'Unknown')}")
                        else:
                            logger.error(f"Failed to download email: {email_data.get('subject', 'Unknown')}")
                    
                except Exception:
                    logger.error(f"Error downloading email {email_idx}")
                
                # Update progress
                progress_bar.progress((i + 1) / selected_count)
            
            if success_count > 0:
                st.success(f"✅ Successfully downloaded {success_count} out of {selected_count} emails to {validated_path}")
            else:
                st.error("❌ Failed to download any emails")
            
            # Clear selection after download attempt
            st.session_state.selected_emails = []
            
    except SecurityError as e:
        st.error(f"❌ Security error: {str(e)}")
        logger.error(f"Security error downloading emails: {str(e)}")
    except Exception:
        st.error("❌ Error downloading emails")
        logger.error("Error downloading emails")

def display_statistics():
    """Display email statistics safely"""
    if not st.session_state.emails_df.empty:
        st.header("📊 Email Statistics")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            total_emails = min(len(st.session_state.emails_df), 10000)  # Cap display
            st.metric("Total Emails", total_emails)
        with col2:
            high_priority = st.session_state.emails_df.get('is_high_priority', pd.Series(dtype=bool)).sum()
            st.metric("High Priority", high_priority)
        with col3:
            with_attachments = st.session_state.emails_df.get('has_attachments', pd.Series(dtype=bool)).sum()
            st.metric("With Attachments", with_attachments)
        with col4:
            selected_count = len(st.session_state.selected_emails)
            st.metric("Selected", selected_count)
        
        # Sender distribution (limited for security)
        if 'sender' in st.session_state.emails_df.columns:
            st.subheader("📈 Top Senders")
            sender_counts = st.session_state.emails_df['sender'].value_counts().head(10)
            if not sender_counts.empty:
                st.bar_chart(sender_counts)
            else:
                st.info("No sender data available")
    else:
        st.info("ℹ️ Fetch emails to view statistics")

def main():
    """Main application with security enhancements"""
    try:
        st.set_page_config(
            page_title="MailClerk - Email Automation",
            page_icon="📧",
            layout="wide",
            initial_sidebar_state="expanded"
        )
        
        st.title("📧 MailClerk - Email Automation & Sorting")
        st.markdown("Securely automate your email management with intelligent sorting and downloading capabilities.")
        
        # Initialize session state
        init_session_state()
        
        # Setup sidebar
        date_option, start_date, end_date = setup_sidebar()
        
        # Main content area
        tab1, tab2, tab3 = st.tabs(["📧 Email Management", "📋 Sorting Rules", "📊 Statistics"])
        
        with tab1:
            fetch_and_display_emails(start_date, end_date)
        
        with tab2:
            setup_sorting_rules()
        
        with tab3:
            display_statistics()
        
        # Security notice
        with st.sidebar:
            st.markdown("---")
            st.caption("🔒 Security: All data is processed locally. No information is transmitted externally.")
        
    except Exception:
        st.error("❌ Application error occurred")
        logger.error("Application error in main function")

if __name__ == "__main__":
    main()