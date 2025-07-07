import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import logging
from email_handler import OutlookEmailHandler
from config import EmailConfig

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('mailclerk.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def init_session_state():
    """Initialize Streamlit session state variables"""
    if 'email_handler' not in st.session_state:
        st.session_state.email_handler = None
    if 'emails_df' not in st.session_state:
        st.session_state.emails_df = pd.DataFrame()
    if 'selected_emails' not in st.session_state:
        st.session_state.selected_emails = []
    if 'sorting_rules' not in st.session_state:
        st.session_state.sorting_rules = EmailConfig.get_default_rules()
    if 'download_path' not in st.session_state:
        st.session_state.download_path = os.path.expanduser("~/Downloads/MailClerk")

def setup_sidebar():
    """Setup sidebar with configuration options"""
    st.sidebar.title("📧 MailClerk Configuration")
    
    # Outlook Connection
    st.sidebar.header("Outlook Connection")
    if st.sidebar.button("Connect to Outlook"):
        try:
            with st.spinner("Connecting to Outlook..."):
                st.session_state.email_handler = OutlookEmailHandler()
                if st.session_state.email_handler.connect():
                    st.sidebar.success("✅ Connected to Outlook")
                    logger.info("Successfully connected to Outlook")
                else:
                    st.sidebar.error("❌ Failed to connect to Outlook")
                    logger.error("Failed to connect to Outlook")
        except Exception as e:
            st.sidebar.error(f"Connection error: {str(e)}")
            logger.error(f"Outlook connection error: {str(e)}")
    
    # Download Settings
    st.sidebar.header("Download Settings")
    st.session_state.download_path = st.sidebar.text_input(
        "Download Path", 
        value=st.session_state.download_path
    )
    
    # Create download directory if it doesn't exist
    if not os.path.exists(st.session_state.download_path):
        os.makedirs(st.session_state.download_path, exist_ok=True)
    
    # Date Range Filter
    st.sidebar.header("Date Filters")
    date_option = st.sidebar.selectbox(
        "Select Date Range",
        ["Today", "Yesterday", "Last 3 days", "Last Week", "Custom Range"]
    )
    
    if date_option == "Custom Range":
        start_date = st.sidebar.date_input("Start Date", datetime.now().date())
        end_date = st.sidebar.date_input("End Date", datetime.now().date())
    else:
        start_date, end_date = get_date_range(date_option)
    
    return date_option, start_date, end_date

def get_date_range(option):
    """Get date range based on selected option"""
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
    """Setup sorting rules configuration"""
    st.header("📋 Email Sorting Rules")
    
    with st.expander("Configure Sorting Rules", expanded=False):
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Sender Filters")
            sender_keywords = st.text_area(
                "Sender Keywords (one per line)",
                value="\n".join(st.session_state.sorting_rules.get('sender_keywords', [])),
                height=100
            )
            
            st.subheader("Subject Filters")
            subject_keywords = st.text_area(
                "Subject Keywords (one per line)",
                value="\n".join(st.session_state.sorting_rules.get('subject_keywords', [])),
                height=100
            )
        
        with col2:
            st.subheader("Priority Settings")
            high_priority_senders = st.text_area(
                "High Priority Senders (one per line)",
                value="\n".join(st.session_state.sorting_rules.get('high_priority_senders', [])),
                height=100
            )
            
            st.subheader("Exclude Filters")
            exclude_keywords = st.text_area(
                "Exclude Keywords (one per line)",
                value="\n".join(st.session_state.sorting_rules.get('exclude_keywords', [])),
                height=100
            )
        
        if st.button("Save Sorting Rules"):
            st.session_state.sorting_rules = {
                'sender_keywords': [s.strip() for s in sender_keywords.split('\n') if s.strip()],
                'subject_keywords': [s.strip() for s in subject_keywords.split('\n') if s.strip()],
                'high_priority_senders': [s.strip() for s in high_priority_senders.split('\n') if s.strip()],
                'exclude_keywords': [s.strip() for s in exclude_keywords.split('\n') if s.strip()]
            }
            EmailConfig.save_rules(st.session_state.sorting_rules)
            st.success("✅ Sorting rules saved!")

def fetch_and_display_emails(start_date, end_date):
    """Fetch and display emails based on filters"""
    if st.session_state.email_handler is None:
        st.warning("⚠️ Please connect to Outlook first")
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
                    st.session_state.emails_df = pd.DataFrame(emails)
                    st.success(f"✅ Fetched {len(emails)} emails")
                    logger.info(f"Fetched {len(emails)} emails")
                else:
                    st.info("ℹ️ No emails found for the selected criteria")
                    st.session_state.emails_df = pd.DataFrame()
        except Exception as e:
            st.error(f"Error fetching emails: {str(e)}")
            logger.error(f"Error fetching emails: {str(e)}")
    
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
            if st.button("📥 Download Selected"):
                download_selected_emails()
        
        # Display emails with checkboxes
        for idx, email in st.session_state.emails_df.iterrows():
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
                    
                    st.markdown(f"""
                    **{priority_icon} From:** {email['sender']} {attachment_icon}  
                    **Subject:** {email['subject']}  
                    **Date:** {email['received_date']}  
                    **Chain Timestamp:** {email['chain_timestamp']}  
                    **Size:** {email.get('size', 'N/A')}  
                    **Preview:** {email.get('preview', 'No preview available')[:100]}...
                    """)
                    
                st.divider()

def download_selected_emails():
    """Download selected emails to local storage"""
    if not st.session_state.selected_emails:
        st.warning("⚠️ No emails selected for download")
        return
    
    try:
        with st.spinner(f"Downloading {len(st.session_state.selected_emails)} emails..."):
            progress_bar = st.progress(0)
            
            for i, email_idx in enumerate(st.session_state.selected_emails):
                email = st.session_state.emails_df.iloc[email_idx]
                
                # Download email
                success = st.session_state.email_handler.download_email(
                    email,
                    st.session_state.download_path
                )
                
                if success:
                    logger.info(f"Downloaded email: {email['subject']}")
                else:
                    logger.error(f"Failed to download email: {email['subject']}")
                
                # Update progress
                progress_bar.progress((i + 1) / len(st.session_state.selected_emails))
            
            st.success(f"✅ Downloaded {len(st.session_state.selected_emails)} emails to {st.session_state.download_path}")
            
            # Clear selection after download
            st.session_state.selected_emails = []
            
    except Exception as e:
        st.error(f"Error downloading emails: {str(e)}")
        logger.error(f"Error downloading emails: {str(e)}")

def main():
    """Main application"""
    st.set_page_config(
        page_title="MailClerk - Email Automation",
        page_icon="📧",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    st.title("📧 MailClerk - Email Automation & Sorting")
    st.markdown("Automate your email management with intelligent sorting and downloading capabilities.")
    
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
        if not st.session_state.emails_df.empty:
            st.header("📊 Email Statistics")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Emails", len(st.session_state.emails_df))
            with col2:
                high_priority = st.session_state.emails_df.get('is_high_priority', pd.Series()).sum()
                st.metric("High Priority", high_priority)
            with col3:
                with_attachments = st.session_state.emails_df.get('has_attachments', pd.Series()).sum()
                st.metric("With Attachments", with_attachments)
            with col4:
                selected_count = len(st.session_state.selected_emails)
                st.metric("Selected", selected_count)
            
            # Sender distribution
            if 'sender' in st.session_state.emails_df.columns:
                st.subheader("📈 Sender Distribution")
                sender_counts = st.session_state.emails_df['sender'].value_counts().head(10)
                st.bar_chart(sender_counts)
        else:
            st.info("ℹ️ Fetch emails to view statistics")

if __name__ == "__main__":
    main()