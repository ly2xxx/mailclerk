# MailClerk - Email Automation & Sorting

MailClerk is a Streamlit-based email automation application that helps you sort, filter, and download emails from Microsoft Outlook with intelligent rules and user-friendly interface.

## Features

- 🌐 **Dual Email Access**: Choose between web-based (Outlook.com) or desktop Outlook integration
- 📧 **Browser Handler**: Access Outlook.com emails via Microsoft Graph API (works on any platform)
- 💻 **Desktop Handler**: Connect to your local Outlook using win32com.client (Windows only)
- 📋 **Smart Filtering**: Define custom sorting rules based on sender, subject, and keywords
- 📅 **Date Range Filtering**: Filter emails by today, yesterday, last week, or custom date ranges
- 📎 **Attachment Handling**: Download emails with all attachments preserved
- 🔍 **Email Chain Analysis**: Track latest timestamps in email conversations
- 📊 **Statistics Dashboard**: View email distribution and metrics
- 💾 **Local Storage**: Save emails to your local drive with organized folder structure
- ⚙️ **Configuration Management**: Save and load your sorting preferences

## Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/ly2xxx/mailclerk.git
   cd mailclerk
   ```

2. **Create and activate a virtual environment** (recommended):
   ```bash
   # Create virtual environment
   python -m venv venv
   
   # Activate virtual environment
   # On Windows:
   venv\Scripts\activate
   
   # On macOS/Linux (if running via WSL):
   source venv/bin/activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Choose your email access method**:
   - **Browser/Web (Recommended)**: Works with Outlook.com accounts, any platform
   - **Desktop**: Requires Microsoft Outlook installed on Windows

## Usage

1. **Start the application** (ensure virtual environment is activated):
   ```bash
   streamlit run app.py
   ```

2. **Choose Email Handler**:
   - **Browser/Web**: Select this for Outlook.com, Hotmail, or Office 365 accounts
   - **Desktop**: Select this if you have Outlook installed locally on Windows

3. **Connect to Outlook**:
   - **For Browser**: Click "Connect to Browser/Web" and complete OAuth authentication in your browser
   - **For Desktop**: Click "Connect to Desktop Outlook" to connect to your local installation

4. **Configure Sorting Rules**:
   - Go to the "Sorting Rules" tab
   - Define keywords for sender filtering, subject filtering, high priority senders, and exclusions
   - Save your rules for future use

5. **Filter and View Emails**:
   - Select date range in the sidebar
   - Click "Fetch Emails" to retrieve emails based on your filters
   - Review the email list with sender, subject, date, and preview information

6. **Download Emails**:
   - Select emails using checkboxes
   - Click "Download Selected" to save emails to your local drive
   - Each email is saved with metadata, body content, and attachments

## Email Handler Comparison

### 🌐 Browser/Web Handler (Recommended)
- **Pros**: Works on any platform, no software installation needed, reliable authentication
- **Cons**: Requires internet connection, OAuth setup
- **Best for**: Outlook.com, Hotmail, Office 365 personal accounts
- **Authentication**: OAuth 2.0 via web browser

### 💻 Desktop Handler  
- **Pros**: Direct access to local Outlook, works offline
- **Cons**: Windows only, COM registration issues, requires Outlook installation
- **Best for**: Local Outlook installations, corporate environments
- **Authentication**: Direct COM interface

## Configuration

### Sorting Rules

The application supports several types of filtering rules:

- **Sender Keywords**: Filter emails from specific senders
- **Subject Keywords**: Filter emails with specific subjects
- **High Priority Senders**: Mark emails from important senders as high priority
- **Exclude Keywords**: Exclude emails containing specific terms

### Download Settings

- **Download Path**: Customize where emails are saved (default: ~/Downloads/MailClerk)
- **File Organization**: Emails are organized in timestamped folders with metadata
- **Attachment Handling**: All attachments are preserved and saved separately

## File Structure

```
mailclerk/
├── app.py                 # Main Streamlit application
├── email_handler.py       # Outlook email operations
├── config.py             # Configuration management
├── requirements.txt      # Python dependencies
├── README.md            # This file
└── mailclerk_config.json # User configuration (created automatically)
```

## Downloaded Email Structure

When emails are downloaded, they are organized as follows:

```
Downloads/MailClerk/
└── 20240707_143022_Important_Meeting_Notes/
    ├── metadata.json         # Email metadata
    ├── email_body.txt       # Plain text body
    ├── email_body.html      # HTML body (if available)
    └── attachments/         # Email attachments
        ├── document.pdf
        └── image.jpg
```

## System Requirements

### For Browser/Web Handler (Recommended)
- **Operating System**: Any (Windows, macOS, Linux)
- **Python**: 3.8 or higher
- **Internet Connection**: Required for API access
- **Email Account**: Outlook.com, Hotmail, or Office 365 personal account
- **Dependencies**: See requirements.txt

### For Desktop Handler
- **Operating System**: Windows (required for win32com.client)
- **Python**: 3.8 or higher  
- **Microsoft Outlook**: Installed and configured locally
- **Dependencies**: See requirements.txt

### Both Handlers
- **Virtual Environment**: Recommended for isolation

## Troubleshooting

### Common Issues

1. **"Failed to connect to Outlook"** / **COM Error (-2147221005)**:
   - **QUICK FIX**: Run `quick_fix.bat` as Administrator (double-click and select "Run as Administrator")
   - **ADVANCED FIX**: Run `python fix_outlook_com.py` as Administrator
   - **DIAGNOSTIC**: Run `python outlook_diagnostic.py` to identify the issue
   - Ensure Microsoft Outlook is installed and running
   - Check that your Outlook profile is configured
   - Try restarting Outlook and the application
   - Verify that your virtual environment is activated
   - **Run as Administrator** (most common solution)
   - Ensure Python and Office have matching architecture (both 32-bit or both 64-bit)

2. **"No emails found"**:
   - Check your date range filters
   - Verify your sorting rules aren't too restrictive
   - Ensure you have emails in your inbox for the selected period

3. **"Download failed"**:
   - Check that the download path exists and is writable
   - Ensure sufficient disk space
   - Some emails with large attachments may take longer to download

### Logging

The application creates a `mailclerk.log` file that contains detailed information about operations and any errors encountered.

## Security Considerations

- The application uses your local Outlook installation and doesn't store email credentials
- Downloaded emails are saved locally on your machine
- Configuration files are stored locally and contain only filtering rules
- No email data is transmitted over the internet by this application

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is open source and available under the MIT License.

## Support

If you encounter any issues or have questions, please create an issue in the GitHub repository.