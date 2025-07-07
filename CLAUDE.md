# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

### Running the application
- `streamlit run app.py` - Start the Streamlit web application
- `pip install -r requirements.txt` - Install dependencies

### Testing
- `pytest` - Run all tests
- `pytest test_integration.py` - Run integration tests specifically
- `pytest -m "not slow"` - Run tests excluding slow ones
- `pytest -m security` - Run security-focused tests only
- `pytest -v` - Run tests with verbose output

### Dependencies
Main dependencies include:
- `streamlit` - Web application framework
- `pandas` - Data manipulation
- `pywin32` - Windows COM interface for Outlook
- `pytest` - Testing framework

## Architecture Overview

This is a **Streamlit-based email automation application** that connects to Microsoft Outlook to fetch, filter, sort, and download emails with intelligent rules and security features.

### Core Components

**Main Application (`app.py`)**:
- Streamlit web interface with tabs for Email Management, Sorting Rules, and Statistics
- Dynamic handler selection based on available email handler modules
- Session state management for connection status, email data, and user selections
- Security-focused input validation and sanitization throughout
- Adaptive UI that changes based on available handlers

**Email Handler Factory (`email_handler_factory.py`)**:
- `EmailHandlerFactory` class for dynamic discovery and creation of email handlers
- Automatic platform compatibility checking (Windows-only handlers, cross-platform handlers)
- Dependency validation to ensure required modules are available
- Runtime handler refresh capability for development and modularity
- Supports plugin-like architecture for adding new email handlers

**Email Handler Base (`email_handler_base.py`)**:
- Abstract base class defining the common interface for all email handlers
- Shared validation and filtering methods used by all implementations
- Ensures consistent API across different handler types

**Browser Email Handler (`email_browser_handler.py`)**:
- `BrowserEmailHandler` class using Microsoft Graph API for web-based email access
- OAuth 2.0 authentication flow with token management
- Cross-platform compatibility (Windows, macOS, Linux)
- Secure API requests with proper error handling and rate limiting

**Desktop Email Handler (`email_handler.py`)**:
- `OutlookEmailHandler` class for Windows COM-based Outlook integration
- Secure email fetching with date range validation and content size limits
- Attachment handling with file extension validation and size restrictions
- Email downloading with path traversal protection and filename sanitization
- Built-in diagnostics for COM registration and Outlook installation issues

**Configuration Management (`config.py`)**:
- `EmailConfig` class handles loading, saving, and validation of user preferences
- Secure file operations with path validation and size limits
- Rule validation with keyword sanitization and length restrictions
- Support for sorting rules, download settings, and connection parameters

**Security Architecture**:
- Comprehensive input sanitization using regex patterns to remove control characters
- Path traversal protection for downloads and file operations
- File extension blocking for dangerous attachment types (.exe, .bat, .cmd, etc.)
- Size limits on emails (100MB), attachments (50MB), and configuration files (1MB)
- Maximum limits on keywords per category (50) and keyword length (100 chars)

### Key Technical Details

**Email Processing Flow**:
1. Connect to Outlook via win32com.client (Windows only)
2. Fetch emails using date range filters with SQL-like syntax
3. Apply user-defined sorting rules (sender, subject, priority, exclusion keywords)
4. Extract and sanitize email metadata (subject, sender, size, attachments)
5. Generate secure identifiers using SHA256 hashing
6. Download selected emails with attachments to organized folder structure

**Data Validation**:
- All user inputs are sanitized using `sanitize_user_input()` function
- Date ranges are validated to prevent fetching emails older than 365 days
- Email sizes are checked against security limits before processing
- File paths are validated to prevent access to system directories

**Configuration Structure**:
- Rules stored in `mailclerk_config.json` with atomic read/write operations
- Default rules provided for sender keywords, subject keywords, high priority senders, and exclusion keywords
- Download settings include path validation, subfolder creation, and attachment type filtering

**Testing Strategy**:
- Comprehensive integration tests in `test_integration.py`
- Security-focused tests for injection attacks and path traversal
- Performance tests for large email volumes (up to 1000 emails)
- Memory management and resource cleanup validation
- Unicode and timezone handling tests

### Security Considerations

The application implements multiple layers of security:
- **Input Validation**: All user inputs sanitized and length-limited
- **Path Security**: Download paths validated to prevent traversal attacks
- **File Security**: Dangerous file extensions blocked for attachments
- **Size Limits**: Email and attachment sizes capped to prevent DoS
- **COM Security**: Windows-only operation with proper error handling
- **Configuration Security**: JSON config files with integrity validation

### Error Handling

Robust error handling throughout:
- Connection failures gracefully handled with user feedback
- Invalid email data processed without crashing
- Security errors logged without exposing sensitive information
- Fallback to default configurations on validation failures
- Progress indicators for long-running operations like email downloads