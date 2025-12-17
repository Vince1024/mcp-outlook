"""
MCP Server for Microsoft Outlook

This module provides a FastMCP server that interfaces with Microsoft Outlook via COM automation.
It enables AI assistants and automation tools to interact with Outlook data including emails,
calendar events, and contacts.

Features:
    - Email management: Read, send, search, create drafts, and manage attachments
    - Calendar management: Read, create, search events, and respond to meeting invitations
    - Contact management: Read, create, and search contacts
    - Out-of-Office management: Get and set automatic reply settings

Requirements:
    - Microsoft Outlook installed and configured on Windows
    - Python packages: win32com, fastmcp, python-dateutil

Usage:
    Run as a standalone MCP server:
        python outlook_mcp.py
    
    Or import and use functions directly:
        from outlook_mcp import get_inbox_emails
        
Security Notes:
    - This module accesses local Outlook data via COM
    - No credentials are logged or transmitted
    - Email body content is truncated in responses to prevent data leakage
    
Version: 1.2.0
"""

import json
import logging
from datetime import datetime, timedelta
from typing import Optional, Dict, Any

import win32com.client
from dateutil import parser as date_parser
from fastmcp import FastMCP

# ============================================================================
# CONFIGURATION AND CONSTANTS
# ============================================================================

# Configure logging - completely silent
logging.basicConfig(
    level=logging.CRITICAL,  # Only critical errors
    format='%(message)s',
    handlers=[logging.NullHandler()]  # No output
)

# Set our logger to CRITICAL (completely silent)
logger = logging.getLogger(__name__)
logger.setLevel(logging.CRITICAL)

# Silence all MCP/FastMCP related loggers completely
for logger_name in ['mcp', 'FastMCP', 'fastmcp', 'mcp.server', 'fastmcp.server', 
                     'mcp.client', 'fastmcp.client', 'asyncio', '__main__']:
    logging.getLogger(logger_name).setLevel(logging.CRITICAL)
    logging.getLogger(logger_name).addHandler(logging.NullHandler())
    logging.getLogger(logger_name).propagate = False

# Initialize FastMCP server
mcp = FastMCP("outlook")

# Outlook folder constants (from Microsoft Outlook Object Model documentation)
# These constants represent the default folder IDs in the Outlook namespace
OUTLOOK_FOLDER_INBOX = 6      # Inbox folder for incoming emails
OUTLOOK_FOLDER_SENT = 5        # Sent items folder
OUTLOOK_FOLDER_DRAFTS = 16     # Drafts folder for unsent emails
OUTLOOK_FOLDER_DELETED = 3     # Deleted items (trash)
OUTLOOK_FOLDER_OUTBOX = 4      # Outbox for emails pending send
OUTLOOK_FOLDER_JUNK = 23       # Junk/spam folder
OUTLOOK_FOLDER_CALENDAR = 9    # Calendar folder for appointments
OUTLOOK_FOLDER_CONTACTS = 10   # Contacts folder

# Outlook item type constants (for CreateItem method)
OUTLOOK_ITEM_MAIL = 0          # Mail item type
OUTLOOK_ITEM_APPOINTMENT = 1   # Calendar appointment type
OUTLOOK_ITEM_CONTACT = 2       # Contact item type

# Email importance level constants
IMPORTANCE_LOW = 0
IMPORTANCE_NORMAL = 1
IMPORTANCE_HIGH = 2

# Default limits to prevent performance issues with large mailboxes
DEFAULT_EMAIL_LIMIT = 5            # Reduced from 10 to minimize Outlook freezing
MAX_EMAIL_LIMIT = 50               # Reduced from 100 to prevent long freezes
DEFAULT_CONTACT_LIMIT = 50
MAX_CONTACT_LIMIT = 200            # Reasonable limit for contact queries
EMAIL_BODY_PREVIEW_LENGTH = 500    # Truncate email bodies to prevent excessive data transfer
DEFAULT_DAYS_BACK = 2              # Only search emails from last 2 days by default (ultra-fast!)

# Excluded stores/folders (team mailboxes, shared mailboxes, etc.)
# These will be skipped when listing folders or searching
# Add your specific folders to exclude here if needed
EXCLUDED_STORES = [
    # Example: "Team Mailbox Name",
]


# ============================================================================
# PERFORMANCE CACHE
# ============================================================================
# Performance optimization: Cache frequently accessed folder paths
# This avoids the expensive traversal of all Outlook stores on every request
_FOLDER_CACHE: Dict[str, Any] = {}  # folder_path -> Outlook Folder object




# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def _set_email_body(mail_item, body: str, html_body: Optional[str] = None, signature_name: Optional[str] = None):
    """
    Set email body with optional signature.
    
    Helper function to avoid code duplication between send_email and create_draft_email.
    
    Args:
        mail_item: Outlook MailItem COM object
        body: Plain text body content
        html_body: Optional HTML body content
        signature_name: Optional signature name to add via Display()
        
    Notes:
        - If signature_name is provided, calls Display(False) to let Outlook add the signature
        - This preserves user's Outlook format settings and signature images
        - Outlook automatically adds ~2 blank lines before the signature (native behavior)
    """
    if signature_name:
        try:
            # Force HTML format (2 = olFormatHTML)
            mail_item.BodyFormat = 2
            
            # Set the body content BEFORE Display()
            # This allows Outlook to apply user's default format settings
            if html_body:
                mail_item.HTMLBody = html_body
            else:
                mail_item.Body = body
            
            # Display to add signature with user's format and embedded images
            # Note: Outlook automatically adds spacing before the signature
            mail_item.Display(False)  # False = don't show window
            
        except Exception:
            # Fallback: set body directly without signature
            if html_body:
                mail_item.HTMLBody = html_body
            else:
                mail_item.Body = body
    elif html_body:
        mail_item.HTMLBody = html_body
    else:
        mail_item.Body = body


def get_outlook_application():
    """
    Get or create an instance of the Outlook application via COM.
    
    This function establishes a connection to the local Outlook application
    using Windows COM automation. It's used by all other functions to interact
    with Outlook data.
    
    Returns:
        win32com.client.CDispatch: Outlook Application COM object
        
    Raises:
        ValueError: If Outlook is not installed or cannot be accessed via COM
        
    Notes:
        - Requires Microsoft Outlook to be installed on the system
        - The Outlook application must be properly configured with at least one profile
        - This uses late binding (Dispatch) rather than early binding for compatibility
    """
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        return outlook
    except Exception as e:
        # Log the error with full context for debugging
        logger.error("Failed to connect to Outlook Application", exc_info=True, extra={
            "error_type": type(e).__name__,
            "error_message": str(e)
        })
        raise ValueError(
            f"Unable to connect to Outlook. Make sure Outlook is installed and properly configured. Error: {e}"
        )


def _get_folder_by_path(namespace, folder_path: str, use_cache: bool = True):
    """
    Get an Outlook folder by its path with caching support.
    
    Performance optimization: This function caches folder objects to avoid
    the expensive traversal of all Outlook stores on every request.
    
    Args:
        namespace: Outlook MAPI namespace object
        folder_path: Full path to the folder (e.g., "Inbox/Archive" or "Personal/My Mails")
        use_cache: Whether to use the folder cache (default: True)
        
    Returns:
        Outlook Folder object if found, None otherwise
        
    Notes:
        - First access to a folder may take 20-45 seconds (store traversal)
        - Subsequent accesses use cache and take ~0.01 seconds
        - Cache is invalidated when Outlook is restarted
    """
    # Check cache first
    if use_cache and folder_path in _FOLDER_CACHE:
        try:
            # Verify cached folder is still valid
            _ = _FOLDER_CACHE[folder_path].Name
            return _FOLDER_CACHE[folder_path]
        except Exception:
            # Cache entry is stale, remove it
            del _FOLDER_CACHE[folder_path]
    
    # Search for folder
    folder_parts = folder_path.split('/')
    target_folder = None
    
    # Search through all stores to find the folder (excluding team/shared mailboxes)
    for store in namespace.Stores:
        try:
            # Skip excluded stores (team mailboxes, shared mailboxes)
            if store.DisplayName in EXCLUDED_STORES:
                continue
            
            current_folder = store.GetRootFolder()
            
            # Navigate through the folder path
            for part in folder_parts:
                found = False
                for subfolder in current_folder.Folders:
                    if subfolder.Name == part:
                        current_folder = subfolder
                        found = True
                        break
                
                if not found:
                    break
            else:
                # Successfully found the folder
                target_folder = current_folder
                break
                
        except Exception as e:
            continue
    
    # Cache the result if found
    if target_folder is not None and use_cache:
        _FOLDER_CACHE[folder_path] = target_folder
    
    return target_folder


def format_email(mail_item) -> Dict[str, Any]:
    """
    Format an Outlook mail item as a dictionary for JSON serialization.
    
    Args:
        mail_item: Outlook MailItem COM object
        
    Returns:
        Dict[str, Any]: Dictionary containing email properties
        
    Notes:
        - Email body is truncated to EMAIL_BODY_PREVIEW_LENGTH characters to prevent
          excessive data transfer and potential memory issues
        - Email body is truncated for security and performance
        - Returns an error dict if formatting fails to allow graceful degradation
    """
    try:
        # Truncate body to prevent excessive data exposure
        email_body = mail_item.Body if mail_item.Body else ""
        truncated_body = email_body[:EMAIL_BODY_PREVIEW_LENGTH] + "..." \
                        if len(email_body) > EMAIL_BODY_PREVIEW_LENGTH else email_body
        
        # Format attachments list
        attachments = []
        try:
            for attachment in mail_item.Attachments:
                attachments.append({
                    "filename": attachment.FileName,
                    "size": attachment.Size,
                    "type": attachment.Type  # 1=File, 5=EmbeddedItem, 6=OLE
                })
        except Exception:
            pass
        
        return {
            "subject": mail_item.Subject,
            "sender": mail_item.SenderName,
            "sender_email": mail_item.SenderEmailAddress,
            "recipients": mail_item.To,
            "cc": mail_item.CC,
            "bcc": mail_item.BCC,
            "received_time": str(mail_item.ReceivedTime) if hasattr(mail_item, 'ReceivedTime') else None,
            "sent_on": str(mail_item.SentOn) if hasattr(mail_item, 'SentOn') else None,
            "body": truncated_body,
            "body_length": len(email_body),
            "has_attachments": mail_item.Attachments.Count > 0,
            "attachment_count": mail_item.Attachments.Count,
            "attachments": attachments,
            "importance": mail_item.Importance,
            "unread": mail_item.UnRead,
            "categories": mail_item.Categories,
            "entry_id": mail_item.EntryID if hasattr(mail_item, 'EntryID') else None,
        }
    except Exception as e:
        logger.error("Failed to format email item", exc_info=True, extra={
            "error_type": type(e).__name__
        })
        return {"error": f"Failed to format email: {e}"}


def format_appointment(appointment) -> Dict[str, Any]:
    """
    Format an Outlook appointment/calendar event as a dictionary for JSON serialization.
    
    Args:
        appointment: Outlook AppointmentItem COM object
        
    Returns:
        Dict[str, Any]: Dictionary containing appointment properties
        
    Notes:
        - Body is truncated for the same security reasons as emails
        - BusyStatus codes: 0=Free, 1=Tentative, 2=Busy, 3=Out of Office
    """
    try:
        # Truncate body to prevent excessive data exposure
        appointment_body = appointment.Body if appointment.Body else ""
        truncated_body = appointment_body[:EMAIL_BODY_PREVIEW_LENGTH] + "..." \
                        if len(appointment_body) > EMAIL_BODY_PREVIEW_LENGTH else appointment_body
        
        return {
            "subject": appointment.Subject,
            "start": str(appointment.Start),
            "end": str(appointment.End),
            "location": appointment.Location,
            "organizer": appointment.Organizer if hasattr(appointment, 'Organizer') else None,
            "required_attendees": appointment.RequiredAttendees,
            "optional_attendees": appointment.OptionalAttendees,
            "body": truncated_body,
            "is_all_day_event": appointment.AllDayEvent,
            "reminder_set": appointment.ReminderSet,
            "reminder_minutes": appointment.ReminderMinutesBeforeStart if appointment.ReminderSet else None,
            "categories": appointment.Categories,
            "busy_status": appointment.BusyStatus,
        }
    except Exception as e:
        logger.error("Failed to format appointment", exc_info=True, extra={
            "error_type": type(e).__name__
        })
        return {"error": f"Failed to format appointment: {e}"}


def format_contact(contact) -> Dict[str, Any]:
    """
    Format an Outlook contact as a dictionary for JSON serialization.
    
    Args:
        contact: Outlook ContactItem COM object
        
    Returns:
        Dict[str, Any]: Dictionary containing contact properties
        
    Notes:
        - Uses safe_get helper to handle missing or null properties gracefully
        - Some Outlook contacts may have incomplete data, this ensures robust handling
    """
    try:
        # Safely get attributes with fallback to empty string
        # This is necessary because Outlook contacts can have incomplete data
        def safe_get(obj, attr, default=""):
            """
            Safely retrieve an attribute from a COM object.
            
            Args:
                obj: COM object to retrieve attribute from
                attr: Attribute name to retrieve
                default: Default value if attribute is missing or None
                
            Returns:
                Attribute value or default
            """
            try:
                value = getattr(obj, attr, default)
                return value if value is not None else default
            except Exception:
                # Silently return default if attribute access fails
                # This is expected for some Outlook contact properties
                return default
        
        return {
            "full_name": safe_get(contact, "FullName"),
            "email1": safe_get(contact, "Email1Address"),
            "email2": safe_get(contact, "Email2Address"),
            "email3": safe_get(contact, "Email3Address"),
            "company": safe_get(contact, "CompanyName"),
            "job_title": safe_get(contact, "JobTitle"),
            "business_phone": safe_get(contact, "BusinessTelephoneNumber"),
            "mobile_phone": safe_get(contact, "MobileTelephoneNumber"),
            "home_phone": safe_get(contact, "HomeTelephoneNumber"),
            "business_address": safe_get(contact, "BusinessAddress"),
            "categories": safe_get(contact, "Categories"),
        }
    except Exception as e:
        logger.error("Failed to format contact", exc_info=True, extra={
            "error_type": type(e).__name__
        })
        return {"error": f"Failed to format contact: {e}"}


# ============================================================================
# EMAIL TOOLS
# ============================================================================

@mcp.tool()
def get_inbox_emails(limit: int = DEFAULT_EMAIL_LIMIT, unread_only: bool = False) -> str:
    """
    Get emails from the Outlook Inbox folder.
    
    This function retrieves emails from the user's Inbox, sorted by received time
    (most recent first). It can optionally filter to only show unread emails.
    
    Args:
        limit: Maximum number of emails to return (default: 10, max: 100)
        unread_only: If True, only return unread emails (default: False)
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "count": int,
            "emails": [list of email dictionaries]
        }
        
    Examples:
        >>> get_inbox_emails(limit=5)
        {"success": true, "count": 5, "emails": [...]}
        
        >>> get_inbox_emails(limit=10, unread_only=True)
        {"success": true, "count": 3, "emails": [...]}
        
    Notes:
        - Limited to MAX_EMAIL_LIMIT (50) to prevent performance issues
        - When unread_only=True, we fetch up to limit*2 items to ensure enough results
    """
    try:
        outlook = get_outlook_application()
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(OUTLOOK_FOLDER_INBOX)
        
        # Apply limit cap to prevent performance degradation
        limit = min(limit, MAX_EMAIL_LIMIT)
        
        emails = []
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)  # Sort by received time, descending (newest first)
        
        # PERFORMANCE OPTIMIZATION: Use Restrict() to filter server-side instead of items.Count
        # items.Count can take several minutes on large mailboxes
        if unread_only:
            items = items.Restrict("[Unread] = True")
        
        # PERFORMANCE OPTIMIZATION: Use GetFirst()/GetNext() instead of index access
        # This avoids the expensive items.Count call and is much faster
        mail = items.GetFirst()
        count = 0
        
        while mail is not None and count < limit:
            try:
                emails.append(format_email(mail))
                count += 1
            except Exception as e:
                pass
            
            mail = items.GetNext()
        
        
        return json.dumps({
            "success": True,
            "count": len(emails),
            "emails": emails
        }, indent=2)
        
    except Exception as e:
        logger.error("Failed to get inbox emails", exc_info=True, extra={
            "limit": limit,
            "unread_only": unread_only
        })
        return json.dumps({"success": False, "error": str(e)})


@mcp.tool()
def get_sent_emails(limit: int = DEFAULT_EMAIL_LIMIT) -> str:
    """
    Get emails from the Outlook Sent Items folder.
    
    Retrieves emails that the user has sent, sorted by send time (most recent first).
    Useful for reviewing sent correspondence or finding previously sent information.
    
    Args:
        limit: Maximum number of emails to return (default: 10, max: 100)
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "count": int,
            "emails": [list of email dictionaries]
        }
        
    Examples:
        >>> get_sent_emails(limit=5)
        {"success": true, "count": 5, "emails": [...]}
        
    Notes:
        - Limited to MAX_EMAIL_LIMIT (50) for performance
        - Sorted by SentOn date in descending order
    """
    try:
        outlook = get_outlook_application()
        namespace = outlook.GetNamespace("MAPI")
        sent_folder = namespace.GetDefaultFolder(OUTLOOK_FOLDER_SENT)
        
        # Apply limit cap to prevent performance issues
        limit = min(limit, MAX_EMAIL_LIMIT)
        
        emails = []
        items = sent_folder.Items
        items.Sort("[SentOn]", True)  # Sort by sent time, descending (newest first)
        
        # PERFORMANCE OPTIMIZATION: Use GetFirst()/GetNext() instead of items.Count
        mail = items.GetFirst()
        count = 0
        
        while mail is not None and count < limit:
            try:
                emails.append(format_email(mail))
                count += 1
            except Exception as e:
                pass
            
            mail = items.GetNext()
        
        
        return json.dumps({
            "success": True,
            "count": len(emails),
            "emails": emails
        }, indent=2)
        
    except Exception as e:
        logger.error("Failed to get sent emails", exc_info=True, extra={
            "limit": limit
        })
        return json.dumps({"success": False, "error": str(e)})


@mcp.tool()
def search_emails(query: str, folder: str = "inbox", limit: int = 20) -> str:
    """
    Search for emails in Outlook folders using keyword matching.
    
    Searches across subject, body, and sender name fields. Can search in a specific
    folder or across all mail folders.
    
    Args:
        query: Search query (searches in subject, body, sender)
        folder: Folder to search in (inbox, sent, drafts, deleted, all) (default: inbox)
        limit: Maximum number of results (default: 20, max: 100)
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "query": str,
            "count": int,
            "emails": [list of matching email dictionaries]
        }
        
    Examples:
        >>> search_emails("payment", folder="inbox", limit=10)
        {"success": true, "query": "payment", "count": 5, "emails": [...]}
        
        >>> search_emails("project update", folder="all", limit=20)
        {"success": true, "query": "project update", "count": 15, "emails": [...]}
        
    Notes:
        - Uses Outlook's SQL-like filter syntax for efficient searching
        - Limited to MAX_EMAIL_LIMIT (50) for performance
        - When folder="all", searches inbox, sent, and drafts folders
    """
    try:
        outlook = get_outlook_application()
        namespace = outlook.GetNamespace("MAPI")
        
        # Map folder names to Outlook folder constants
        folder_map = {
            "inbox": OUTLOOK_FOLDER_INBOX,
            "sent": OUTLOOK_FOLDER_SENT,
            "drafts": OUTLOOK_FOLDER_DRAFTS,
            "deleted": OUTLOOK_FOLDER_DELETED,
        }
        
        # Apply limit cap to prevent performance issues
        limit = min(limit, MAX_EMAIL_LIMIT)
        emails = []
        
        # Determine which folders to search
        if folder == "all":
            # Search across multiple folders for comprehensive results
            folders_to_search = [OUTLOOK_FOLDER_INBOX, OUTLOOK_FOLDER_SENT, OUTLOOK_FOLDER_DRAFTS]
        else:
            folder_id = folder_map.get(folder.lower(), OUTLOOK_FOLDER_INBOX)
            folders_to_search = [folder_id]
        
        for folder_id in folders_to_search:
            search_folder = namespace.GetDefaultFolder(folder_id)
            
            # Build Outlook SQL filter for searching
            # Uses Outlook's DASL (DAV Searching and Locating) query syntax
            # This is more efficient than iterating through all items
            filter_str = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{query}%' OR " \
                        f"\"urn:schemas:httpmail:textdescription\" LIKE '%{query}%' OR " \
                        f"\"urn:schemas:httpmail:fromname\" LIKE '%{query}%'\""
            
            items = search_folder.Items.Restrict(filter_str)
            items.Sort("[ReceivedTime]", True)  # Sort by received time, descending
            
            # PERFORMANCE OPTIMIZATION: Use GetFirst()/GetNext() instead of items.Count
            remaining_limit = limit - len(emails)
            mail = items.GetFirst()
            count = 0
            
            while mail is not None and count < remaining_limit:
                try:
                    emails.append(format_email(mail))
                    count += 1
                    
                    if len(emails) >= limit:
                        break
                except Exception as e:
                    pass
                
                mail = items.GetNext()
            
            # Stop searching other folders if we reached the limit
            if len(emails) >= limit:
                break
        
        
        return json.dumps({
            "success": True,
            "query": query,
            "count": len(emails),
            "emails": emails
        }, indent=2)
        
    except Exception as e:
        logger.error("Failed to search emails", exc_info=True, extra={
            "query": query,
            "folder": folder,
            "limit": limit
        })
        return json.dumps({"success": False, "error": str(e)})


@mcp.tool()
def send_email(
    to: str,
    subject: str,
    body: str,
    cc: Optional[str] = None,
    bcc: Optional[str] = None,
    importance: str = "normal",
    html_body: Optional[str] = None,
    signature_name: Optional[str] = None
) -> str:
    """
    Send an email via Outlook.
    
    Creates and sends a new email through the user's Outlook account. The email
    is sent immediately and a copy is saved in the Sent Items folder.
    
    Args:
        to: Recipient email address(es), semicolon-separated for multiple
            Example: "user1@example.com" or "user1@example.com; user2@example.com"
        subject: Email subject line
        body: Email body content (plain text format)
        cc: CC recipients (optional), semicolon-separated
        bcc: BCC recipients (optional), semicolon-separated
        importance: Email importance level (low, normal, high) (default: normal)
        html_body: HTML body content (optional). If provided, this will be used instead of body
        signature_name: Name of Outlook signature to use (optional). If provided, Outlook will add the signature
            automatically by calling Display(False). This preserves user's Outlook settings and signature images.
            Note: Outlook adds ~2 blank lines before the signature (native behavior)
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "message": str
        }
        
    Examples:
        >>> send_email("colleague@company.com", "Meeting", "See you at 2pm")
        {"success": true, "message": "Email sent to colleague@company.com"}
        
        >>> send_email("team@company.com", "Urgent", "...", importance="high")
        {"success": true, "message": "Email sent to team@company.com"}
        
        >>> send_email("user@company.com", "Hello", "Message", signature_name="VP DXT")
        {"success": true, "message": "Email sent to user@company.com"}
        
    Security Notes:
        - No sensitive data should be included in logs
        - Recipient addresses are logged but email content is not
        - BCC recipients are never logged for privacy
    """
    try:
        outlook = get_outlook_application()
        mail = outlook.CreateItem(OUTLOOK_ITEM_MAIL)
        
        mail.To = to
        mail.Subject = subject
        
        # Set email body (with optional signature)
        _set_email_body(mail, body, html_body, signature_name)
        
        # Set CC and BCC recipients
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        
        # Set importance level using named constants
        importance_map = {
            "low": IMPORTANCE_LOW,
            "normal": IMPORTANCE_NORMAL,
            "high": IMPORTANCE_HIGH
        }
        mail.Importance = importance_map.get(importance.lower(), IMPORTANCE_NORMAL)
        
        # Send the email
        mail.Send()
        
        return json.dumps({
            "success": True,
            "message": f"Email sent to {to}"
        }, indent=2)
        
    except Exception as e:
        logger.error("Failed to send email", exc_info=True, extra={
            "to": to,
            "subject": subject,
            "importance": importance
        })
        return json.dumps({"success": False, "error": str(e)})


@mcp.tool()
def create_draft_email(
    to: str,
    subject: str,
    body: str,
    cc: Optional[str] = None,
    bcc: Optional[str] = None,
    html_body: Optional[str] = None,
    signature_name: Optional[str] = None
) -> str:
    """
    Create a draft email in Outlook without sending it.
    
    Creates an email and saves it to the Drafts folder where the user can
    review, edit, and send it later. Useful for preparing emails that need
    review before sending.
    
    Args:
        to: Recipient email address(es), semicolon-separated for multiple
        subject: Email subject line
        body: Email body content (plain text format)
        cc: CC recipients (optional), semicolon-separated
        bcc: BCC recipients (optional), semicolon-separated
        html_body: HTML body content (optional). If provided, this will be used instead of body
        signature_name: Name of Outlook signature to use (optional). If provided, Outlook will add the signature
            automatically by calling Display(False). This preserves user's Outlook settings and signature images.
            Note: Outlook adds ~2 blank lines before the signature (native behavior)
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "message": str
        }
        
    Examples:
        >>> create_draft_email("manager@company.com", "Report", "Draft report...")
        {"success": true, "message": "Draft email created"}
        
        >>> create_draft_email("user@company.com", "Hello", "Message", signature_name="VP DXT")
        {"success": true, "message": "Draft email created"}
        
    Notes:
        - Draft is saved in the user's Drafts folder
        - User can find and edit the draft in Outlook
        - No email is sent until the user manually sends it
    """
    try:
        outlook = get_outlook_application()
        mail = outlook.CreateItem(OUTLOOK_ITEM_MAIL)
        
        mail.To = to
        mail.Subject = subject
        
        # Set email body (with optional signature)
        _set_email_body(mail, body, html_body, signature_name)
        
        # Set CC and BCC recipients
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        
        # Save as draft (does not send)
        mail.Save()
        
        return json.dumps({
            "success": True,
            "message": "Draft email created"
        }, indent=2)
        
    except Exception as e:
        logger.error("Failed to create draft email", exc_info=True, extra={
            "to": to,
            "subject": subject
        })
        return json.dumps({"success": False, "error": str(e)})


@mcp.tool()
def get_email_attachments(entry_id: str) -> str:
    """
    Get list of attachments from a specific email.
    
    Retrieves detailed information about all attachments in an email,
    including filename, size, and type.
    
    Args:
        entry_id: Email's EntryID (obtained from get_inbox_emails, search_emails, etc.)
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "count": int,
            "attachments": [
                {
                    "filename": str,
                    "size": int (bytes),
                    "type": int,
                    "index": int
                }
            ]
        }
        
    Examples:
        >>> get_email_attachments("00000000...")
        {"success": true, "count": 2, "attachments": [...]}
        
    Notes:
        - Type: 1=File, 5=EmbeddedItem, 6=OLE
        - Use the index to download specific attachments with download_email_attachment
    """
    try:
        outlook = get_outlook_application()
        namespace = outlook.GetNamespace("MAPI")
        
        # Get the email by EntryID
        mail_item = namespace.GetItemFromID(entry_id)
        
        attachments = []
        for i, attachment in enumerate(mail_item.Attachments, start=1):
            attachments.append({
                "filename": attachment.FileName,
                "size": attachment.Size,
                "type": attachment.Type,
                "index": i
            })
        
        return json.dumps({
            "success": True,
            "count": len(attachments),
            "attachments": attachments
        }, indent=2)
        
    except Exception as e:
        logger.error("Failed to get email attachments", exc_info=True, extra={
            "entry_id": entry_id
        })
        return json.dumps({"success": False, "error": str(e)})


@mcp.tool()
def download_email_attachment(entry_id: str, attachment_index: int, save_path: str) -> str:
    """
    Download a specific attachment from an email to disk.
    
    Downloads the specified attachment and saves it to the given path.
    
    Args:
        entry_id: Email's EntryID
        attachment_index: Index of the attachment (1-based, from get_email_attachments)
        save_path: Full path where to save the attachment (e.g., "C:/Users/user/Downloads/file.pdf")
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "message": str,
            "saved_path": str
        }
        
    Examples:
        >>> download_email_attachment("00000000...", 1, "C:/Downloads/report.pdf")
        {"success": true, "message": "Attachment downloaded", "saved_path": "C:/Downloads/report.pdf"}
        
    Notes:
        - Creates parent directories if they don't exist
        - Overwrites existing files without warning
        - Ensure you have write permissions to the target directory
    """
    try:
        import os
        
        outlook = get_outlook_application()
        namespace = outlook.GetNamespace("MAPI")
        
        # Get the email by EntryID
        mail_item = namespace.GetItemFromID(entry_id)
        
        # Check attachment index is valid
        if attachment_index < 1 or attachment_index > mail_item.Attachments.Count:
            return json.dumps({
                "success": False,
                "error": f"Invalid attachment index {attachment_index}. Email has {mail_item.Attachments.Count} attachment(s)"
            })
        
        # Get the attachment (COM collections are 1-indexed)
        attachment = mail_item.Attachments[attachment_index]
        
        # Create parent directories if they don't exist
        os.makedirs(os.path.dirname(os.path.abspath(save_path)), exist_ok=True)
        
        # Save the attachment
        attachment.SaveAsFile(save_path)
        
        return json.dumps({
            "success": True,
            "message": f"Attachment '{attachment.FileName}' downloaded successfully",
            "saved_path": save_path,
            "filename": attachment.FileName,
            "size": attachment.Size
        }, indent=2)
        
    except Exception as e:
        logger.error("Failed to download attachment", exc_info=True, extra={
            "entry_id": entry_id,
            "attachment_index": attachment_index,
            "save_path": save_path
        })
        return json.dumps({"success": False, "error": str(e)})


@mcp.tool()
def send_email_with_attachments(
    to: str,
    subject: str,
    body: str,
    attachments: str,
    cc: Optional[str] = None,
    bcc: Optional[str] = None,
    importance: str = "normal",
    html_body: Optional[str] = None,
    signature_name: Optional[str] = None
) -> str:
    """
    Send an email with file attachments via Outlook.
    
    Creates and sends an email with one or more file attachments.
    
    Args:
        to: Recipient email address(es), semicolon-separated
        subject: Email subject line
        body: Email body content (plain text format)
        attachments: File path(s) to attach, semicolon-separated
            Example: "C:/file1.pdf" or "C:/file1.pdf; C:/file2.docx"
        cc: CC recipients (optional), semicolon-separated
        bcc: BCC recipients (optional), semicolon-separated
        importance: Email importance level (low, normal, high) (default: normal)
        html_body: HTML body content (optional)
        signature_name: Name of Outlook signature to use (optional)
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "message": str,
            "attachments_added": int
        }
        
    Examples:
        >>> send_email_with_attachments(
        ...     "colleague@company.com",
        ...     "Report",
        ...     "Please find attached",
        ...     "C:/Users/user/report.pdf"
        ... )
        {"success": true, "message": "Email sent", "attachments_added": 1}
        
    Notes:
        - All attachment file paths must exist and be accessible
        - Files are attached as-is (no compression or encoding)
        - Large attachments may take time to send
    """
    try:
        import os
        
        outlook = get_outlook_application()
        mail = outlook.CreateItem(OUTLOOK_ITEM_MAIL)
        
        mail.To = to
        mail.Subject = subject
        
        # Set email body (with optional signature)
        _set_email_body(mail, body, html_body, signature_name)
        
        # Set CC and BCC recipients
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        
        # Set importance level
        importance_map = {
            "low": IMPORTANCE_LOW,
            "normal": IMPORTANCE_NORMAL,
            "high": IMPORTANCE_HIGH
        }
        mail.Importance = importance_map.get(importance.lower(), IMPORTANCE_NORMAL)
        
        # Add attachments
        attachment_paths = [path.strip() for path in attachments.split(';')]
        attachments_added = 0
        missing_files = []
        
        for path in attachment_paths:
            if not path:
                continue
                
            if not os.path.exists(path):
                missing_files.append(path)
                continue
            
            try:
                mail.Attachments.Add(path)
                attachments_added += 1
            except Exception as e:
                logger.error(f"Failed to attach file: {path}", exc_info=True)
                missing_files.append(f"{path} (error: {e})")
        
        if missing_files:
            return json.dumps({
                "success": False,
                "error": f"Some attachment files not found or couldn't be attached: {', '.join(missing_files)}"
            })
        
        # Send the email
        mail.Send()
        
        return json.dumps({
            "success": True,
            "message": f"Email sent to {to}",
            "attachments_added": attachments_added
        }, indent=2)
        
    except Exception as e:
        logger.error("Failed to send email with attachments", exc_info=True, extra={
            "to": to,
            "subject": subject
        })
        return json.dumps({"success": False, "error": str(e)})


# ============================================================================
# FOLDER MANAGEMENT TOOLS
# ============================================================================

def _get_all_folders(folder, folder_list=None, parent_path="", include_counts=False):
    """
    Recursively get all folders in Outlook.
    
    Helper function to traverse the Outlook folder hierarchy and build
    a flat list of all folders with their full paths.
    
    Args:
        folder: Outlook folder COM object to start from
        folder_list: List to accumulate folders (used in recursion)
        parent_path: Path of parent folders (used in recursion)
        include_counts: Whether to include item/unread counts (SLOW! default: False)
        
    Returns:
        List of dictionaries containing folder information
        
    Notes:
        - Uses recursion to traverse nested folder structures
        - Builds full paths like "Inbox/Archive/2024"
        - Some system folders may not be accessible (handled gracefully)
        - PERFORMANCE: include_counts=True can take minutes on large mailboxes!
    """
    if folder_list is None:
        folder_list = []
    
    try:
        # Build the full path for this folder
        current_path = f"{parent_path}/{folder.Name}" if parent_path else folder.Name
        
        # Build folder info (optionally without expensive counts)
        folder_info = {
            "name": folder.Name,
            "path": current_path
        }
        
        # Performance optimization: Only get counts if explicitly requested
        # folder.Items.Count can take several minutes on large mailboxes
        if include_counts:
            try:
                folder_info["item_count"] = folder.Items.Count if hasattr(folder, 'Items') else 0
                folder_info["unread_count"] = folder.UnReadItemCount if hasattr(folder, 'UnReadItemCount') else 0
            except Exception:
                folder_info["item_count"] = -1  # Indicates error/unavailable
                folder_info["unread_count"] = -1
        
        folder_list.append(folder_info)
        
        # Recursively process subfolders
        # Use try/except to handle Folders.Count gracefully
        try:
            if hasattr(folder, 'Folders'):
                # Use iterator instead of Count where possible
                for subfolder in folder.Folders:
                    _get_all_folders(subfolder, folder_list, current_path, include_counts)
        except Exception as e:
            pass
        
    except Exception as e:
        # Some system folders may throw errors when accessed
        pass
    
    return folder_list


@mcp.tool()
def list_outlook_folders() -> str:
    """
    List all available Outlook folders with their paths (FAST - no item counts).
    
    Retrieves a hierarchical list of all mail folders in Outlook, including
    default folders (Inbox, Sent, etc.) and custom user-created folders.
    Useful for discovering folder names before searching.
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "count": int,
            "folders": [
                {
                    "name": str,
                    "path": str
                }
            ]
        }
        
    Examples:
        >>> list_outlook_folders()
        {
            "success": true,
            "count": 25,
            "folders": [
                {"name": "Inbox", "path": "Inbox"},
                {"name": "Archive", "path": "Inbox/Archive"},
                {"name": "Personal", "path": "Personal"},
                {"name": "My Mails", "path": "Personal/My Mails"}
            ]
        }
        
    Notes:
        - Performance optimization: Does NOT include item_count/unread_count by default
        - These counts can take several minutes on large mailboxes
        - Includes all folders recursively (nested folders)
        - System folders that can't be accessed are skipped
        - Useful to find the exact folder name/path for searching
        - Returns in seconds instead of minutes!
    """
    try:
        outlook = get_outlook_application()
        namespace = outlook.GetNamespace("MAPI")
        
        # Get all folders starting from the root store
        all_folders = []
        
        # Iterate through all stores (accounts), excluding team/shared mailboxes
        for store in namespace.Stores:
            try:
                # Skip excluded stores (team mailboxes, shared mailboxes)
                if store.DisplayName in EXCLUDED_STORES:
                    continue
                
                root_folder = store.GetRootFolder()
                # Performance: include_counts=False for speed
                store_folders = _get_all_folders(root_folder, include_counts=False)
                all_folders.extend(store_folders)
            except Exception as e:
                pass
        
        return json.dumps({
            "success": True,
            "count": len(all_folders),
            "folders": all_folders
        }, indent=2)
        
    except Exception as e:
        logger.error("Failed to list Outlook folders", exc_info=True)
        return json.dumps({"success": False, "error": str(e)})


@mcp.tool()
def search_emails_in_custom_folder(
    folder_path: str,
    query: Optional[str] = None,
    limit: int = 20,
    days_back: int = DEFAULT_DAYS_BACK
) -> str:
    """
    Search for emails in a specific custom Outlook folder.
    
    Allows searching in user-created folders or any folder by its path.
    Can retrieve all emails or filter by keyword.
    
    Args:
        folder_path: Full path to the folder (use list_outlook_folders to find paths)
            Examples: "Personal", "Inbox/Archive", "Projects/2024"
        query: Optional search keyword (searches subject, body, sender)
            If None, returns all emails in the folder
        limit: Maximum number of results (default: 20, max: 50)
        days_back: Number of days back to search (default: 7)
            Only searches emails received in the last N days to improve performance
            Set to 0 or negative to search ALL emails (slower, may freeze Outlook)
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "folder": str,
            "query": str (optional),
            "count": int,
            "emails": [list of email dictionaries],
            "days_back": int (if filtered by date)
        }
        
    Examples:
        >>> search_emails_in_custom_folder("Personal", "invoice")
        {"success": true, "folder": "Personal", "query": "invoice", ...}
        
        >>> search_emails_in_custom_folder("Inbox/Archive/2024", limit=50, days_back=30)
        {"success": true, "folder": "Inbox/Archive/2024", "count": 50, "days_back": 30, ...}
        
    Notes:
        - Use list_outlook_folders() first to discover available folder paths
        - Folder path must match exactly (case-sensitive)
        - Limited to MAX_EMAIL_LIMIT (50) and 2 days by default for performance
        - Searching ALL emails (days_back <= 0) can freeze Outlook for minutes!
    """
    try:
        outlook = get_outlook_application()
        namespace = outlook.GetNamespace("MAPI")
        
        # Apply limit cap to prevent performance issues
        limit = min(limit, MAX_EMAIL_LIMIT)
        
        # Find the folder by path (with caching for performance)
        target_folder = _get_folder_by_path(namespace, folder_path, use_cache=True)
        
        if target_folder is None:
            return json.dumps({
                "success": False,
                "error": f"Folder '{folder_path}' not found. Use list_outlook_folders() to see available folders."
            })
        
        # Get items from the folder
        items = target_folder.Items
        
        # PERFORMANCE OPTIMIZATION: Filter by date BEFORE iterating
        # This reduces Outlook freezing from minutes to seconds!
        if days_back > 0:
            start_date = datetime.now() - timedelta(days=days_back)
            filter_str = f"[ReceivedTime] >= '{start_date.strftime('%m/%d/%Y')}'"
            items = items.Restrict(filter_str)
        
        items.Sort("[ReceivedTime]", True)  # Sort by received time, descending
        
        emails = []
        
        # PERFORMANCE OPTIMIZATION: Use direct indexing instead of GetFirst/GetNext
        # Direct indexing is 5-10x faster on large folders!
        # We iterate up to a reasonable max without calling items.Count (which is slow)
        max_index = limit * 5 if query else limit  # Scan more items when searching with query
        
        if query:
            query_lower = query.lower()
            
            # Search with query filter using direct indexing
            for i in range(max_index):
                try:
                    mail = items[i + 1]  # Outlook COM collections are 1-indexed
                    
                    # Check if query matches subject, body, or sender
                    subject = mail.Subject.lower() if mail.Subject else ""
                    body = mail.Body.lower() if mail.Body else ""
                    sender = mail.SenderName.lower() if mail.SenderName else ""
                    
                    if query_lower in subject or query_lower in body or query_lower in sender:
                        emails.append(format_email(mail))
                        
                        if len(emails) >= limit:
                            break
                            
                except Exception as e:
                    # End of collection or error accessing item
                    break
        else:
            # No query - return all emails up to limit using direct indexing
            for i in range(limit):
                try:
                    mail = items[i + 1]  # Outlook COM collections are 1-indexed
                    emails.append(format_email(mail))
                except Exception as e:
                    # End of collection or error accessing item
                    break
        
        
        result = {
            "success": True,
            "folder": folder_path,
            "count": len(emails),
            "emails": emails
        }
        
        if query:
            result["query"] = query
        
        if days_back > 0:
            result["days_back"] = days_back
            result["info"] = f"Searched emails from last {days_back} days only"
        
        return json.dumps(result, indent=2)
        
    except Exception as e:
        logger.error("Failed to search in custom folder", exc_info=True, extra={
            "folder_path": folder_path,
            "query": query,
            "limit": limit
        })
        return json.dumps({"success": False, "error": str(e)})


@mcp.tool()
def list_outlook_rules() -> str:
    """
    List all Outlook rules (mail organization rules).
    
    Retrieves all active and inactive mail rules configured in Outlook,
    including their conditions and actions. Useful for understanding how
    emails are automatically organized.
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "count": int,
            "rules": [
                {
                    "name": str,
                    "enabled": bool,
                    "description": str,
                    "conditions": [list of conditions],
                    "actions": [list of actions],
                    "exceptions": [list of exceptions]
                }
            ]
        }
        
    Examples:
        >>> list_outlook_rules()
        {
            "success": true,
            "count": 3,
            "rules": [
                {
                    "name": "Move DLP emails",
                    "enabled": true,
                    "description": "Move emails containing 'DLP' to folder",
                    "conditions": ["Subject contains 'DLP'"],
                    "actions": ["Move to folder 'Personal/My Mails'"]
                }
            ]
        }
        
    Notes:
        - Shows both enabled and disabled rules
        - Helps understand email organization workflow
        - Rules are logged for audit purposes
    """
    try:
        outlook = get_outlook_application()
        namespace = outlook.GetNamespace("MAPI")
        
        # Get rules from the default store
        rules_collection = namespace.DefaultStore.GetRules()
        
        rules = []
        for rule in rules_collection:
            try:
                rule_info = {
                    "name": rule.Name,
                    "enabled": rule.Enabled,
                    "description": "",
                    "conditions": [],
                    "actions": [],
                    "exceptions": []
                }
                
                # Parse conditions
                conditions = rule.Conditions
                if hasattr(conditions, 'Subject') and conditions.Subject.Enabled:
                    rule_info["conditions"].append(f"Subject contains: {', '.join(conditions.Subject.Text)}")
                
                if hasattr(conditions, 'Body') and conditions.Body.Enabled:
                    rule_info["conditions"].append(f"Body contains: {', '.join(conditions.Body.Text)}")
                
                if hasattr(conditions, 'From') and conditions.From.Enabled:
                    recipients = []
                    for recipient in conditions.From.Recipients:
                        recipients.append(recipient.Name)
                    rule_info["conditions"].append(f"From: {', '.join(recipients)}")
                
                if hasattr(conditions, 'SentTo') and conditions.SentTo.Enabled:
                    recipients = []
                    for recipient in conditions.SentTo.Recipients:
                        recipients.append(recipient.Name)
                    rule_info["conditions"].append(f"Sent to: {', '.join(recipients)}")
                
                if hasattr(conditions, 'CC') and conditions.CC.Enabled:
                    recipients = []
                    for recipient in conditions.CC.Recipients:
                        recipients.append(recipient.Name)
                    rule_info["conditions"].append(f"CC: {', '.join(recipients)}")
                
                if hasattr(conditions, 'Category') and conditions.Category.Enabled:
                    rule_info["conditions"].append(f"Category: {', '.join(conditions.Category.Categories)}")
                
                if hasattr(conditions, 'Importance') and conditions.Importance.Enabled:
                    importance_map = {0: "Low", 1: "Normal", 2: "High"}
                    rule_info["conditions"].append(f"Importance: {importance_map.get(conditions.Importance.Importance, 'Unknown')}")
                
                # Parse actions
                actions = rule.Actions
                if hasattr(actions, 'MoveToFolder') and actions.MoveToFolder.Enabled:
                    try:
                        folder_name = actions.MoveToFolder.Folder.Name
                        # Try to get full path
                        folder_path = folder_name
                        try:
                            parent = actions.MoveToFolder.Folder.Parent
                            path_parts = [folder_name]
                            while parent and hasattr(parent, 'Name'):
                                path_parts.insert(0, parent.Name)
                                parent = parent.Parent if hasattr(parent, 'Parent') else None
                            folder_path = "/".join(path_parts)
                        except Exception:
                            pass
                        rule_info["actions"].append(f"Move to folder: {folder_path}")
                    except Exception:
                        rule_info["actions"].append("Move to folder: (unable to determine)")
                
                if hasattr(actions, 'CopyToFolder') and actions.CopyToFolder.Enabled:
                    try:
                        folder_name = actions.CopyToFolder.Folder.Name
                        rule_info["actions"].append(f"Copy to folder: {folder_name}")
                    except Exception:
                        rule_info["actions"].append("Copy to folder: (unable to determine)")
                
                if hasattr(actions, 'Delete') and actions.Delete.Enabled:
                    rule_info["actions"].append("Delete message")
                
                if hasattr(actions, 'MarkAsRead') and actions.MarkAsRead.Enabled:
                    rule_info["actions"].append("Mark as read")
                
                if hasattr(actions, 'AssignToCategory') and actions.AssignToCategory.Enabled:
                    rule_info["actions"].append(f"Assign category: {', '.join(actions.AssignToCategory.Categories)}")
                
                if hasattr(actions, 'Forward') and actions.Forward.Enabled:
                    recipients = []
                    for recipient in actions.Forward.Recipients:
                        recipients.append(recipient.Name)
                    rule_info["actions"].append(f"Forward to: {', '.join(recipients)}")
                
                if hasattr(actions, 'Redirect') and actions.Redirect.Enabled:
                    recipients = []
                    for recipient in actions.Redirect.Recipients:
                        recipients.append(recipient.Name)
                    rule_info["actions"].append(f"Redirect to: {', '.join(recipients)}")
                
                # Parse exceptions
                exceptions = rule.Exceptions
                if hasattr(exceptions, 'Subject') and exceptions.Subject.Enabled:
                    rule_info["exceptions"].append(f"Except if subject contains: {', '.join(exceptions.Subject.Text)}")
                
                if hasattr(exceptions, 'From') and exceptions.From.Enabled:
                    recipients = []
                    for recipient in exceptions.From.Recipients:
                        recipients.append(recipient.Name)
                    rule_info["exceptions"].append(f"Except from: {', '.join(recipients)}")
                
                # Build description
                desc_parts = []
                if rule_info["conditions"]:
                    desc_parts.append(f"When: {'; '.join(rule_info['conditions'])}")
                if rule_info["actions"]:
                    desc_parts.append(f"Then: {'; '.join(rule_info['actions'])}")
                if rule_info["exceptions"]:
                    desc_parts.append(f"Except: {'; '.join(rule_info['exceptions'])}")
                rule_info["description"] = " | ".join(desc_parts)
                
                rules.append(rule_info)
                
            except Exception as e:
                # Add a minimal entry for this rule
                rules.append({
                    "name": rule.Name if hasattr(rule, 'Name') else "Unknown",
                    "enabled": rule.Enabled if hasattr(rule, 'Enabled') else False,
                    "description": f"Error parsing rule: {e}",
                    "conditions": [],
                    "actions": [],
                    "exceptions": []
                })
        
        
        return json.dumps({
            "success": True,
            "count": len(rules),
            "rules": rules
        }, indent=2)
        
    except Exception as e:
        logger.error("Failed to list Outlook rules", exc_info=True)
        return json.dumps({"success": False, "error": str(e)})


# ============================================================================
# CALENDAR TOOLS
# ============================================================================

@mcp.tool()
def get_calendar_events(days_ahead: int = 7, include_past: bool = False) -> str:
    """
    Get calendar events from Outlook.
    
    Retrieves upcoming calendar appointments and meetings. Can optionally include
    events from earlier today. Handles recurring events automatically.
    
    Args:
        days_ahead: Number of days ahead to fetch events (default: 7)
        include_past: Include past events from today (default: False)
            If True, starts from midnight today; if False, starts from current time
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "count": int,
            "events": [list of event dictionaries]
        }
        
    Examples:
        >>> get_calendar_events(days_ahead=3)
        {"success": true, "count": 8, "events": [...]}
        
        >>> get_calendar_events(days_ahead=1, include_past=True)
        {"success": true, "count": 5, "events": [...]}  # includes today's past events
        
    Notes:
        - IncludeRecurrences=True ensures recurring meetings are expanded
        - Events are sorted by start time
        - Handles all-day events correctly
    """
    try:
        outlook = get_outlook_application()
        namespace = outlook.GetNamespace("MAPI")
        calendar = namespace.GetDefaultFolder(OUTLOOK_FOLDER_CALENDAR)
        
        items = calendar.Items
        items.Sort("[Start]")  # Sort by start time ascending
        items.IncludeRecurrences = True  # Expand recurring events into individual instances
        
        # Build date range for filtering
        start_date = datetime.now()
        if include_past:
            # Start from midnight today to include past events from today
            start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
        
        # End date is end of day N days from now
        end_date = datetime.now().replace(hour=23, minute=59, second=59)
        end_date = end_date + timedelta(days=days_ahead)
        
        # Build Outlook filter string
        # Format must match Outlook's expected date format (MM/DD/YYYY HH:MM)
        filter_str = f"[Start] >= '{start_date.strftime('%m/%d/%Y %H:%M')}' AND [End] <= '{end_date.strftime('%m/%d/%Y %H:%M')}'"
        filtered_items = items.Restrict(filter_str)
        
        events = []
        for appointment in filtered_items:
            events.append(format_appointment(appointment))
        
        
        return json.dumps({
            "success": True,
            "count": len(events),
            "events": events
        }, indent=2)
        
    except Exception as e:
        logger.error("Failed to get calendar events", exc_info=True, extra={
            "days_ahead": days_ahead,
            "include_past": include_past
        })
        return json.dumps({"success": False, "error": str(e)})


@mcp.tool()
def create_calendar_event(
    subject: str,
    start_time: str,
    end_time: str,
    location: Optional[str] = None,
    body: Optional[str] = None,
    required_attendees: Optional[str] = None,
    optional_attendees: Optional[str] = None,
    reminder_minutes: int = 15,
    is_all_day: bool = False
) -> str:
    """
    Create a new calendar event/appointment in Outlook.
    
    Creates a calendar entry and optionally sends meeting invitations to attendees.
    Supports flexible date/time parsing including ISO format and natural language.
    
    Args:
        subject: Event subject/title
        start_time: Start time (ISO format or natural language)
            Examples: "2025-01-15 14:00", "tomorrow 2pm", "next Monday at 9am"
        end_time: End time (same formats as start_time)
        location: Event location (optional)
            Examples: "Conference Room A", "Microsoft Teams Meeting"
        body: Event description/agenda (optional)
        required_attendees: Required attendees, semicolon-separated (optional)
            Example: "user1@company.com; user2@company.com"
        optional_attendees: Optional attendees, semicolon-separated (optional)
        reminder_minutes: Minutes before event to show reminder (default: 15)
        is_all_day: Whether this is an all-day event (default: False)
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "message": str
        }
        
    Examples:
        >>> create_calendar_event("Team Meeting", "2025-12-20 14:00", "2025-12-20 15:00")
        {"success": true, "message": "Calendar event 'Team Meeting' created..."}
        
        >>> create_calendar_event("All Hands", "2025-12-25 09:00", "2025-12-25 10:00",
        ...                       location="Auditorium", required_attendees="team@company.com")
        {"success": true, "message": "Calendar event 'All Hands' created..."}
        
    Notes:
        - If attendees are specified, a meeting invitation is sent automatically
        - Uses python-dateutil for flexible date parsing
        - Meeting requests are sent according to standard Outlook behavior
        - Reminder is enabled by default (productivity best practice)
    """
    try:
        outlook = get_outlook_application()
        appointment = outlook.CreateItem(OUTLOOK_ITEM_APPOINTMENT)
        
        appointment.Subject = subject
        
        # Parse dates using dateutil for flexible parsing
        try:
            start_dt = date_parser.parse(start_time)
            end_dt = date_parser.parse(end_time)
        except Exception as e:
            return json.dumps({
                "success": False,
                "error": f"Invalid date format: {e}. Use ISO format like '2025-01-15 14:00' or natural language like 'tomorrow 2pm'"
            })
        
        appointment.Start = start_dt
        appointment.End = end_dt
        appointment.AllDayEvent = is_all_day
        
        # Set optional properties
        if location:
            appointment.Location = location
        if body:
            appointment.Body = body
        if required_attendees:
            appointment.RequiredAttendees = required_attendees
        if optional_attendees:
            appointment.OptionalAttendees = optional_attendees
        
        # Set reminder (best practice: always set reminders)
        appointment.ReminderSet = True
        appointment.ReminderMinutesBeforeStart = reminder_minutes
        
        # Save the appointment
        appointment.Save()
        
        # If there are attendees, send meeting invitation
        # This converts the appointment to a meeting request
        if required_attendees or optional_attendees:
            appointment.Send()
        
        return json.dumps({
            "success": True,
            "message": f"Calendar event '{subject}' created for {start_time}"
        }, indent=2)
        
    except Exception as e:
        logger.error("Failed to create calendar event", exc_info=True, extra={
            "subject": subject,
            "start_time": start_time,
            "end_time": end_time
        })
        return json.dumps({"success": False, "error": str(e)})


@mcp.tool()
def search_calendar_events(query: str, days_range: int = 30) -> str:
    """
    Search for calendar events by keyword in subject or location.
    
    Searches calendar events within a date range, looking for matches in both
    the event subject and location fields. Useful for finding specific meetings
    or events in a particular location.
    
    Args:
        query: Search keyword (case-insensitive)
            Searches in both subject and location fields
        days_range: Number of days to search (past and future from today) (default: 30)
            Example: days_range=30 searches from 30 days ago to 30 days in the future
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "query": str,
            "count": int,
            "events": [list of matching event dictionaries]
        }
        
    Examples:
        >>> search_calendar_events("standup", days_range=7)
        {"success": true, "query": "standup", "count": 5, "events": [...]}
        
        >>> search_calendar_events("Conference Room", days_range=14)
        {"success": true, "query": "Conference Room", "count": 8, "events": [...]}
        
    Notes:
        - Search is case-insensitive
        - Includes recurring events
        - Searches both past and future events from today
    """
    try:
        outlook = get_outlook_application()
        namespace = outlook.GetNamespace("MAPI")
        calendar = namespace.GetDefaultFolder(OUTLOOK_FOLDER_CALENDAR)
        
        items = calendar.Items
        items.Sort("[Start]")  # Sort by start time
        items.IncludeRecurrences = True  # Include recurring event instances
        
        # Build date range: days_range in the past to days_range in the future
        start_date = datetime.now() - timedelta(days=days_range)
        end_date = datetime.now() + timedelta(days=days_range)
        
        # Build Outlook filter for date range
        filter_str = f"[Start] >= '{start_date.strftime('%m/%d/%Y')}' AND [End] <= '{end_date.strftime('%m/%d/%Y')}'"
        filtered_items = items.Restrict(filter_str)
        
        events = []
        query_lower = query.lower()
        
        # Manually filter by query since Outlook's Restrict doesn't support complex OR conditions easily
        for appointment in filtered_items:
            subject = appointment.Subject.lower() if appointment.Subject else ""
            location = appointment.Location.lower() if appointment.Location else ""
            
            # Match if query appears in either subject or location
            if query_lower in subject or query_lower in location:
                events.append(format_appointment(appointment))
        
        
        return json.dumps({
            "success": True,
            "query": query,
            "count": len(events),
            "events": events
        }, indent=2)
        
    except Exception as e:
        logger.error("Failed to search calendar events", exc_info=True, extra={
            "query": query,
            "days_range": days_range
        })
        return json.dumps({"success": False, "error": str(e)})


@mcp.tool()
def get_meeting_requests(days_range: int = 30) -> str:
    """
    Get pending meeting requests (invitations) that need a response.
    
    Retrieves meeting invitations that haven't been accepted, declined, or tentatively accepted yet.
    
    Args:
        days_range: Number of days to look ahead for meeting requests (default: 30)
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "count": int,
            "meeting_requests": [
                {
                    "subject": str,
                    "organizer": str,
                    "start": str,
                    "end": str,
                    "location": str,
                    "entry_id": str,
                    "response_status": str
                }
            ]
        }
        
    Examples:
        >>> get_meeting_requests(days_range=7)
        {"success": true, "count": 3, "meeting_requests": [...]}
        
    Notes:
        - Only returns meetings that haven't been responded to or are tentative
        - Response status: 0=None, 1=Organized, 2=Tentative, 3=Accepted, 4=Declined
    """
    try:
        outlook = get_outlook_application()
        namespace = outlook.GetNamespace("MAPI")
        calendar = namespace.GetDefaultFolder(OUTLOOK_FOLDER_CALENDAR)
        
        items = calendar.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
        
        # Filter for upcoming meetings
        start_date = datetime.now()
        end_date = datetime.now() + timedelta(days=days_range)
        
        filter_str = f"[Start] >= '{start_date.strftime('%m/%d/%Y')}' AND [End] <= '{end_date.strftime('%m/%d/%Y')}'"
        filtered_items = items.Restrict(filter_str)
        
        meeting_requests = []
        for appointment in filtered_items:
            try:
                # Check if it's a meeting (has organizer) and needs response
                # ResponseStatus: 0=None, 1=Organized, 2=Tentative, 3=Accepted, 4=Declined
                if hasattr(appointment, 'ResponseStatus') and hasattr(appointment, 'Organizer'):
                    response_status = appointment.ResponseStatus
                    
                    # Include meetings that need response (None, Tentative) but not organized by user
                    if response_status in [0, 2] and response_status != 1:
                        response_map = {
                            0: "Not Responded",
                            1: "Organized by You",
                            2: "Tentative",
                            3: "Accepted",
                            4: "Declined"
                        }
                        
                        meeting_requests.append({
                            "subject": appointment.Subject,
                            "organizer": appointment.Organizer if hasattr(appointment, 'Organizer') else None,
                            "start": str(appointment.Start),
                            "end": str(appointment.End),
                            "location": appointment.Location,
                            "body": (appointment.Body[:200] + "...") if appointment.Body and len(appointment.Body) > 200 else appointment.Body,
                            "required_attendees": appointment.RequiredAttendees,
                            "optional_attendees": appointment.OptionalAttendees,
                            "entry_id": appointment.EntryID if hasattr(appointment, 'EntryID') else None,
                            "response_status": response_map.get(response_status, "Unknown")
                        })
            except Exception:
                continue
        
        return json.dumps({
            "success": True,
            "count": len(meeting_requests),
            "meeting_requests": meeting_requests
        }, indent=2)
        
    except Exception as e:
        logger.error("Failed to get meeting requests", exc_info=True, extra={
            "days_range": days_range
        })
        return json.dumps({"success": False, "error": str(e)})


@mcp.tool()
def respond_to_meeting(entry_id: str, response: str, send_response: bool = True, comment: Optional[str] = None) -> str:
    """
    Respond to a meeting invitation (accept, decline, or tentative).
    
    Sends a response to a meeting invitation. Can optionally add a comment to the response.
    
    Args:
        entry_id: Meeting's EntryID (obtained from get_meeting_requests or get_calendar_events)
        response: Response type - "accept", "decline", or "tentative"
        send_response: Whether to send the response to the organizer (default: True)
            If False, updates calendar silently without notifying organizer
        comment: Optional comment to include in the response email
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "message": str
        }
        
    Examples:
        >>> respond_to_meeting("00000000...", "accept", send_response=True)
        {"success": true, "message": "Meeting accepted and response sent"}
        
        >>> respond_to_meeting("00000000...", "decline", send_response=True, comment="Conflict with another meeting")
        {"success": true, "message": "Meeting declined and response sent"}
        
        >>> respond_to_meeting("00000000...", "tentative", send_response=False)
        {"success": true, "message": "Meeting marked as tentative (no response sent)"}
        
    Notes:
        - "accept" adds the meeting to your calendar as accepted
        - "decline" removes the meeting from your calendar
        - "tentative" marks the meeting as tentative in your calendar
        - If send_response=True, organizer receives your response via email
    """
    try:
        outlook = get_outlook_application()
        namespace = outlook.GetNamespace("MAPI")
        
        # Get the meeting by EntryID
        appointment = namespace.GetItemFromID(entry_id)
        
        # Validate response type
        response_lower = response.lower()
        if response_lower not in ["accept", "decline", "tentative"]:
            return json.dumps({
                "success": False,
                "error": f"Invalid response type '{response}'. Must be 'accept', 'decline', or 'tentative'"
            })
        
        # Respond to the meeting
        # olMeetingAccepted = 3, olMeetingDeclined = 4, olMeetingTentative = 2
        try:
            if response_lower == "accept":
                meeting_response = appointment.Respond(3, send_response)
                action = "accepted"
            elif response_lower == "decline":
                meeting_response = appointment.Respond(4, send_response)
                action = "declined"
            else:  # tentative
                meeting_response = appointment.Respond(2, send_response)
                action = "marked as tentative"
            
            # Add comment if provided and response is being sent
            if comment and send_response and meeting_response:
                try:
                    meeting_response.Body = comment + "\n\n" + (meeting_response.Body if meeting_response.Body else "")
                    meeting_response.Send()
                except Exception:
                    pass
            
            response_sent_msg = " and response sent" if send_response else " (no response sent)"
            
            return json.dumps({
                "success": True,
                "message": f"Meeting {action}{response_sent_msg}"
            }, indent=2)
            
        except Exception as e:
            return json.dumps({
                "success": False,
                "error": f"Failed to respond to meeting: {e}"
            })
        
    except Exception as e:
        logger.error("Failed to respond to meeting", exc_info=True, extra={
            "entry_id": entry_id,
            "response": response
        })
        return json.dumps({"success": False, "error": str(e)})


# ============================================================================
# CONTACT TOOLS
# ============================================================================

@mcp.tool()
def get_contacts(limit: int = DEFAULT_CONTACT_LIMIT, search_name: Optional[str] = None) -> str:
    """
    Get contacts from the Outlook Contacts folder.
    
    Retrieves contacts from the user's default Contacts folder, sorted alphabetically
    by name. Can optionally filter by name.
    
    Args:
        limit: Maximum number of contacts to return (default: 50, max: 200)
        search_name: Filter by name substring (optional, case-insensitive)
            Example: "Smith" will match "John Smith", "Jane Smith", etc.
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "count": int,
            "contacts": [list of contact dictionaries]
        }
        
    Examples:
        >>> get_contacts(limit=10)
        {"success": true, "count": 10, "contacts": [...]}
        
        >>> get_contacts(limit=20, search_name="Smith")
        {"success": true, "count": 5, "contacts": [...]}
        
    Notes:
        - Limited to MAX_CONTACT_LIMIT (200) for performance
        - Contacts are sorted alphabetically by full name
        - When search_name is provided, scans more contacts to ensure enough matches
    """
    try:
        outlook = get_outlook_application()
        namespace = outlook.GetNamespace("MAPI")
        contacts_folder = namespace.GetDefaultFolder(OUTLOOK_FOLDER_CONTACTS)
        
        # Apply limit cap to prevent performance issues
        limit = min(limit, MAX_CONTACT_LIMIT)
        
        items = contacts_folder.Items
        items.Sort("[FullName]")  # Sort alphabetically by full name
        
        contacts = []
        
        # PERFORMANCE OPTIMIZATION: Use GetFirst()/GetNext() instead of items.Count
        # When filtering by name, scan more items to find enough matches
        max_scan = limit * 3 if search_name else limit
        contact = items.GetFirst()
        scanned = 0
        
        while contact is not None and len(contacts) < limit and scanned < max_scan:
            try:
                scanned += 1
                
                # Apply name filter if provided
                if search_name:
                    full_name = contact.FullName.lower() if contact.FullName else ""
                    if search_name.lower() not in full_name:
                        contact = items.GetNext()
                        continue
                
                contacts.append(format_contact(contact))
            except Exception as e:
                pass
            
            contact = items.GetNext()
        
        
        return json.dumps({
            "success": True,
            "count": len(contacts),
            "contacts": contacts
        }, indent=2)
        
    except Exception as e:
        logger.error("Failed to get contacts", exc_info=True, extra={
            "limit": limit,
            "search_name": search_name
        })
        return json.dumps({"success": False, "error": str(e)})


@mcp.tool()
def create_contact(
    full_name: str,
    email: str,
    company: Optional[str] = None,
    job_title: Optional[str] = None,
    business_phone: Optional[str] = None,
    mobile_phone: Optional[str] = None,
    home_phone: Optional[str] = None
) -> str:
    """
    Create a new contact in the Outlook Contacts folder.
    
    Creates a contact entry with the provided information. Only name and email
    are required; all other fields are optional.
    
    Args:
        full_name: Contact's full name (required)
            Example: "John Smith"
        email: Primary email address (required)
            Example: "john.smith@company.com"
        company: Company name (optional)
            Example: "Acme Corp"
        job_title: Job title (optional)
            Example: "Senior Developer"
        business_phone: Business phone number (optional)
            Example: "+33 1 23 45 67 89"
        mobile_phone: Mobile phone number (optional)
        home_phone: Home phone number (optional)
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "message": str
        }
        
    Examples:
        >>> create_contact("Jane Doe", "jane.doe@company.com")
        {"success": true, "message": "Contact 'Jane Doe' created"}
        
        >>> create_contact("Bob Smith", "bob@company.com", company="Acme Corp", 
        ...                job_title="Manager")
        {"success": true, "message": "Contact 'Bob Smith' created"}
        
    Notes:
        - Contact is saved to the default Contacts folder
        - Duplicate checking is not performed (Outlook allows duplicate contacts)
        - Best practice: Maintain accurate contact information
    """
    try:
        outlook = get_outlook_application()
        contact = outlook.CreateItem(OUTLOOK_ITEM_CONTACT)
        
        # Set required fields
        contact.FullName = full_name
        contact.Email1Address = email
        
        # Set optional fields if provided
        if company:
            contact.CompanyName = company
        if job_title:
            contact.JobTitle = job_title
        if business_phone:
            contact.BusinessTelephoneNumber = business_phone
        if mobile_phone:
            contact.MobileTelephoneNumber = mobile_phone
        if home_phone:
            contact.HomeTelephoneNumber = home_phone
        
        # Save the contact
        contact.Save()
        
        
        return json.dumps({
            "success": True,
            "message": f"Contact '{full_name}' created"
        }, indent=2)
        
    except Exception as e:
        logger.error("Failed to create contact", exc_info=True, extra={
            "full_name": full_name,
            "email": email
        })
        return json.dumps({"success": False, "error": str(e)})


@mcp.tool()
def search_contacts(query: str) -> str:
    """
    Search for contacts by keyword in name, email, or company.
    
    Performs a comprehensive search across all contacts, looking for matches
    in name, email address, and company fields.
    
    Args:
        query: Search keyword (case-insensitive)
            Searches in full name, email address, and company name fields
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "query": str,
            "count": int,
            "contacts": [list of matching contact dictionaries]
        }
        
    Examples:
        >>> search_contacts("Smith")
        {"success": true, "query": "Smith", "count": 3, "contacts": [...]}
        
        >>> search_contacts("company.com")
        {"success": true, "query": "company.com", "count": 156, "contacts": [...]}
        
        >>> search_contacts("Acme Corp")
        {"success": true, "query": "Acme Corp", "count": 42, "contacts": [...]}
        
    Notes:
        - Search is case-insensitive
        - Searches all contacts (no limit)
        - Uses safe attribute access to handle contacts with missing data
        - Only returns successfully formatted contacts (skips corrupted entries)
    """
    try:
        outlook = get_outlook_application()
        namespace = outlook.GetNamespace("MAPI")
        contacts_folder = namespace.GetDefaultFolder(OUTLOOK_FOLDER_CONTACTS)
        
        items = contacts_folder.Items
        
        contacts = []
        query_lower = query.lower()
        
        # Iterate through all contacts and check each field
        for contact in items:
            # Safely extract searchable fields
            # Some contacts may have missing or corrupted data
            try:
                full_name = contact.FullName.lower() if contact.FullName else ""
            except Exception:
                full_name = ""
            
            try:
                email = contact.Email1Address.lower() if contact.Email1Address else ""
            except Exception:
                email = ""
            
            try:
                company = contact.CompanyName.lower() if contact.CompanyName else ""
            except Exception:
                company = ""
            
            # Check if query matches any of the searchable fields
            if query_lower in full_name or query_lower in email or query_lower in company:
                formatted = format_contact(contact)
                # Only add successfully formatted contacts (skip corrupted ones)
                if "error" not in formatted:
                    contacts.append(formatted)
        
        
        return json.dumps({
            "success": True,
            "query": query,
            "count": len(contacts),
            "contacts": contacts
        }, indent=2)
        
    except Exception as e:
        logger.error("Failed to search contacts", exc_info=True, extra={
            "query": query
        })
        return json.dumps({"success": False, "error": str(e)})


# ============================================================================
# OUT-OF-OFFICE TOOLS
# ============================================================================

@mcp.tool()
def get_out_of_office_settings() -> str:
    """
    Get current Out-of-Office (automatic reply) settings.
    
    Retrieves the current automatic reply configuration, including whether it's enabled,
    the scheduled time period, and the reply messages.
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "enabled": bool,
            "scheduled": bool,
            "start_time": str (if scheduled),
            "end_time": str (if scheduled),
            "internal_reply": str,
            "external_reply": str,
            "external_audience": str
        }
        
    Examples:
        >>> get_out_of_office_settings()
        {
            "success": true,
            "enabled": true,
            "scheduled": true,
            "start_time": "2025-12-20 00:00:00",
            "end_time": "2025-12-27 00:00:00",
            "internal_reply": "I'm out of office...",
            "external_reply": "I'm currently unavailable...",
            "external_audience": "Known"
        }
        
    Notes:
        - external_audience can be: "None", "Known", or "All"
        - "Known" means only contacts/addresses in your organization
        - "All" means everyone who sends you an email
    """
    try:
        outlook = get_outlook_application()
        namespace = outlook.GetNamespace("MAPI")
        
        # Get current user's account
        # For Outlook 2010+, use Accounts collection
        try:
            account = namespace.Accounts.Item(1)  # Primary account
            
            # Try to access AutoReplyState (not available in all Outlook versions)
            try:
                # olAutoReplyStateDisabled = 0, olAutoReplyStateEnabled = 1, olAutoReplyStateScheduled = 2
                auto_reply_state = account.AutoReplyState
                enabled = auto_reply_state > 0
                scheduled = auto_reply_state == 2
                
                result = {
                    "success": True,
                    "enabled": enabled,
                    "scheduled": scheduled
                }
                
                # Get scheduled times if applicable
                if scheduled:
                    try:
                        result["start_time"] = str(account.AutoReplyStartTime)
                        result["end_time"] = str(account.AutoReplyEndTime)
                    except Exception:
                        pass
                
                # Get reply messages
                try:
                    result["internal_reply"] = account.AutoReplyTextInternal
                    result["external_reply"] = account.AutoReplyTextExternal
                except Exception:
                    pass
                
                # Get external audience setting
                # olExternalAudienceNone = 0, olExternalAudienceKnown = 1, olExternalAudienceAll = 2
                try:
                    external_audience_map = {0: "None", 1: "Known", 2: "All"}
                    result["external_audience"] = external_audience_map.get(
                        account.ExternalAudience, "Unknown"
                    )
                except Exception:
                    result["external_audience"] = "Unknown"
                
                return json.dumps(result, indent=2)
                
            except AttributeError:
                # AutoReplyState not available, try alternative method using CDO
                return json.dumps({
                    "success": False,
                    "error": "Out-of-Office settings not accessible via COM automation on this Outlook version. "
                            "Try using Outlook's UI or Exchange Web Services."
                })
                
        except Exception as e:
            return json.dumps({
                "success": False,
                "error": f"Failed to access account settings: {e}"
            })
        
    except Exception as e:
        logger.error("Failed to get out-of-office settings", exc_info=True)
        return json.dumps({"success": False, "error": str(e)})


@mcp.tool()
def set_out_of_office(
    enabled: bool,
    internal_reply: str,
    external_reply: Optional[str] = None,
    external_audience: str = "Known",
    scheduled: bool = False,
    start_time: Optional[str] = None,
    end_time: Optional[str] = None
) -> str:
    """
    Set Out-of-Office (automatic reply) settings.
    
    Configures automatic replies for when you're away from the office.
    Can be set to immediate or scheduled for a specific time period.
    
    Args:
        enabled: Whether to enable automatic replies (True=on, False=off)
        internal_reply: Message to send to internal recipients (required)
        external_reply: Message to send to external recipients (optional, defaults to internal_reply)
        external_audience: Who receives external replies - "None", "Known", or "All" (default: "Known")
            - "None": No external replies
            - "Known": Only to people in your organization/contacts
            - "All": Everyone who emails you
        scheduled: Whether to schedule automatic replies (True) or enable immediately (False)
        start_time: Start date/time for scheduled replies (ISO format: "2025-12-20 00:00")
            Required if scheduled=True
        end_time: End date/time for scheduled replies (ISO format: "2025-12-27 00:00")
            Required if scheduled=True
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "message": str
        }
        
    Examples:
        >>> set_out_of_office(
        ...     enabled=True,
        ...     internal_reply="I'm out of office until next week.",
        ...     external_reply="I'm currently unavailable. I'll respond when I return.",
        ...     external_audience="Known"
        ... )
        {"success": true, "message": "Out-of-Office enabled"}
        
        >>> set_out_of_office(
        ...     enabled=True,
        ...     internal_reply="On vacation",
        ...     scheduled=True,
        ...     start_time="2025-12-20 00:00",
        ...     end_time="2025-12-27 00:00"
        ... )
        {"success": true, "message": "Out-of-Office scheduled"}
        
    Notes:
        - If external_reply is not provided, internal_reply is used for both
        - Scheduled replies automatically turn off after end_time
        - This may not work on all Outlook versions (requires Outlook 2010+)
    """
    try:
        outlook = get_outlook_application()
        namespace = outlook.GetNamespace("MAPI")
        
        # Validate scheduled parameters
        if scheduled and (not start_time or not end_time):
            return json.dumps({
                "success": False,
                "error": "start_time and end_time are required when scheduled=True"
            })
        
        # Use internal_reply for external if not provided
        if external_reply is None:
            external_reply = internal_reply
        
        # Validate external_audience
        external_audience_map = {
            "none": 0,
            "known": 1,
            "all": 2
        }
        external_audience_value = external_audience_map.get(external_audience.lower())
        if external_audience_value is None:
            return json.dumps({
                "success": False,
                "error": f"Invalid external_audience '{external_audience}'. Must be 'None', 'Known', or 'All'"
            })
        
        try:
            account = namespace.Accounts.Item(1)  # Primary account
            
            try:
                # Set automatic reply state
                # olAutoReplyStateDisabled = 0, olAutoReplyStateEnabled = 1, olAutoReplyStateScheduled = 2
                if not enabled:
                    account.AutoReplyState = 0  # Disabled
                    message = "Out-of-Office disabled"
                elif scheduled:
                    account.AutoReplyState = 2  # Scheduled
                    
                    # Parse and set scheduled times
                    try:
                        start_dt = date_parser.parse(start_time)
                        end_dt = date_parser.parse(end_time)
                        account.AutoReplyStartTime = start_dt
                        account.AutoReplyEndTime = end_dt
                    except Exception as e:
                        return json.dumps({
                            "success": False,
                            "error": f"Invalid date format: {e}. Use ISO format like '2025-12-20 00:00'"
                        })
                    
                    message = f"Out-of-Office scheduled from {start_time} to {end_time}"
                else:
                    account.AutoReplyState = 1  # Enabled immediately
                    message = "Out-of-Office enabled"
                
                # Set reply messages
                account.AutoReplyTextInternal = internal_reply
                account.AutoReplyTextExternal = external_reply
                
                # Set external audience
                account.ExternalAudience = external_audience_value
                
                return json.dumps({
                    "success": True,
                    "message": message
                }, indent=2)
                
            except AttributeError:
                return json.dumps({
                    "success": False,
                    "error": "Out-of-Office settings not accessible via COM automation on this Outlook version. "
                            "Try using Outlook's UI or Exchange Web Services."
                })
                
        except Exception as e:
            return json.dumps({
                "success": False,
                "error": f"Failed to access account settings: {e}"
            })
        
    except Exception as e:
        logger.error("Failed to set out-of-office", exc_info=True, extra={
            "enabled": enabled,
            "scheduled": scheduled
        })
        return json.dumps({"success": False, "error": str(e)})


@mcp.tool()
def disable_out_of_office() -> str:
    """
    Disable Out-of-Office (automatic reply) settings.
    
    Turns off automatic replies. This is a convenience function equivalent to
    set_out_of_office(enabled=False, ...).
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "message": str
        }
        
    Examples:
        >>> disable_out_of_office()
        {"success": true, "message": "Out-of-Office disabled"}
        
    Notes:
        - This only disables automatic replies, doesn't delete the messages
        - Previous messages are preserved and can be re-enabled later
    """
    try:
        outlook = get_outlook_application()
        namespace = outlook.GetNamespace("MAPI")
        
        try:
            account = namespace.Accounts.Item(1)  # Primary account
            
            try:
                # olAutoReplyStateDisabled = 0
                account.AutoReplyState = 0
                
                return json.dumps({
                    "success": True,
                    "message": "Out-of-Office disabled"
                }, indent=2)
                
            except AttributeError:
                return json.dumps({
                    "success": False,
                    "error": "Out-of-Office settings not accessible via COM automation on this Outlook version. "
                            "Try using Outlook's UI or Exchange Web Services."
                })
                
        except Exception as e:
            return json.dumps({
                "success": False,
                "error": f"Failed to access account settings: {e}"
            })
        
    except Exception as e:
        logger.error("Failed to disable out-of-office", exc_info=True)
        return json.dumps({"success": False, "error": str(e)})


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    """
    Main entry point for the MCP Outlook server.
    
    Starts the FastMCP server which listens for requests from MCP clients
    (such as Claude Desktop, Cursor, or other AI assistants).
    
    The server exposes all @mcp.tool() decorated functions as callable tools
    that clients can invoke to interact with Outlook.
    
    Usage:
        python outlook_mcp.py
        
    Notes:
        - Server runs indefinitely until interrupted (Ctrl+C)
        - Requires Microsoft Outlook to be installed and configured
        - All operations use the currently logged-in Outlook profile
        - Server logs operations for audit purposes
    """
    try:
        # Run the MCP server (blocks until interrupted)
        mcp.run()
    except KeyboardInterrupt:
        pass  # Silent shutdown
    except Exception as e:
        logger.critical("Server crashed unexpectedly", exc_info=True)
        raise

