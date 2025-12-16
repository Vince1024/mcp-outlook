"""
Tests for MCP Outlook Server

Note: These tests require Microsoft Outlook to be installed and configured.
"""

import json
import pytest
from unittest.mock import Mock, patch, MagicMock

# Import the module to test
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from outlook_mcp import (
    format_email,
    format_appointment,
    format_contact,
    get_outlook_application,
)


@pytest.fixture
def mock_outlook():
    """Mock Outlook application for testing."""
    with patch('outlook_mcp.win32com.client.Dispatch') as mock_dispatch:
        outlook_mock = MagicMock()
        mock_dispatch.return_value = outlook_mock
        yield outlook_mock


@pytest.fixture
def mock_mail_item():
    """Mock Outlook mail item."""
    mail = MagicMock()
    mail.Subject = "Test Email"
    mail.SenderName = "John Doe"
    mail.SenderEmailAddress = "john@example.com"
    mail.To = "recipient@example.com"
    mail.CC = ""
    mail.BCC = ""
    mail.ReceivedTime = "2025-01-15 10:00:00"
    mail.SentOn = "2025-01-15 09:55:00"
    mail.Body = "This is a test email body."
    mail.Attachments = MagicMock(Count=0)
    mail.Importance = 1
    mail.UnRead = False
    mail.Categories = ""
    return mail


@pytest.fixture
def mock_appointment():
    """Mock Outlook appointment."""
    appt = MagicMock()
    appt.Subject = "Test Meeting"
    appt.Start = "2025-01-15 14:00:00"
    appt.End = "2025-01-15 15:00:00"
    appt.Location = "Conference Room A"
    appt.Organizer = "organizer@example.com"
    appt.RequiredAttendees = "attendee1@example.com"
    appt.OptionalAttendees = ""
    appt.Body = "Meeting agenda..."
    appt.AllDayEvent = False
    appt.ReminderSet = True
    appt.ReminderMinutesBeforeStart = 15
    appt.Categories = ""
    appt.BusyStatus = 2
    return appt


@pytest.fixture
def mock_contact():
    """Mock Outlook contact."""
    contact = MagicMock()
    contact.FullName = "Jane Smith"
    contact.Email1Address = "jane.smith@example.com"
    contact.Email2Address = ""
    contact.Email3Address = ""
    contact.CompanyName = "Acme Corp"
    contact.JobTitle = "Manager"
    contact.BusinessTelephoneNumber = "+1-555-1234"
    contact.MobileTelephoneNumber = "+1-555-5678"
    contact.HomeTelephoneNumber = ""
    contact.BusinessAddress = "123 Main St"
    contact.Categories = ""
    return contact


class TestFormatters:
    """Test formatting functions."""
    
    def test_format_email(self, mock_mail_item):
        """Test email formatting."""
        result = format_email(mock_mail_item)
        
        assert result["subject"] == "Test Email"
        assert result["sender"] == "John Doe"
        assert result["sender_email"] == "john@example.com"
        assert result["unread"] is False
        assert result["has_attachments"] is False
    
    def test_format_appointment(self, mock_appointment):
        """Test appointment formatting."""
        result = format_appointment(mock_appointment)
        
        assert result["subject"] == "Test Meeting"
        assert result["location"] == "Conference Room A"
        assert result["is_all_day_event"] is False
        assert result["reminder_set"] is True
        assert result["reminder_minutes"] == 15
    
    def test_format_contact(self, mock_contact):
        """Test contact formatting."""
        result = format_contact(mock_contact)
        
        assert result["full_name"] == "Jane Smith"
        assert result["email1"] == "jane.smith@example.com"
        assert result["company"] == "Acme Corp"
        assert result["job_title"] == "Manager"


class TestOutlookConnection:
    """Test Outlook connection."""
    
    def test_get_outlook_application_success(self, mock_outlook):
        """Test successful Outlook connection."""
        outlook = get_outlook_application()
        assert outlook is not None
    
    def test_get_outlook_application_failure(self):
        """Test Outlook connection failure."""
        with patch('outlook_mcp.win32com.client.Dispatch', side_effect=Exception("Outlook not found")):
            with pytest.raises(ValueError, match="Unable to connect to Outlook"):
                get_outlook_application()


class TestEmailTools:
    """Test email-related tools."""
    
    @pytest.mark.integration
    def test_get_inbox_emails_integration(self):
        """Integration test for getting inbox emails (requires Outlook)."""
        # This test would require actual Outlook instance
        # Skip if Outlook is not available
        pytest.skip("Requires actual Outlook installation")
    
    @pytest.mark.integration
    def test_send_email_integration(self):
        """Integration test for sending email (requires Outlook)."""
        pytest.skip("Requires actual Outlook installation")


class TestCalendarTools:
    """Test calendar-related tools."""
    
    @pytest.mark.integration
    def test_get_calendar_events_integration(self):
        """Integration test for getting calendar events (requires Outlook)."""
        pytest.skip("Requires actual Outlook installation")


class TestContactTools:
    """Test contact-related tools."""
    
    @pytest.mark.integration
    def test_get_contacts_integration(self):
        """Integration test for getting contacts (requires Outlook)."""
        pytest.skip("Requires actual Outlook installation")


# Run tests
if __name__ == "__main__":
    pytest.main([__file__, "-v"])

