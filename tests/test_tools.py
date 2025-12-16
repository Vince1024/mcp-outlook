"""
Test MCP Outlook tools manually
"""

import sys
import json

# Add src to path
sys.path.insert(0, 'src')

from outlook_mcp import (
    get_calendar_events,
    get_contacts,
    get_inbox_emails,
)

print("=" * 60)
print("TESTING MCP OUTLOOK TOOLS")
print("=" * 60)

# Test 1: Get Calendar Events
print("\n[TEST 1] Getting calendar events (next 7 days)...")
print("-" * 60)
try:
    result = get_calendar_events(days_ahead=7, include_past=False)
    data = json.loads(result)
    if data.get("success"):
        print(f"[OK] Found {data['count']} events")
        if data['count'] > 0:
            # Show first event
            first_event = data['events'][0]
            print(f"\nFirst event:")
            print(f"  Subject: {first_event['subject']}")
            print(f"  Start: {first_event['start']}")
            print(f"  Location: {first_event.get('location', 'N/A')}")
    else:
        print(f"[FAIL] {data.get('error', 'Unknown error')}")
except Exception as e:
    print(f"[ERROR] {e}")

# Test 2: Get Contacts
print("\n[TEST 2] Getting contacts (limit 5)...")
print("-" * 60)
try:
    result = get_contacts(limit=5)
    data = json.loads(result)
    if data.get("success"):
        print(f"[OK] Found {data['count']} contacts")
        if data['count'] > 0:
            # Show first contact
            first_contact = data['contacts'][0]
            print(f"\nFirst contact:")
            print(f"  Name: {first_contact['full_name']}")
            print(f"  Email: {first_contact.get('email1', 'N/A')}")
            print(f"  Company: {first_contact.get('company', 'N/A')}")
    else:
        print(f"[FAIL] {data.get('error', 'Unknown error')}")
except Exception as e:
    print(f"[ERROR] {e}")

# Test 3: Get Inbox Emails
print("\n[TEST 3] Getting inbox emails (limit 5)...")
print("-" * 60)
try:
    result = get_inbox_emails(limit=5, unread_only=False)
    data = json.loads(result)
    if data.get("success"):
        print(f"[OK] Found {data['count']} emails")
        if data['count'] > 0:
            # Show first email
            first_email = data['emails'][0]
            print(f"\nFirst email:")
            print(f"  Subject: {first_email['subject']}")
            print(f"  From: {first_email['sender']}")
            print(f"  Received: {first_email.get('received_time', 'N/A')}")
        else:
            print("  (Inbox is empty)")
    else:
        print(f"[FAIL] {data.get('error', 'Unknown error')}")
except Exception as e:
    print(f"[ERROR] {e}")

print("\n" + "=" * 60)
print("TESTS COMPLETED")
print("=" * 60)

