# Examples - MCP Outlook

Real-world examples of using MCP Outlook with AI assistants.

## Email Examples

### Example 1: Check Unread Emails

**User**: "Show me my unread emails"

**AI Assistant**:
```
I'll check your unread emails in Outlook.
```

**Tool Call**:
```json
{
  "tool": "get_inbox_emails",
  "arguments": {
    "limit": 10,
    "unread_only": true
  }
}
```

**Response**:
```json
{
  "success": true,
  "count": 3,
  "emails": [
    {
      "subject": "Q1 Budget Review",
      "sender": "Jane Smith",
      "sender_email": "jane.smith@company.com",
      "received_time": "2025-01-15 09:30:00",
      "body": "Hi team, please review...",
      "unread": true
    }
  ]
}
```

**AI Response**:
```
You have 3 unread emails:
1. From Jane Smith: "Q1 Budget Review" (received at 9:30 AM)
2. ...
```

---

### Example 2: Search for Specific Emails

**User**: "Find all emails about 'project alpha' from last week"

**AI Assistant**:
```
I'll search your emails for "project alpha".
```

**Tool Call**:
```json
{
  "tool": "search_emails",
  "arguments": {
    "query": "project alpha",
    "folder": "all",
    "limit": 20
  }
}
```

**AI Response**:
```
Found 5 emails about "project alpha":
1. From John Doe (Jan 12): "Project Alpha - Status Update"
2. From Marketing Team (Jan 10): "Project Alpha Launch Plan"
...
```

---

### Example 3: Send a Follow-up Email

**User**: "Send a follow-up email to john@example.com thanking him for the meeting"

**AI Assistant**:
```
I'll send a follow-up email to John.
```

**Tool Call**:
```json
{
  "tool": "send_email",
  "arguments": {
    "to": "john@example.com",
    "subject": "Thank You - Meeting Follow-up",
    "body": "Hi John,\n\nThank you for taking the time to meet with me today. I appreciate your insights on the project.\n\nBest regards",
    "importance": "normal"
  }
}
```

**AI Response**:
```
✓ Email sent successfully to john@example.com
```

---

### Example 4: Draft Email with Multiple Recipients

**User**: "Create a draft email to the team about tomorrow's standup, CC my manager"

**AI Assistant**:
```
I'll create a draft email about tomorrow's standup.
```

**Tool Call**:
```json
{
  "tool": "create_draft_email",
  "arguments": {
    "to": "team@company.com",
    "subject": "Reminder: Tomorrow's Standup",
    "body": "Hi Team,\n\nJust a friendly reminder about tomorrow's standup at 9:00 AM.\n\nSee you there!",
    "cc": "manager@company.com"
  }
}
```

**AI Response**:
```
✓ Draft email created in your Outlook Drafts folder. You can review and send it when ready.
```

---

## 📅 Calendar Examples

### Example 5: Check Today's Meetings

**User**: "What meetings do I have today?"

**AI Assistant**:
```
Let me check your calendar for today.
```

**Tool Call**:
```json
{
  "tool": "get_calendar_events",
  "arguments": {
    "days_ahead": 0,
    "include_past": true
  }
}
```

**AI Response**:
```
You have 3 meetings today:

1. Team Standup (9:00 AM - 9:30 AM)
   Location: Conference Room A
   
2. Project Review (2:00 PM - 3:00 PM)
   Location: Zoom
   Attendees: John, Sarah, Mike
   
3. 1-on-1 with Manager (4:00 PM - 4:30 PM)
   Location: Manager's office
```

---

### Example 6: Schedule a New Meeting

**User**: "Schedule a team meeting for tomorrow at 2pm about sprint planning"

**AI Assistant**:
```
I'll schedule a team meeting for tomorrow at 2 PM.
```

**Tool Call**:
```json
{
  "tool": "create_calendar_event",
  "arguments": {
    "subject": "Sprint Planning",
    "start_time": "2025-01-16 14:00",
    "end_time": "2025-01-16 15:00",
    "location": "Conference Room B",
    "body": "Sprint planning session for the next sprint.\n\nAgenda:\n- Review previous sprint\n- Plan upcoming stories\n- Assign tasks",
    "required_attendees": "team@company.com",
    "reminder_minutes": 15
  }
}
```

**AI Response**:
```
✓ Meeting scheduled successfully!

Sprint Planning
Tomorrow, January 16 at 2:00 PM - 3:00 PM
Location: Conference Room B
Attendees: team@company.com
Reminder: 15 minutes before
```

---

### Example 7: Find Meetings About a Topic

**User**: "When was my last meeting about budget?"

**AI Assistant**:
```
I'll search your calendar for meetings about budget.
```

**Tool Call**:
```json
{
  "tool": "search_calendar_events",
  "arguments": {
    "query": "budget",
    "days_range": 90
  }
}
```

**AI Response**:
```
Found 2 meetings about "budget" in the last 90 days:

1. Q1 Budget Review (January 10, 2025)
   9:00 AM - 10:30 AM with Finance Team
   
2. Budget Planning Session (December 15, 2024)
   2:00 PM - 4:00 PM with Management Team

Your most recent budget meeting was on January 10.
```

---

## 👥 Contact Examples

### Example 8: Find a Contact

**User**: "What's Jane Smith's email address?"

**AI Assistant**:
```
Let me look up Jane Smith's contact information.
```

**Tool Call**:
```json
{
  "tool": "search_contacts",
  "arguments": {
    "query": "Jane Smith"
  }
}
```

**AI Response**:
```
Found Jane Smith:

Email: jane.smith@company.com
Company: Acme Corp
Job Title: Product Manager
Business Phone: +1-555-1234
Mobile: +1-555-5678
```

---

### Example 9: Add a New Contact

**User**: "Add John Doe from TechCorp as a contact, his email is john@techcorp.com"

**AI Assistant**:
```
I'll add John Doe to your contacts.
```

**Tool Call**:
```json
{
  "tool": "create_contact",
  "arguments": {
    "full_name": "John Doe",
    "email": "john@techcorp.com",
    "company": "TechCorp"
  }
}
```

**AI Response**:
```
✓ Contact added successfully!

Name: John Doe
Email: john@techcorp.com
Company: TechCorp
```

---

### Example 10: Find All Contacts from a Company

**User**: "Show me all contacts from Acme Corp"

**AI Assistant**:
```
I'll search for all Acme Corp contacts.
```

**Tool Call**:
```json
{
  "tool": "search_contacts",
  "arguments": {
    "query": "Acme Corp"
  }
}
```

**AI Response**:
```
Found 3 contacts from Acme Corp:

1. Jane Smith - Product Manager
   jane.smith@company.com
   
2. Bob Johnson - Engineer
   bob.johnson@company.com
   
3. Sarah Williams - Designer
   sarah.williams@company.com
```

---

## 🔄 Complex Workflows

### Workflow 1: Morning Briefing

**User**: "Give me my morning briefing"

**AI Assistant** performs multiple tool calls:

1. **Check Unread Emails**
```json
{"tool": "get_inbox_emails", "arguments": {"limit": 5, "unread_only": true}}
```

2. **Check Today's Calendar**
```json
{"tool": "get_calendar_events", "arguments": {"days_ahead": 0, "include_past": true}}
```

**AI Response**:
```
EMAILS (3 unread):
- Jane Smith: "Q1 Budget Review"
- Team: "Sprint Planning Notes"
- HR: "Benefits Update"

📅 TODAY'S SCHEDULE:
- 9:00 AM: Team Standup (30 min)
- 2:00 PM: Project Review (1 hour)
- 4:00 PM: 1-on-1 with Manager (30 min)

You have a fairly busy day ahead!
```

---

### Workflow 2: Meeting Preparation

**User**: "I have a meeting with John in 10 minutes. Show me our recent email thread and his contact info."

**AI Assistant**:

1. **Search Emails from John**
```json
{"tool": "search_emails", "arguments": {"query": "from:john", "limit": 5}}
```

2. **Get John's Contact**
```json
{"tool": "search_contacts", "arguments": {"query": "John"}}
```

**AI Response**:
```
RECENT EMAILS WITH JOHN:
1. "Project Update" (Jan 15, 9:00 AM)
2. "Meeting Agenda" (Jan 14, 3:00 PM)
3. "Budget Questions" (Jan 12, 10:00 AM)

👤 JOHN'S CONTACT INFO:
Name: John Doe
Email: john@company.com
Phone: +1-555-1234
Job Title: Senior Engineer
Company: Acme Corp

Ready for your meeting!
```

---

### Workflow 3: End of Day Summary

**User**: "What did I accomplish today?"

**AI Assistant**:

1. **Check Sent Emails**
```json
{"tool": "get_sent_emails", "arguments": {"limit": 10}}
```

2. **Check Today's Meetings**
```json
{"tool": "get_calendar_events", "arguments": {"days_ahead": 0, "include_past": true}}
```

**AI Response**:
```
📊 YOUR DAY SUMMARY:

EMAILS SENT: 7
- Follow-up to Jane about budget
- Team update on sprint progress
- Response to HR about benefits
...

📅 MEETINGS ATTENDED: 3
- Team Standup (9:00 AM)
- Project Review (2:00 PM)
- 1-on-1 with Manager (4:00 PM)

Great work today! ```

---

## 💡 Pro Tips

### Tip 1: Use Natural Language

Instead of: "Call get_inbox_emails with limit 5"

Say: "Show me my last 5 emails"

---

### Tip 2: Be Specific with Time

Instead of: "Show my meetings"

Say: "Show my meetings for next week" or "What's on my calendar tomorrow?"

---

### Tip 3: Combine Requests

Instead of:
- "Show my emails"
- "Show my calendar"

Say: "Give me my morning briefing with emails and calendar"

---

### Tip 4: Use Context

Instead of: "Send email to john@example.com..."

Say: "Reply to John's last email saying thanks"
(AI can search for John's email and compose a reply)

---

### Tip 5: Leverage Search

Instead of: "Show all my emails"

Say: "Find emails about project alpha from last month"

---

## Use Cases

### Personal Productivity
- Morning briefings
- End-of-day summaries
- Meeting preparation
- Email triage

### Team Coordination
- Schedule team meetings
- Send status updates
- Share contact information
- Coordinate calendars

### Project Management
- Track project-related emails
- Schedule project meetings
- Find project documentation links
- Follow up with stakeholders

### Customer Relations
- Quick access to customer emails
- Schedule customer meetings
- Maintain customer contacts
- Track communication history

---

## Getting Started

Try these commands to get familiar:

1. `"Show me my last 5 emails"`
2. `"What's on my calendar today?"`
3. `"Find John Smith's email address"`
4. `"Create a draft email to my team about tomorrow's meeting"`
5. `"Schedule a 30-minute 1-on-1 with Sarah for next Tuesday at 2pm"`

---

**Happy Automating! **

For more information, see:
- [README.md](README.md) - Full documentation
- [QUICK_START.md](QUICK_START.md) - Setup guide
- [ARCHITECTURE.md](ARCHITECTURE.md) - Technical details

