# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2025-12-16

### 🎉 Initial Public Release

#### Added
- **Email Management**
  - `get_inbox_emails` - Retrieve emails from inbox with unread filtering
  - `get_sent_emails` - Retrieve sent emails
  - `search_emails` - Search emails across standard folders
  - `send_email` - Send emails with CC/BCC and importance levels
  - `create_draft_email` - Create draft emails without sending

- **Folder Management**
  - `list_outlook_folders` - List all Outlook folders (optimized, no counts)
  - `search_emails_in_custom_folder` - Search in specific custom folders with date filtering
  - `list_outlook_rules` - List all Outlook mail rules

- **Calendar Management**
  - `get_calendar_events` - Get upcoming calendar events
  - `create_calendar_event` - Create new calendar events with attendees
  - `search_calendar_events` - Search events by subject or location

- **Contact Management**
  - `get_contacts` - Retrieve contacts with optional name filtering
  - `create_contact` - Create new contacts
  - `search_contacts` - Search contacts by name, email, or company

#### Performance Optimizations
- 🚀 Folder caching (45x faster on repeated searches)
- 🚀 Date filtering (search only recent emails, default: 2 days)
- 🚀 Direct indexing (faster iteration without items.Count)
- 🚀 Reduced limits (prevents long freezes, max 50 emails)
- 🚀 Smart defaults optimized for daily usage

#### Documentation
- Comprehensive README with installation and usage
- QUICK_START guide for rapid setup
- EXAMPLES with real-world use cases
- OPTIMIZATIONS documentation with performance details
- PUBLISHING_GUIDE for GitHub publication

#### Technical Details
- Built with FastMCP framework
- Windows COM automation via pywin32
- Python 3.10+ support
- Full docstring coverage
- Structured logging
- Robust error handling

---

## Future Roadmap

### Planned Features
- [ ] Attachment download/upload support
- [ ] Task management integration
- [ ] Folder management (create, move, delete)
- [ ] Advanced filtering (flags, categories, custom properties)
- [ ] Meeting response handling (accept/decline)
- [ ] Email rules management (create, modify, delete)
- [ ] Out-of-office settings
- [ ] Cross-platform support exploration

### Performance Improvements
- [ ] Async operations for better responsiveness
- [ ] Batch operations for multiple items
- [ ] Enhanced caching strategies
- [ ] Background sync capabilities

---

## Contributing

See [PUBLISHING_GUIDE.md](PUBLISHING_GUIDE.md) for information on how to contribute to this project.

---

[1.0.0]: https://github.com/YOUR_USERNAME/mcp-outlook/releases/tag/v1.0.0
