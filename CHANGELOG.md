# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.2.1] - 2025-12-17

### Changed
- **Documentation Improvements**
  - Translated all documentation from French to English (DOCUMENTATION.md, CONTRIBUTING.md)
  - Replaced ASCII art architecture diagram with modern Mermaid diagram
  - Added emojis and soft colors (Material Design palette) to architecture diagram
  - Improved visual consistency across all documentation files

## [1.2.0] - 2025-12-17

### Added
- **Attachment Management** (3 new tools)
  - `get_email_attachments` - Get list of attachments from a specific email with details (filename, size, type, index)
  - `download_email_attachment` - Download specific attachment from an email to disk with automatic directory creation
  - `send_email_with_attachments` - Send emails with file attachments (single or multiple files)
  - Enhanced `format_email()` to include detailed attachment information in email metadata
  
- **Meeting Response Management** (2 new tools)
  - `get_meeting_requests` - Get pending meeting invitations that need a response with filtering by date range
  - `respond_to_meeting` - Accept, decline, or tentatively respond to meeting invitations with optional comments
  - Support for silent responses (update calendar without notifying organizer)
  
- **Out-of-Office Settings** (3 new tools)
  - `get_out_of_office_settings` - Get current automatic reply configuration
  - `set_out_of_office` - Configure automatic replies (immediate or scheduled) with separate messages for internal/external
  - `disable_out_of_office` - Disable automatic replies while preserving messages
  - Support for scheduling OOO with start/end times
  - Configurable external audience (None/Known/All)

### Changed
- Enhanced email metadata to include `entry_id` field for easier attachment and meeting operations
- Updated attachment list in email metadata to include filename, size, and type details
- **Complete Documentation Overhaul**
  - Created comprehensive technical documentation (DOCUMENTATION.md - 1500+ lines)
  - Created detailed contribution guide (CONTRIBUTING.md - 400+ lines)
  - Updated README.md with clear navigation and links to all documentation
  - Removed redundant documentation files (NEW_FEATURES.md, OPTIMIZATIONS.md, PUBLISHING_GUIDE.md)
  - Added cross-references between all documentation files
  - Enhanced QUICK_START.md and EXAMPLES.md integration

### Improved
- **Code Quality**
  - Removed unused helper functions (`get_outlook_signature`, `get_outlook_signature_via_display`)
  - Factored out duplicate code into `_set_email_body()` helper function
  - Optimized imports (removed unused `List` type)
  - Cleaned up misleading comments about Outlook format inheritance
  - Reduced codebase from 1981 to ~1840 lines (~140 lines saved)

### Notes
- Out-of-Office features require Outlook 2010+ with Exchange Server
- OOO settings may not be accessible via COM automation on all Outlook configurations
- All new features follow existing error handling patterns

## [1.1.0] - 2025-12-17

### Added
- **HTML Email Support**
  - `send_email` now accepts `html_body` parameter for rich HTML content
  - `create_draft_email` now accepts `html_body` parameter
  
- **Outlook Signature Integration**
  - `send_email` now accepts `signature_name` parameter to automatically include Outlook signatures
  - `create_draft_email` now accepts `signature_name` parameter
  - Automatic signature loading from user's Outlook Signatures folder
  - Preserves inline images and embedded content in signatures
  - Helper function `get_outlook_signature()` to load signature files

### Changed
- **Silent Logging**
  - Configured logging to CRITICAL level for minimal output
  - Added NullHandler to prevent log spam
  - Silenced all MCP/FastMCP internal loggers
  - Disabled log propagation for cleaner integration
  - Removed all informational log messages from tool execution

### Fixed
- Improved error handling in signature loading with fallback mechanisms
- Fixed inline image preservation when using Outlook signatures

## [1.0.0] - 2025-12-16

### Initial Public Release

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
- Folder caching (45x faster on repeated searches)
- Date filtering (search only recent emails, default: 2 days)
- Direct indexing (faster iteration without items.Count)
- Reduced limits (prevents long freezes, max 50 emails)
- Smart defaults optimized for daily usage

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

### Completed Features
- [x] Attachment download/upload support (v1.2.0)
- [x] Meeting response handling (accept/decline) (v1.2.0)
- [x] Out-of-office settings (v1.2.0)

### Planned Features
- [ ] Task management integration
- [ ] Folder management (create, move, delete)
- [ ] Advanced filtering (flags, categories, custom properties)
- [ ] Email rules management (create, modify, delete)
- [ ] Cross-platform support exploration
- [ ] Attachment content preview/thumbnails
- [ ] Bulk attachment operations

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
