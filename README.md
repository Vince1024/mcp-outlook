MCP Outlook

[![Python Version](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/platform-Windows-blue.svg)](https://www.microsoft.com/windows)
[![MCP](https://img.shields.io/badge/MCP-Compatible-green.svg)](https://modelcontextprotocol.io)

<img src="https://api.iconify.design/material-symbols/mail-outline.svg?color=%23004080" width="64" height="64" alt="mail icon"> 

A Model Context Protocol (MCP) server for Microsoft Outlook integration.

## Overview

This MCP server provides AI assistants with the ability to interact with Microsoft Outlook, including:

- **Email Management**: Read, search, send, and draft emails with HTML support and Outlook signatures
- **Calendar Management**: View, create, and search calendar events
- **Contact Management**: View, create, and search contacts

## Features

### Email Tools

- `get_inbox_emails` - Retrieve emails from inbox with filtering options
- `get_sent_emails` - Retrieve sent emails
- `search_emails` - Search emails across folders by subject, body, or sender
- `send_email` - Send emails with CC/BCC support, HTML content, and Outlook signature integration
- `create_draft_email` - Create draft emails without sending, with HTML and signature support

### Calendar Tools

- `get_calendar_events` - Get upcoming calendar events
- `create_calendar_event` - Create new calendar events with attendees
- `search_calendar_events` - Search events by subject or location

### Contact Tools

- `get_contacts` - Retrieve contacts with optional name filtering
- `create_contact` - Create new contacts
- `search_contacts` - Search contacts by name, email, or company

### Folder Tools

- `list_outlook_folders` - List all Outlook folders (ultra-fast, no item counts)
- `search_emails_in_custom_folder` - Search in specific custom folders with date filtering

## Performance Optimizations

This MCP has been heavily optimized for **large mailboxes** and to **minimize Outlook freezing**:

- **Folder caching** - 45x faster on repeated searches
- **Date filtering** - Search only recent emails (default: 2 days)
- **Direct indexing** - Faster iteration without `items.Count`
- **Reduced limits** - Prevents long freezes (max 50 emails)
- **Smart defaults** - Optimized for daily usage
- **Silent logging** - Minimal log output for cleaner integration

**See [OPTIMIZATIONS.md](OPTIMIZATIONS.md) for detailed performance information.**

## Installation

### Prerequisites

- **Windows OS** (required for COM automation)
- **Microsoft Outlook** installed and configured
- **Python 3.10+**

### Setup

1. Clone or download this repository

2. Install dependencies:

```bash
pip install -r requirements.txt
```

Or using the project file:

```bash
pip install -e .
```

3. Verify Outlook is running and configured with an account

## Usage

### Running the Server

Run the MCP server directly:

```bash
python src/outlook_mcp.py
```

Or using FastMCP's built-in CLI:

```bash
fastmcp run src/outlook_mcp.py
```

### Configuration for Cursor/Claude Desktop

Add this configuration to your MCP settings file:

**For Cursor** (`~/.cursor/mcp.json` or workspace settings):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": [
        "C:/Users/YOUR_USERNAME/source/repos/MCP/src/outlook_mcp.py"
      ],
      "env": {}
    }
  }
}
```

**For Claude Desktop** (`%APPDATA%/Claude/claude_desktop_config.json` on Windows):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": [
        "C:/Users/YOUR_USERNAME/source/repos/MCP/src/outlook_mcp.py"
      ]
    }
  }
}
```

**Important**: Replace `YOUR_USERNAME` with your actual Windows username.

### Testing the Server

You can test the server using FastMCP's interactive mode:

```bash
fastmcp dev src/outlook_mcp.py
```

This will open an interactive prompt where you can test the tools.

## Tool Examples

### Reading Emails

```python
# Get last 10 unread emails
get_inbox_emails(limit=10, unread_only=True)

# Search for emails about "meeting"
search_emails(query="meeting", folder="inbox", limit=20)
```

### Sending Emails

```python
# Send a simple email
send_email(
    to="colleague@company.com",
    subject="Meeting Follow-up",
    body="Hi, following up on our meeting...",
    importance="high"
)

# Send email with HTML content and Outlook signature
send_email(
    to="colleague@company.com",
    subject="Project Update",
    html_body="<h1>Update</h1><p>Here are the details...</p>",
    signature_name="My Signature"
)

# Create a draft with multiple recipients and signature
create_draft_email(
    to="team@company.com",
    subject="Project Update",
    body="Here's the latest update...",
    cc="manager@company.com",
    signature_name="My Signature"
)
```

### Calendar Management

```python
# Get next 7 days of events
get_calendar_events(days_ahead=7)

# Create a meeting
create_calendar_event(
    subject="Team Standup",
    start_time="2025-01-15 09:00",
    end_time="2025-01-15 09:30",
    location="Conference Room A",
    required_attendees="team@company.com",
    reminder_minutes=15
)

# Search for meetings
search_calendar_events(query="standup", days_range=30)
```

### Contact Management

```python
# Get all contacts
get_contacts(limit=50)

# Search for a contact
search_contacts(query="John Smith")

# Create a new contact
create_contact(
    full_name="Jane Doe",
    email="jane.doe@company.com",
    company="Acme Corp",
    job_title="Product Manager",
    mobile_phone="+1-555-1234"
)
```

## Security & Permissions

- This server requires access to your Outlook data
- It uses Windows COM automation (no credentials stored)
- All operations are performed with your current Outlook profile's permissions
- Make sure Outlook is running and configured before starting the server

## Troubleshooting

### "Unable to connect to Outlook"

- Ensure Microsoft Outlook is installed and running
- Verify Outlook is configured with at least one email account
- Try restarting Outlook

### "ImportError: No module named 'win32com'"

- Install pywin32: `pip install pywin32`
- After installation, run: `python Scripts/pywin32_postinstall.py -install` (if needed)

### Permission Errors

- Run your terminal/IDE as Administrator (may be required for COM automation)
- Check that Outlook is not blocked by security policies

### Date Parsing Issues

- Use ISO format for dates: `2025-01-15 14:00`
- Supported formats: "YYYY-MM-DD HH:MM", "tomorrow 2pm", "next Monday 10am"

## Development

### Project Structure

```
mcp-outlook/
├── src/
│   ├── __init__.py
│   └── outlook_mcp.py       # Main MCP server
├── pyproject.toml           # Project configuration
├── requirements.txt         # Dependencies
├── .gitignore
└── README.md
```

### Adding New Tools

To add a new tool, use the `@mcp.tool()` decorator:

```python
@mcp.tool()
def my_new_tool(param1: str, param2: int = 10) -> str:
    """
    Tool description.
    
    Args:
        param1: Description of param1
        param2: Description of param2 (default: 10)
    
    Returns:
        JSON string with results
    """
    # Implementation
    return json.dumps({"success": True, "data": "..."})
```

### Running Tests

```bash
pytest
```

### Code Formatting

```bash
black src/
ruff check src/
```

## Limitations

- **Windows Only**: Uses COM automation which is Windows-specific
- **Outlook Required**: Microsoft Outlook must be installed and running
- **Single Account**: Works with the default Outlook profile only
- **Performance**: Large mailboxes may have slower search performance
- **Attachments**: Current version doesn't support attachment handling (planned)

## Roadmap

### Recently Added
- [x] HTML email support
- [x] Outlook signature integration
- [x] Silent logging for cleaner integration
- [x] Outlook rules listing

### Planned Features
- [ ] Attachment download/upload support
- [ ] Task management integration
- [ ] Folder management (create, move, delete)
- [ ] Advanced filtering (flags, categories, custom properties)
- [ ] Meeting response handling (accept/decline)
- [ ] Email rules creation and modification
- [ ] Out-of-office settings
- [ ] Cross-platform support (investigate MAPI alternatives)

## License

MIT License - See LICENSE file for details.

## Contributing

Contributions are welcome! Please follow these guidelines:

1. Test your changes with a real Outlook installation
2. Follow the existing code style (Black + Ruff)
3. Add docstrings to all public functions
4. Update the README if adding new features
5. Submit a pull request with a clear description of changes

## Support

For issues or questions:
- Create an issue in the GitHub repository
- Check existing issues for similar problems
- Provide detailed information about your setup (Windows version, Outlook version, Python version)

## Acknowledgments

- Built with [FastMCP](https://github.com/jlowin/fastmcp)
- Uses [pywin32](https://github.com/mhammond/pywin32) for COM automation
- Inspired by the MCP Atlassian server architecture

---

**Note**: This tool accesses your local Outlook data. Ensure you follow your organization's security policies when handling email and calendar data.

