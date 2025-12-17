# Quick Start Guide - MCP Outlook

Get up and running with MCP Outlook in 5 minutes!

## Prerequisites Check

- [ ] Windows OS
- [ ] Microsoft Outlook installed and configured with at least one email account
- [ ] Python 3.10+ installed
- [ ] Outlook is currently running

## 📦 Installation (3 steps)

### Option A: Automatic (Recommended)

Double-click `install.bat` and follow the prompts.

### Option B: Manual

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Test connection
python test_connection.py

# 3. If tests pass, you're ready!
```

## 🧪 Test Your Installation

```bash
python test_connection.py
```

Expected output:
```
✓ PASS: Imports
✓ PASS: Outlook Connection
✓ PASS: Server File
```

## Running the Server

### Quick Run

Double-click `run_server.bat` or:

```bash
python src/outlook_mcp.py
```

### For Cursor

1. Open Cursor settings (Ctrl+Shift+P → "Preferences: Open Settings (JSON)")

2. Add MCP configuration:

```json
{
  "mcp": {
    "servers": {
      "outlook": {
        "command": "python",
        "args": [
          "C:/Users/YOUR_USERNAME/source/repos/MCP/src/outlook_mcp.py"
        ]
      }
    }
  }
}
```

3. Restart Cursor

4. Verify: You should see "outlook" in the MCP servers list

### For Claude Desktop

1. Open `%APPDATA%/Claude/claude_desktop_config.json`

2. Add:

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

3. Restart Claude Desktop

## Try These Commands

Once configured, try asking your AI assistant:

### Email Examples

```
"Show me my last 5 unread emails"
"Search for emails about 'project alpha'"
"Send an email to john@example.com with subject 'Meeting Follow-up'"
```

### Calendar Examples

```
"What meetings do I have this week?"
"Create a meeting for tomorrow at 2pm with the team"
"Find all meetings about 'sprint planning'"
```

### Contact Examples

```
"Find contact information for Jane Smith"
"Show me all contacts from Acme Corp"
"Create a new contact for John Doe"
```

## Troubleshooting

### "Unable to connect to Outlook"

**Solutions:**
1. Make sure Outlook is running
2. Open Outlook and check it's working normally
3. Try restarting Outlook
4. Run the script as Administrator

### "No module named 'win32com'"

**Solution:**
```bash
pip install pywin32
python Scripts/pywin32_postinstall.py -install
```

### "Permission denied" or COM errors

**Solutions:**
1. Run terminal/IDE as Administrator
2. Check Windows Defender or antivirus settings
3. Ensure Outlook is not blocked by IT policies

### Server starts but tools don't work

**Check:**
1. Is Outlook running?
2. Do you have an email account configured?
3. Check the terminal for error messages
4. Run `python test_connection.py` to diagnose

## Next Steps

- Read the full [README.md](README.md) for detailed documentation
- Check [CONTRIBUTING.md](CONTRIBUTING.md) if you want to add features
- View [CHANGELOG.md](CHANGELOG.md) for version history

## 🆘 Need Help?

1. Check the error message in the terminal
2. Run `python test_connection.py` for diagnostics
3. Review the troubleshooting section above
4. Check if Outlook is working normally outside the MCP
5. Create an issue with your error details

## Verification Checklist

Before asking for help, verify:

- [ ] Python 3.10+ is installed (`python --version`)
- [ ] Dependencies are installed (`pip list | findstr fastmcp`)
- [ ] Outlook is running and has an account configured
- [ ] You can access emails/calendar/contacts in Outlook manually
- [ ] Test script passes (`python test_connection.py`)
- [ ] You're using the correct Python path in MCP config
- [ ] You've restarted Cursor/Claude Desktop after config changes

---

**Time to first tool call**: ~5 minutes **Enjoy your AI-powered Outlook assistant!** 