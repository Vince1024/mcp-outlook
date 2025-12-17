# Contributing Guide - MCP Outlook

Thank you for your interest in contributing to MCP Outlook! This guide will help you participate in the project.

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [How to Contribute](#how-to-contribute)
- [Development Setup](#development-setup)
- [Code Standards](#code-standards)
- [Pull Request Process](#pull-request-process)
- [Roadmap](#roadmap)

---

## Code of Conduct

### Our Pledge

We are committed to making participation in this project a harassment-free experience for everyone, regardless of:
- Age
- Body size
- Disability
- Ethnicity
- Gender identity and expression
- Level of experience
- Nationality
- Personal appearance
- Race
- Religion
- Sexual identity and orientation

### Expected Behaviors

- Use welcoming and inclusive language
- Respect differing viewpoints and experiences
- Gracefully accept constructive criticism
- Focus on what is best for the community
- Show empathy towards other community members

### Unacceptable Behaviors

- Use of sexualized language or imagery
- Trolling, insulting or derogatory comments
- Public or private harassment
- Publishing others' private information without explicit permission
- Other conduct inappropriate in a professional setting

---

## How to Contribute

### Reporting Bugs

Before creating an issue:

1. **Check existing issues** to avoid duplicates
2. **Test with latest version**
3. **Reproduce the bug** reliably

#### Bug Issue Template

```markdown
**Bug Description**
Clear and concise description of the bug.

**Steps to Reproduce**
1. Go to '...'
2. Click on '...'
3. See error

**Expected Behavior**
What should happen.

**Actual Behavior**
What actually happens.

**Environment**
- OS: Windows 10/11
- Outlook Version: 2016/2019/365
- Python Version: 3.10/3.11/3.12
- MCP Outlook Version: 1.2.0

**Logs/Screenshots**
Output of `python tests/test_connection.py` and error messages.

**Additional Context**
Any other relevant information.
```

### Proposing New Features

#### Feature Issue Template

```markdown
**Problem/Need**
Describe the problem or need this feature would solve.

**Proposed Solution**
Describe the solution you envision.

**Alternatives Considered**
Other approaches you've considered.

**Impact**
- On existing users
- On performance
- On compatibility

**Implementation**
Implementation sketch if you have ideas.
```

### Improving Documentation

Documentation is as important as code!

Contributions welcome:
- Fix typos/grammar
- Add examples
- Clarify explanations
- Translate (if multilingual in future)

---

## Development Setup

### Prerequisites

- **Windows 10/11**
- **Microsoft Outlook** installed and configured
- **Python 3.10+**
- **Git** for version control

### Installation

```bash
# 1. Fork the project on GitHub

# 2. Clone your fork
git clone https://github.com/YOUR_USERNAME/mcp-outlook.git
cd mcp-outlook

# 3. Add upstream remote
git remote add upstream https://github.com/ORIGINAL_OWNER/mcp-outlook.git

# 4. Create a virtual environment
python -m venv venv
venv\Scripts\activate  # On Windows

# 5. Install dependencies
pip install -r requirements.txt

# 6. Install development dependencies
pip install pytest black ruff

# 7. Verify installation
python tests/test_connection.py
```

### Project Structure

```
mcp-outlook/
├── src/
│   ├── __init__.py
│   └── outlook_mcp.py       # Main MCP server
├── tests/
│   ├── __init__.py
│   ├── test_connection.py   # Outlook connection test
│   ├── test_outlook_mcp.py  # Unit tests
│   ├── test_advanced.py     # Advanced tests
│   └── test_tools.py        # Tool tests
├── docs/                    # Documentation (if needed)
├── .gitignore
├── pyproject.toml
├── requirements.txt
├── README.md
├── DOCUMENTATION.md
├── CONTRIBUTING.md          # This file
├── CHANGELOG.md
└── LICENSE
```

---

## Code Standards

### Python Style

This project follows **PEP 8** with some adjustments:

```python
# Line length: 100 characters (not 79)
# Use Black for automatic formatting

# Good
def my_function(param1: str, param2: int = 10) -> str:
    """
    Function description.
    
    Args:
        param1: Description of parameter 1
        param2: Description of parameter 2 (default: 10)
    
    Returns:
        Description of return value
    """
    return f"{param1}: {param2}"

# Bad
def myFunction(p1,p2=10):
    return f"{p1}: {p2}"
```

### Quality Tools

#### Black (Automatic Formatting)

```bash
# Format all code
black src/ tests/

# Check without modifying
black --check src/ tests/
```

#### Ruff (Linter)

```bash
# Lint
ruff check src/ tests/

# Auto-fix
ruff check --fix src/ tests/
```

### Naming Conventions

```python
# Variables and functions: snake_case
user_name = "John"
def get_inbox_emails(): ...

# Classes: PascalCase
class EmailManager: ...

# Constants: UPPER_SNAKE_CASE
MAX_EMAIL_LIMIT = 50
OUTLOOK_FOLDER_INBOX = 6

# Private: _ prefix
_FOLDER_CACHE = {}
def _get_folder_by_path(): ...
```

### Docstrings

Use **Google** style:

```python
def send_email(
    to: str,
    subject: str,
    body: str,
    cc: Optional[str] = None
) -> str:
    """
    Send an email via Outlook.
    
    Creates and sends a new email through the user's Outlook account.
    The email is sent immediately and a copy is saved in the Sent Items folder.
    
    Args:
        to: Recipient email address(es), semicolon-separated for multiple.
            Example: "user1@example.com" or "user1@example.com; user2@example.com"
        subject: Email subject line
        body: Email body content (plain text format)
        cc: CC recipients (optional), semicolon-separated
    
    Returns:
        JSON string with structure:
        {
            "success": bool,
            "message": str
        }
    
    Examples:
        >>> send_email("colleague@company.com", "Meeting", "See you at 2pm")
        {"success": true, "message": "Email sent to colleague@company.com"}
    
    Raises:
        ValueError: If Outlook is not accessible
    
    Notes:
        - Recipient addresses are logged but email content is not
        - BCC recipients are never logged for privacy
    """
    # Implementation
```

### Type Hints

Use type hints everywhere:

```python
from typing import Optional, Dict, Any, List

def format_email(mail_item) -> Dict[str, Any]:
    """Format an email item."""
    ...

def get_inbox_emails(
    limit: int = 10,
    unread_only: bool = False
) -> str:
    """Get inbox emails."""
    ...
```

### Error Handling

```python
# Consistent JSON return
try:
    # Code
    return json.dumps({
        "success": True,
        "data": "..."
    }, indent=2)
except Exception as e:
    logger.error("Failed to ...", exc_info=True, extra={
        "param1": value1,
        "param2": value2
    })
    return json.dumps({
        "success": False,
        "error": str(e)
    })
```

### Tests

Each new feature must have tests:

```python
# tests/test_new_feature.py
import pytest
from src.outlook_mcp import new_function

def test_new_function_success():
    """Test new_function with valid input."""
    result = new_function("valid_input")
    assert result["success"] is True

def test_new_function_error():
    """Test new_function with invalid input."""
    result = new_function("invalid_input")
    assert result["success"] is False
    assert "error" in result
```

---

## Pull Request Process

### Git Workflow

```bash
# 1. Sync with upstream
git fetch upstream
git checkout main
git merge upstream/main

# 2. Create a branch for your feature
git checkout -b feature/my-awesome-feature
# OR
git checkout -b fix/bug-description

# 3. Make your changes
# Edit files...

# 4. Test
python tests/test_connection.py
pytest tests/

# 5. Format and lint
black src/ tests/
ruff check --fix src/ tests/

# 6. Commit
git add .
git commit -m "feat: add awesome feature"
# OR
git commit -m "fix: resolve bug with email attachments"

# 7. Push
git push origin feature/my-awesome-feature

# 8. Create a Pull Request on GitHub
```

### Commit Convention

Use **Conventional Commits**:

```
<type>(<scope>): <description>

[optional body]

[optional footer]
```

#### Types

- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation only
- `style`: Formatting, indentation (no code change)
- `refactor`: Refactoring (no new feature or fix)
- `perf`: Performance improvement
- `test`: Adding or fixing tests
- `chore`: Maintenance (dependencies, config, etc.)

#### Examples

```bash
feat(email): add support for HTML email attachments
fix(calendar): resolve timezone issue in event creation
docs(readme): update installation instructions
refactor(contacts): simplify search logic
perf(folders): optimize folder cache lookup
test(email): add tests for send_email with attachments
chore(deps): update pywin32 to v306
```

### Pull Request Checklist

Before submitting a PR, check:

- [ ] Code follows style standards (Black + Ruff)
- [ ] Tests pass (`pytest tests/`)
- [ ] New tests added for new features
- [ ] Documentation updated (README, DOCUMENTATION, CHANGELOG)
- [ ] Docstrings are complete
- [ ] No commented code or debug prints
- [ ] Commits follow Conventional Commits
- [ ] PR has clear description

### Pull Request Template

```markdown
## Description

Brief description of changes.

## Type of Change

- [ ] Bug fix (non-breaking change that fixes an issue)
- [ ] New feature (non-breaking change that adds functionality)
- [ ] Breaking change (fix or feature that would break existing functionality)
- [ ] Documentation update

## How to Test

1. Step 1
2. Step 2
3. Expected result

## Checklist

- [ ] My code follows project standards
- [ ] I performed a self-review of my code
- [ ] I commented code in difficult parts
- [ ] I updated documentation
- [ ] My changes generate no new warnings
- [ ] I added tests that prove my fix/feature works
- [ ] Unit tests pass locally
- [ ] I updated CHANGELOG.md

## Screenshots (if applicable)

![Screenshot](url)

## Related Issues

Fixes #123
Relates to #456
```

---

## Roadmap

See [CHANGELOG.md](CHANGELOG.md) for complete roadmap.

### Current Priorities

#### High Priority
- [ ] Task management
- [ ] Folder management (create, move, delete)
- [ ] Advanced filters (flags, categories)

#### Medium Priority
- [ ] Email rules management (create, modify, delete)
- [ ] Attachment preview
- [ ] Batch operations

#### Low Priority
- [ ] Cross-platform support (explore MAPI alternatives)
- [ ] Web interface (optional)

### Completed Features

- [x] Email management (v1.0.0)
- [x] Calendar management (v1.0.0)
- [x] Contact management (v1.0.0)
- [x] Folder management (v1.0.0)
- [x] HTML email support (v1.1.0)
- [x] Outlook signature integration (v1.1.0)
- [x] Attachment management (v1.2.0)
- [x] Meeting response handling (v1.2.0)
- [x] Out-of-Office settings (v1.2.0)

---

## Frequently Asked Questions

### Q: My PR was rejected, what should I do?

**A**: Don't get discouraged! Read the reviewers' comments, make the requested changes, and resubmit. It's a learning process.

### Q: I don't know where to start?

**A**: Look for issues tagged `good first issue` or `help wanted`. These are good starting points for new contributors.

### Q: Can I contribute without coding?

**A**: Absolutely! You can:
- Improve documentation
- Translate (if multilingual)
- Report bugs
- Suggest improvements
- Help other users in issues

### Q: How do I test my changes?

**A**: 
1. Run `python tests/test_connection.py` for basic tests
2. Run `pytest tests/` for all tests
3. Test manually with a real Outlook
4. Check that existing functionality still works

### Q: My tests fail, what should I do?

**A**: 
1. Read error messages carefully
2. Check that Outlook is running
3. Check your Python environment
4. Ask for help in the issue or PR

---

## Resources

### External Documentation

- **Python**: https://docs.python.org/3/
- **FastMCP**: https://github.com/jlowin/fastmcp
- **pywin32**: https://github.com/mhammond/pywin32
- **Outlook COM API**: https://docs.microsoft.com/en-us/office/vba/api/overview/outlook
- **Model Context Protocol**: https://modelcontextprotocol.io

### Project Documentation

- [README.md](README.md) - Overview
- [DOCUMENTATION.md](DOCUMENTATION.md) - Complete technical documentation
- [QUICK_START.md](QUICK_START.md) - Quick start guide
- [EXAMPLES.md](EXAMPLES.md) - Usage examples
- [CHANGELOG.md](CHANGELOG.md) - Version history

---

## Acknowledgments

Thank you to all contributors who have made MCP Outlook what it is today!

### How to Be Listed

If you contribute significantly:
- Important bug fixes
- New features
- Documentation improvements
- Community help

Your name will be added to the acknowledgments section in the README!

---

## Contact

- **GitHub Issues**: [Create an issue](https://github.com/YOUR_USERNAME/mcp-outlook/issues)
- **Discussions**: [GitHub Discussions](https://github.com/YOUR_USERNAME/mcp-outlook/discussions)
- **Email**: For sensitive questions only

---

**Thank you for contributing to MCP Outlook!** 

Every contribution, small or large, makes a difference.

**Version**: 1.2.1  
**Last Updated**: December 17, 2025
