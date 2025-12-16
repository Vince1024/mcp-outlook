"""
Quick test script to verify Outlook connection and basic functionality.
Run this before configuring the MCP server.
"""

import sys
import os

# Force UTF-8 encoding for Windows console
if sys.platform == 'win32':
    os.system('chcp 65001 > nul 2>&1')
    sys.stdout.reconfigure(encoding='utf-8')

def test_imports():
    """Test if required packages are installed."""
    print("Testing imports...")
    try:
        import win32com.client
        print("  [OK] win32com.client")
    except ImportError as e:
        print(f"  [FAIL] win32com.client - {e}")
        return False
    
    try:
        from dateutil import parser
        print("  [OK] dateutil")
    except ImportError as e:
        print(f"  [FAIL] dateutil - {e}")
        return False
    
    try:
        import fastmcp
        print("  [OK] fastmcp")
    except ImportError as e:
        print(f"  [FAIL] fastmcp - {e}")
        return False
    
    return True


def test_outlook_connection():
    """Test connection to Outlook."""
    print("\nTesting Outlook connection...")
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        print("  [OK] Connected to Outlook")
        
        # Try to get namespace
        namespace = outlook.GetNamespace("MAPI")
        print("  [OK] Got MAPI namespace")
        
        # Try to access inbox
        inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
        print(f"  [OK] Accessed Inbox ({inbox.Items.Count} items)")
        
        # Try to access calendar
        calendar = namespace.GetDefaultFolder(9)  # 9 = Calendar
        print(f"  [OK] Accessed Calendar ({calendar.Items.Count} items)")
        
        # Try to access contacts
        contacts = namespace.GetDefaultFolder(10)  # 10 = Contacts
        print(f"  [OK] Accessed Contacts ({contacts.Items.Count} items)")
        
        return True
        
    except Exception as e:
        print(f"  [FAIL] Failed to connect: {e}")
        print("\nTroubleshooting:")
        print("  1. Make sure Microsoft Outlook is installed")
        print("  2. Make sure Outlook is running")
        print("  3. Make sure you have an email account configured")
        print("  4. Try running this script as Administrator")
        return False


def test_mcp_server():
    """Test if the MCP server file is accessible."""
    print("\nTesting MCP server file...")
    try:
        from pathlib import Path
        server_file = Path("src/outlook_mcp.py")
        
        if server_file.exists():
            print(f"  [OK] Server file found: {server_file.absolute()}")
            return True
        else:
            print(f"  [FAIL] Server file not found: {server_file.absolute()}")
            return False
    except Exception as e:
        print(f"  [FAIL] Error checking server file: {e}")
        return False


def main():
    """Run all tests."""
    print("=" * 50)
    print("MCP Outlook - Connection Test")
    print("=" * 50)
    print()
    
    results = []
    
    # Test imports
    results.append(("Imports", test_imports()))
    
    # Test Outlook
    results.append(("Outlook Connection", test_outlook_connection()))
    
    # Test server file
    results.append(("Server File", test_mcp_server()))
    
    # Summary
    print("\n" + "=" * 50)
    print("Test Summary")
    print("=" * 50)
    
    all_passed = True
    for test_name, passed in results:
        status = "[PASS]" if passed else "[FAIL]"
        print(f"{status}: {test_name}")
        if not passed:
            all_passed = False
    
    print()
    if all_passed:
        print("SUCCESS! All tests passed! You're ready to use MCP Outlook.")
        print("\nNext steps:")
        print("1. Run: python src/outlook_mcp.py")
        print("2. Configure in Cursor/Claude Desktop")
    else:
        print("WARNING: Some tests failed. Please fix the issues above.")
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())

