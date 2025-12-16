@echo off
REM Installation script for MCP Outlook
REM Run this from the project directory

echo ====================================
echo  Installing MCP Outlook
echo ====================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.10 or higher from python.org
    pause
    exit /b 1
)

echo [1/4] Python found
python --version

echo.
echo [2/4] Installing dependencies...
pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)

echo.
echo [3/4] Installing pywin32 post-install...
python -c "import win32com.client" >nul 2>&1
if %errorlevel% neq 0 (
    echo Running pywin32 post-install script...
    python Scripts\pywin32_postinstall.py -install
)

echo.
echo [4/4] Verifying Outlook connection...
python -c "import win32com.client; outlook = win32com.client.Dispatch('Outlook.Application'); print('Outlook connection: OK')"

if %errorlevel% neq 0 (
    echo WARNING: Could not connect to Outlook
    echo Make sure Microsoft Outlook is installed and configured
    echo You can still complete the installation
)

echo.
echo ====================================
echo  Installation Complete!
echo ====================================
echo.
echo Next steps:
echo 1. Make sure Outlook is running
echo 2. Test the server: python src\outlook_mcp.py
echo 3. Configure in Cursor/Claude Desktop using mcp_config_example.json
echo.
pause

