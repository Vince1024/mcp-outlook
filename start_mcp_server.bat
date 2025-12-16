@echo off
REM Script to start the MCP Outlook Server
REM This script launches the Outlook MCP server that Cursor will connect to

echo ========================================
echo Starting MCP Outlook Server
echo ========================================
echo.

cd /d "%~dp0"

echo Current directory: %CD%
echo.

echo Checking Python installation...
python --version
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    pause
    exit /b 1
)
echo.

echo Starting MCP Outlook server...
echo Press Ctrl+C to stop the server
echo.

python src\outlook_mcp.py

pause

