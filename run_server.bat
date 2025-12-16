@echo off
REM Quick launcher for MCP Outlook server

echo Starting MCP Outlook Server...
echo.
echo Make sure:
echo - Microsoft Outlook is running
echo - You have an email account configured
echo.

python src\outlook_mcp.py

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Server failed to start
    echo Check the error message above
    pause
)

