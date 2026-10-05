@echo off
:: Exchange AutoDiscover Tester - Launcher
:: Double-click this file to start the tool.
:: Requires: Windows PowerShell 5.1 (built into Windows 10/11)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0ExchangeTester.ps1"

:: If PowerShell exits with an error, keep the window open to show the message
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Script exited with code %ERRORLEVEL%
    echo Make sure ExchangeTester.ps1 is in the same folder as this .bat file.
    pause
)
