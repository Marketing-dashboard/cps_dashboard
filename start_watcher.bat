@echo off
:: Launches the Excel file watcher in a visible PowerShell window.
:: Double-click this file (or add it to Windows Startup) to start monitoring.

set SCRIPT=%~dp0watch_and_push.ps1

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
pause
