@echo off
:: Change to the directory where the batch file is located
cd /d "%~dp0Stuff"

:: Run the PowerShell script
powershell -ExecutionPolicy Bypass -File "Use.ps1"

:: Pause to keep the window open (optional)
pause