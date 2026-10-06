@echo off
set "SCRIPT=%~dp0Disparo-Outlook.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%SCRIPT%"
