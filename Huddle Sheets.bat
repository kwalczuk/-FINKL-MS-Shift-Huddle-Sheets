@echo off
rem Double-click to start, or drag the email (.eml/.msg), a .zip, a folder, or the attachments onto this file.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0HuddleSheets.ps1" %*
if errorlevel 1 pause
