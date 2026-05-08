@echo off
echo Starting script...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0main.ps1"
echo Script finished or crashed
pause