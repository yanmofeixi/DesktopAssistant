@echo off
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install.ps1"
if errorlevel 1 (echo Installation failed. Read the error above or in the administrator window.) else (echo Installation finished.)
pause
