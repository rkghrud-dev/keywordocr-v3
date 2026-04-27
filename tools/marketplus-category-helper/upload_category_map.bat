@echo off
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0marketplus_category_launcher.ps1"
start "" "http://127.0.0.1:5555/upload"
