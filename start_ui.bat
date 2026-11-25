@echo off
cd /d "%~dp0"
python scraper_ui.py
if errorlevel 1 pause
