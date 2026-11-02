@echo off
cd /d "%~dp0"
python release.py
if errorlevel 1 pause
