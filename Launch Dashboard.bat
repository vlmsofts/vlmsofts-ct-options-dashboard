@echo off
cd /d "%~dp0"
start "" python app.py
timeout /t 4 /nobreak >nul
start "" http://127.0.0.1:5050
