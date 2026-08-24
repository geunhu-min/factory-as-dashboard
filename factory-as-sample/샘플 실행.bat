@echo off
cd /d "%~dp0"
wscript "%~dp0run-server-hidden.vbs"
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:5175"
