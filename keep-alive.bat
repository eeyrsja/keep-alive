@echo off

rem copy this file to startup folder to run keep-alive.py on Windows startup
rem startup folder path: %APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup

cd /d "C:\dev\keep-alive"
start /min "" ".venv\Scripts\pythonw.exe" keep-alive.py