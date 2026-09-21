@echo off
cd /d "%~dp0"

where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw main.py
    exit
)

where pyw >nul 2>nul
if %errorlevel%==0 (
    start "" pyw main.py
    exit
)

start "" py main.py
exit