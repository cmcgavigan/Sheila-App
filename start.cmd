@echo off
REM start.cmd — launch the Sheila PWA server in the background (no console window).
REM Run this once to test, or let the Startup shortcut run it automatically at login.
cd /d "%~dp0"

if not exist node_modules\ (
    echo node_modules not found. Run:  npm install
    pause
    exit /b 1
)
if not exist .env (
    echo .env not found. Create it first (see README).
    pause
    exit /b 1
)

wscript "%~dp0launch.vbs"
echo Sheila server starting in the background. Output goes to server.log
