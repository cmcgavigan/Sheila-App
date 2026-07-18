@echo off
REM ---------------------------------------------------------------------------
REM update-restart.cmd — relaunch helper for the in-app one-click update.
REM Invoked (detached) by server.js /api/update AFTER it has run `git pull`.
REM Waits for the old server to exit, optionally reinstalls dependencies, then
REM starts the app again via launch.vbs. Touches no data.
REM   Arg1 = "1" -> run `npm install` first (dependencies changed in the pull).
REM ---------------------------------------------------------------------------
cd /d "%~dp0"
timeout /t 4 /nobreak >nul
if "%~1"=="1" (
    echo [update] dependencies changed - running npm install... >> server.log 2>&1
    call npm install --no-audit --no-fund >> server.log 2>&1
)
wscript "%~dp0launch.vbs"
