@echo off
rem Sheila App v2 - one-click installer. Right-clicking is not needed:
rem it relaunches itself elevated, sets up everything, and starts the service.
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Start-Process powershell -Verb RunAs -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File','\"%~dp0install\install.ps1\"'"
