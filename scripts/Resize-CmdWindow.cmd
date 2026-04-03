@echo off
REM Same console as this cmd: GetConsoleWindow targets this window.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Resize-CmdWindow.ps1"
exit /b %ERRORLEVEL%
