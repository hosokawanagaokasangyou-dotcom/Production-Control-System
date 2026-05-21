@echo off
setlocal EnableDelayedExpansion
set "MVN=%~1"
shift
set "ARGS="
:loop
if "%~1"=="" goto run
if /I "%~1"=="javafx:run" (
  set "ARGS=!ARGS! exec:exec@pm-ai-desktop"
) else (
  set "ARGS=!ARGS! %~1"
)
shift
goto loop
:run
call "!MVN!" !ARGS!
exit /b %ERRORLEVEL%
