@echo off
setlocal EnableDelayedExpansion
set "MVN=%~1"
shift
set "ARGS="
:loop
if "%~1"=="" goto run
if /I "%~1"=="javafx:run" (
  rem compile first so validate build-classpath and target/classes are ready
  set "ARGS=!ARGS! compile exec:exec@pm-ai-desktop"
) else (
  set "ARGS=!ARGS! %~1"
)
shift
goto loop
:run
call "!MVN!" !ARGS!
exit /b %ERRORLEVEL%
