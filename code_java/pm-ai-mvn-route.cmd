@echo off
setlocal EnableDelayedExpansion
set "MVN=%~1"
shift
set "ARGS="
:loop
if "%~1"=="" goto run
if /I "%~1"=="javafx:run" (
  rem validate の build-classpath と target/classes を揃えるため compile を先行
  set "ARGS=!ARGS! compile exec:exec@pm-ai-desktop"
) else (
  set "ARGS=!ARGS! %~1"
)
shift
goto loop
:run
call "!MVN!" !ARGS!
exit /b %ERRORLEVEL%
