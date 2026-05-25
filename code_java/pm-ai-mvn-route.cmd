@echo off
setlocal EnableDelayedExpansion
rem %~dp0 is only reliable here (before shift / call :label). Do not use %~dp0 in subroutines.
set "PM_AI_PROJECT_DIR=%~dp0"
if "!PM_AI_PROJECT_DIR:~-1!"=="\" set "PM_AI_PROJECT_DIR=!PM_AI_PROJECT_DIR:~0,-1!"
set "MVN=%~1"
shift
set "ARGS="
set "HAS_JAVAFX_RUN=0"
:loop
if "%~1"=="" goto after_loop
if /I "%~1"=="javafx:run" (
  set "HAS_JAVAFX_RUN=1"
) else (
  set "ARGS=!ARGS! %~1"
)
shift
goto loop
:after_loop
if "!HAS_JAVAFX_RUN!"=="1" (
  call :pm_ai_javafx_run
  exit /b !ERRORLEVEL!
)
call "!MVN!" !ARGS!
exit /b !ERRORLEVEL!

:pm_ai_javafx_run
call "!MVN!" compile
if !ERRORLEVEL! neq 0 exit /b !ERRORLEVEL!
call :pm_ai_verify_classes
if !ERRORLEVEL! neq 0 (
  echo [pm-ai-desktop] target\classes is incomplete. Running clean compile...
  call "!MVN!" clean compile
  if !ERRORLEVEL! neq 0 exit /b !ERRORLEVEL!
  call :pm_ai_verify_classes
  if !ERRORLEVEL! neq 0 (
    echo [pm-ai-desktop] ERROR: required .class / CSS still missing after compile.
    echo [pm-ai-desktop] Stop Java and Maven, then run: .\mvnw.cmd clean compile
    echo [pm-ai-desktop] Or use: .\run-pm-ai-desktop.ps1
    exit /b 1
  )
)
rem exec alone skips validate; module-path stays empty (see pom.xml comment).
call "!MVN!" validate exec:exec@pm-ai-desktop
exit /b !ERRORLEVEL!

:pm_ai_verify_classes
powershell -NoProfile -ExecutionPolicy Bypass -File "!PM_AI_PROJECT_DIR!\verify-pm-ai-build.ps1" -ProjectRoot "!PM_AI_PROJECT_DIR!"
exit /b !ERRORLEVEL!
