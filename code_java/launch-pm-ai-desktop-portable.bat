@echo off
rem ASCII-only: UTF-8 Japanese in .bat breaks cmd.exe parsing on many PCs.
rem Do not paste lines into PowerShell; double-click or: .\launch-pm-ai-desktop.bat
setlocal EnableExtensions EnableDelayedExpansion

rem Portable launcher: keep next to PmAiDesktop.exe, app\, runtime\, pm-ai-data\

set "ROOT=%~dp0"
if "%ROOT:~-1%"=="\" set "ROOT=%ROOT:~0,-1%"
cd /d "%ROOT%"

if not exist "%ROOT%\app" (
    echo [ERROR] Missing app folder. Put this bat in app-image root (e.g. dist\PmAiDesktop^).
    echo Current: "%ROOT%"
    pause
    exit /b 1
)

set "JAVA_EXE=%ROOT%\runtime\bin\java.exe"
if exist "%JAVA_EXE%" goto :have_java

if defined JAVA_HOME (
    if exist "%JAVA_HOME%\bin\java.exe" (
        set "JAVA_EXE=%JAVA_HOME%\bin\java.exe"
        echo [WARN] Bundled runtime\bin\java.exe missing; using JAVA_HOME.
        echo        "%JAVA_EXE%"
        goto :have_java
    )
)

echo [ERROR] Java not found: "%ROOT%\runtime\bin\java.exe"
echo.
echo Note: pm-ai-data\runtime is Python only. Need runtime\bin\java.exe beside this bat.
echo.
dir /b "%ROOT%"
pause
exit /b 1

:have_java

rem Heap/prism match pom.xml and package_app.ps1. Do NOT use --add-opens javafx.* here:
rem classpath-only JavaFX resolves too late and JDK prints "Unknown module: javafx.controls".

"%JAVA_EXE%" -Dfile.encoding=UTF-8 -Xms3g -Xmx3g -XX:+HeapDumpOnOutOfMemoryError -XX:+UseStringDeduplication -Dprism.order=sw -classpath "%ROOT%\app\*" jp.co.pm.ai.desktop.PmAiFxApp %*

set EXITCODE=!ERRORLEVEL!

if !EXITCODE! neq 0 (
    echo.
    echo [Exit !EXITCODE!] Logs: !USERPROFILE!\.pm-ai-desktop\startup.log  or  !TEMP!\pm-ai-desktop-startup.log
)

exit /b !EXITCODE!
