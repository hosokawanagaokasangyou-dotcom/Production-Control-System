@echo off
rem ASCII-only. Keep javafx-* version suffix in sync with pom.xml javafx.version (currently 26.0.1).
rem Do not paste into PowerShell; run: .\launch-pm-ai-desktop.bat
setlocal EnableExtensions EnableDelayedExpansion

set "ROOT=%~dp0"
if "%ROOT:~-1%"=="\" set "ROOT=%ROOT:~0,-1%"
cd /d "%ROOT%"

if not exist "%ROOT%\app" (
    echo [ERROR] Missing app folder. Put this bat next to PMD.exe / app / runtime.
    echo Current: "%ROOT%"
    pause
    exit /b 1
)

set "JAVA_EXE=%ROOT%\runtime\bin\java.exe"
if exist "%JAVA_EXE%" goto :have_java

if defined JAVA_HOME (
    if exist "%JAVA_HOME%\bin\java.exe" (
        set "JAVA_EXE=%JAVA_HOME%\bin\java.exe"
        echo [WARN] Using JAVA_HOME java.exe (bundled runtime missing).
        goto :have_java
    )
)

echo [ERROR] Java not found: "%ROOT%\runtime\bin\java.exe"
pause
exit /b 1

:have_java

rem OpenJFX Windows modular jars (must match files under app\ from package_app.ps1).
set "PM_AI_JFX_MODPATH=%ROOT%\app\javafx-base-26.0.1-win.jar;%ROOT%\app\javafx-controls-26.0.1-win.jar;%ROOT%\app\javafx-fxml-26.0.1-win.jar;%ROOT%\app\javafx-graphics-26.0.1-win.jar;%ROOT%\app\javafx-swing-26.0.1-win.jar"

"%JAVA_EXE%" -Dfile.encoding=UTF-8 -Xms3g -Xmx3g -XX:+HeapDumpOnOutOfMemoryError -XX:+UseStringDeduplication -Dprism.order=sw --add-opens=javafx.base/com.sun.javafx.event=ALL-UNNAMED --add-opens=javafx.controls/javafx.scene.control.skin=ALL-UNNAMED --add-exports=javafx.controls/com.sun.javafx.scene.control.behavior=ALL-UNNAMED --enable-native-access=javafx.graphics --module-path "%PM_AI_JFX_MODPATH%" --add-modules javafx.controls,javafx.fxml,javafx.graphics,javafx.base,javafx.swing -classpath "%ROOT%\app\*" jp.co.pm.ai.desktop.PmAiFxApp %*

set EXITCODE=!ERRORLEVEL!

if !EXITCODE! neq 0 (
    echo.
    echo [Exit !EXITCODE!] Logs: !USERPROFILE!\.pm-ai-desktop\startup.log  or  !TEMP!\pm-ai-desktop-startup.log
)

exit /b !EXITCODE!
