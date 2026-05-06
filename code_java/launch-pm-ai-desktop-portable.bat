@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem =============================================================================
rem  PmAiDesktop（jpackage app-image）の代替起動: exe の代わりに同梱 JRE で Java を直接起動する。
rem  配置: PmAiDesktop.exe / app / runtime / pm-ai-data と同じフォルダにこの bat を置く。
rem  （package_app.ps1 完了後の dist\PmAiDesktop\ に launch-pm-ai-desktop.bat としてコピーされます）
rem
rem  JVM オプションは package_app.ps1 の $javaOpts および pom の jvm.initial.heap / jvm.max.heap と揃えてください。
rem =============================================================================

set "ROOT=%~dp0"
if "%ROOT:~-1%"=="\" set "ROOT=%ROOT:~0,-1%"
cd /d "%ROOT%"

set "JAVA_EXE=%ROOT%\runtime\bin\java.exe"
if not exist "%JAVA_EXE%" (
    echo [ERROR] 見つかりません: "%JAVA_EXE%"
    echo この bat は app-image のルート（PmAiDesktop.exe と同じ階層）に置いてください。
    pause
    exit /b 1
)
rem 末尾 \ で閉じ引用符が壊れるため、ディレクトリは \ なしで検査する
if not exist "%ROOT%\app" (
    echo [ERROR] app フォルダがありません: "%ROOT%\app"
    pause
    exit /b 1
)

rem --- pom.xml の jvm.initial.heap / jvm.max.heap（既定 2g）と package_app.ps1 の javaOpts に合わせる ---
set "JAVA_OPTS=-Dfile.encoding=UTF-8 -Xms2g -Xmx2g -XX:+HeapDumpOnOutOfMemoryError -XX:+UseStringDeduplication -Dprism.order=sw --add-opens=javafx.base/com.sun.javafx.event=ALL-UNNAMED --add-opens=javafx.controls/javafx.scene.control.skin=ALL-UNNAMED --add-exports=javafx.controls/com.sun.javafx.scene.control.behavior=ALL-UNNAMED"

"%JAVA_EXE%" %JAVA_OPTS% -classpath "%ROOT%\app\*" jp.co.pm.ai.desktop.PmAiFxApp %*
set EXITCODE=!ERRORLEVEL!

if !EXITCODE! neq 0 (
    echo.
    echo [終了コード !EXITCODE!] ログ: %%USERPROFILE%%\.pm-ai-desktop\startup.log または %%TEMP%%\pm-ai-desktop-startup.log
)

exit /b !EXITCODE!
