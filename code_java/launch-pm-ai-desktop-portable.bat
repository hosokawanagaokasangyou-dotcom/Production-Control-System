@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem =============================================================================
rem  PmAiDesktop（jpackage app-image）の代替起動: exe の代わりに Java で直接起動する。
rem  配置: PmAiDesktop.exe / app / runtime / pm-ai-data と同じフォルダにこの bat を置く。
rem  （package_app.ps1 完了後の dist\PmAiDesktop\ に launch-pm-ai-desktop.bat としてコピーされます）
rem
rem  必須: jpackage が生成した runtime フォルダ一式（容量大）。省略すると java.exe が無くなります。
rem  JVM オプションは package_app.ps1 の $javaOpts および pom の jvm.initial.heap / jvm.max.heap と揃えてください。
rem =============================================================================

set "ROOT=%~dp0"
if "%ROOT:~-1%"=="\" set "ROOT=%ROOT:~0,-1%"
cd /d "%ROOT%"

rem 末尾 \ で閉じ引用符が壊れるため、ディレクトリは \ なしで検査する
if not exist "%ROOT%\app" (
    echo [ERROR] app フォルダがありません。bat は dist\PmAiDesktop など app-image のルートに置いてください。
    echo 現在のフォルダ: "%ROOT%"
    pause
    exit /b 1
)

set "JAVA_EXE=%ROOT%\runtime\bin\java.exe"
if exist "%JAVA_EXE%" goto :have_java

rem 開発用: 同梱 runtime が無いときだけ JAVA_HOME（JDK フル、かつ pom の release と整合したもの）を使う
if defined JAVA_HOME (
    if exist "%JAVA_HOME%\bin\java.exe" (
        set "JAVA_EXE=%JAVA_HOME%\bin\java.exe"
        echo [WARN] 同梱の runtime\bin\java.exe がありません。JAVA_HOME の java を使います。
        echo        "%JAVA_EXE%"
        goto :have_java
    )
)

echo [ERROR] 同梱 JRE が見つかりません: "%ROOT%\runtime\bin\java.exe"
echo.
echo jpackage の app-image では「runtime」フォルダ全体が必要です（Git や ZIP で省略していないか確認）。
echo そのまま使う場合は Windows 上で code_java\package_app.ps1 を完走し、dist\PmAiDesktop を丸ごとコピーしてください。
echo.
echo このフォルダの一覧:
dir /b "%ROOT%"
echo.
if exist "%ROOT%\runtime" (
    echo runtime はありますが bin\java.exe がありません。ビルドやコピーが不完全な可能性があります。
) else (
    echo runtime フォルダ自体がありません。
)
echo.
pause
exit /b 1

:have_java

rem --- pom.xml の jvm.initial.heap / jvm.max.heap（既定 2g）と package_app.ps1 の javaOpts に合わせる ---
set "JAVA_OPTS=-Dfile.encoding=UTF-8 -Xms2g -Xmx2g -XX:+HeapDumpOnOutOfMemoryError -XX:+UseStringDeduplication -Dprism.order=sw --add-opens=javafx.base/com.sun.javafx.event=ALL-UNNAMED --add-opens=javafx.controls/javafx.scene.control.skin=ALL-UNNAMED --add-exports=javafx.controls/com.sun.javafx.scene.control.behavior=ALL-UNNAMED"

"%JAVA_EXE%" %JAVA_OPTS% -classpath "%ROOT%\app\*" jp.co.pm.ai.desktop.PmAiFxApp %*
set EXITCODE=!ERRORLEVEL!

if !EXITCODE! neq 0 (
    echo.
    echo [終了コード !EXITCODE!] ログ: %%USERPROFILE%%\.pm-ai-desktop\startup.log または %%TEMP%%\pm-ai-desktop-startup.log
)

exit /b !EXITCODE!
