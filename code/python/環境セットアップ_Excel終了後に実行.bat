@echo off
REM 別ウィンドウで PowerShell を起動し、Excel を保存確認付きで終了してから
REM setup_environment.py（pip / xlwings）を実行し、開いていたブックを開き直します。
chcp 65001 >nul
cd /d "%~dp0"
start "工程管理AI 環境セットアップ" powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0環境セットアップ_Excel終了後に実行.ps1"
