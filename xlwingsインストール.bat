@echo off
chcp 65001 >nul
REM Pip-install xlwings only. Min version 0.30 per requirements.txt.
REM Keep this .bat in code\python. Double-click to run.
setlocal
cd /d "%~dp0"
set PYTHONUTF8=1

where py >nul 2>&1
if %errorlevel% equ 0 (
  echo + py -m pip install -U "xlwings>=0.30"
  py -m pip install -U "xlwings>=0.30"
) else (
  echo + python -m pip install -U "xlwings>=0.30"
  python -m pip install -U "xlwings>=0.30"
)

set "ERR=%errorlevel%"
echo.
if not "%ERR%"=="0" (
  echo インストールに失敗しました（終了コード %ERR%）。
) else (
  echo xlwings のインストールが完了しました。
)
pause
exit /b %ERR%
