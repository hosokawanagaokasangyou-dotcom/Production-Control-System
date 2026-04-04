@echo off
setlocal EnableExtensions
rem マクロブックを、このバッチと同じフォルダをカレントにしたうえで別プロセスの Excel で開く
chcp 65001 >nul 2>&1

cd /d "%~dp0"

set "BOOK_REL=.\生産管理_AI配台テスト.xlsm"
if not exist "%BOOK_REL%" (
  echo ブックが見つかりません: %BOOK_REL%
  exit /b 1
)

set "EXCEL_EXE="
for %%E in (
  "%ProgramFiles%\Microsoft Office\root\Office16\EXCEL.EXE"
  "%ProgramFiles(x86)%\Microsoft Office\root\Office16\EXCEL.EXE"
  "%ProgramFiles%\Microsoft Office\root\Office15\EXCEL.EXE"
  "%ProgramFiles(x86)%\Microsoft Office\root\Office15\EXCEL.EXE"
  "%ProgramFiles%\Microsoft Office\Office16\EXCEL.EXE"
  "%ProgramFiles(x86)%\Microsoft Office\Office16\EXCEL.EXE"
) do if not defined EXCEL_EXE if exist "%%~E" set "EXCEL_EXE=%%~E"

if not defined EXCEL_EXE (
  for /f "delims=" %%W in ('where excel.exe 2^>nul') do (
    set "EXCEL_EXE=%%W"
    goto :excel_found
  )
)
:excel_found

if not defined EXCEL_EXE (
  echo EXCEL.EXE が見つかりません。Microsoft Office のインストールを確認してください。
  exit /b 1
)

rem start "" で新しいプロセスとして Excel を起動（既存ウィンドウに取り込まれる場合は Excel 側の設定を確認）
start "" "%EXCEL_EXE%" "%BOOK_REL%"

endlocal
exit /b 0
