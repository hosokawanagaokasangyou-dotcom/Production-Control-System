@echo off
setlocal EnableExtensions
rem マクロブックを、このバッチと同じフォルダをカレントにしたうえで別プロセスの Excel で開く
cd /d "%~dp0"
chcp 65001 >nul 2>&1

rem NOTE: Long Japanese REM lines with % ~ " \ can break cmd.exe when the file is UTF-8 on a Japanese Windows (CP932).
rem       Keep technical notes ASCII-only below. Save as UTF-8 without BOM (BOM breaks the first line).
rem Relative-only paths can desync from VBA ThisWorkbook.Path handling; pass absolute path to the workbook.
rem Working directory is aligned with the workbook folder (start /D and trailing dot on batch dir).
set "BOOK_ABS=%~dp0生産管理_AI配台テスト.xlsm"
if not exist "%BOOK_ABS%" (
  echo ブックが見つかりません: %BOOK_ABS%
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

rem First quoted argument to CMD start is window title; /D fixes Excel process cwd to the workbook folder.
rem /x forces a new Excel.exe process (required when another workbook is already open in a different instance).
start "" /D "%~dp0." "%EXCEL_EXE%" /x "%BOOK_ABS%"

endlocal
exit /b 0
