@echo off
setlocal EnableExtensions
rem マクロブックを、このバッチと同じフォルダをカレントにしたうえで別プロセスの Excel で開く
cd /d "%~dp0"
chcp 65001 >nul 2>&1

rem Excel へ相対パスだけ渡すと、起動プロセスの作業フォルダやパス解釈が「エクスプローラーから直接開く」とずれ、VBA の ThisWorkbook.Path 基準の処理やマクロで失敗することがある。
rem ブックは絶対パスで渡し、start /D で作業フォルダもブックと同じ場所に固定する（%~dp0 の末尾 \ と " のエスケープ避けに %~dp0. を使用）。
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

rem start: 1 つ目の引用符はウィンドウタイトル。/D で Excel プロセスのカレントをブックフォルダに固定。
start "" /D "%~dp0." "%EXCEL_EXE%" "%BOOK_ABS%"

endlocal
exit /b 0
