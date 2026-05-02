@echo off
REM Production-Control-System 直下で実行: Maven test + pytest（プランの test.bat 相当）
setlocal EnableExtensions
cd /d "%~dp0"

echo == [1/2] Maven test (code_java) ==
call mvn -f code_java\pom.xml test
if errorlevel 1 exit /b 1

echo == [2/2] pytest (code/python/tests) ==
cd code\python
python -m pytest tests\ -q --tb=short %*
if errorlevel 1 exit /b 1

echo OK
