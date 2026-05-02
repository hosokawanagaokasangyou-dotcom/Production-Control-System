@echo off
REM Run from Production-Control-System root: Maven Wrapper (mvnw) + pytest
setlocal EnableExtensions
cd /d "%~dp0"

echo == [1/2] Maven test (code_java) ==
pushd code_java
call mvnw.cmd test
set ERR=%ERRORLEVEL%
popd
if not %ERR%==0 exit /b %ERR%

echo == [2/2] pytest (code/python/tests) ==
cd code\python
python -m pytest tests\ -q --tb=short %*
if errorlevel 1 exit /b 1

echo OK
