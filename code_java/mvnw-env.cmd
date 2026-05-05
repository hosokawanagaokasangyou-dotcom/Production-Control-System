@echo off
REM Maven Wrapper runs with JAVA_HOME from User/Machine registry (WSL PowerShell session fix).
REM Usage: mvnw-env.cmd javafx:run
setlocal
set "PS1=%~dp0mvnw-env-javahome.ps1"
if not exist "%PS1%" (
  echo [mvnw-env] missing: %PS1% 1>&2
  exit /b 1
)
set "JAVA_HOME="
for /f "usebackq delims=" %%A in (`powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%PS1%"`) do set "JAVA_HOME=%%A"
if not defined JAVA_HOME (
  exit /b 1
)
if not exist "%JAVA_HOME%\bin\java.exe" (
  echo [mvnw-env] java.exe not found: "%JAVA_HOME%\bin\java.exe" 1>&2
  exit /b 1
)
set "PATH=%JAVA_HOME%\bin;%PATH%"
call "%~dp0mvnw.cmd" %*
endlocal
