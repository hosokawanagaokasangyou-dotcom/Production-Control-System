@echo off
REM Maven Wrapper を「ユーザー／マシン環境に保存済みの JAVA_HOME」で実行する。
REM WSL 連携の PowerShell など、セッションに古い JAVA_HOME が残る場合に使用する。
REM 使い方: mvnw-env.cmd javafx:run   /   mvnw-env.cmd -q test

for /f "usebackq delims=" %%A in (`powershell.exe -NoLogo -NoProfile -Command "$u=[Environment]::GetEnvironmentVariable('JAVA_HOME','User'); if ([string]::IsNullOrEmpty($u)) { $u=[Environment]::GetEnvironmentVariable('JAVA_HOME','Machine') }; if ($null -eq $u) { '' } elseif ($u.EndsWith('\')) { $u.Substring(0,$u.Length-1) } else { $u }"`) do set "JAVA_HOME=%%A"

if "%JAVA_HOME%"=="" (
  echo [mvnw-env] JAVA_HOME がユーザーまたはマシン環境にありません。 1>&2
  exit /b 1
)

if not exist "%JAVA_HOME%\bin\java.exe" (
  echo [mvnw-env] java.exe が見つかりません: "%JAVA_HOME%\bin\java.exe" 1>&2
  exit /b 1
)

set "PATH=%JAVA_HOME%\bin;%PATH%"
call "%~dp0mvnw.cmd" %*
exit /b %ERRORLEVEL%
