@echo off
cd /d "%~dp0"
call mvnw.cmd -q "-Dtest=EquipmentGanttAssignmentDropSmokeTest" test
exit /b %ERRORLEVEL%
