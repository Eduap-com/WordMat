@echo off

:: Check for administrator privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo This script requires administrator privileges. Please run as administrator.
    pause
    exit /b
)

xcopy "C:\GitHub\WordMat\Shared\Translations\win\*.csv" "C:\Program Files (x86)\WordMat\languages" /Y
