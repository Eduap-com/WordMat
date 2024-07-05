@echo off
setlocal

set "APPDATA=%APPDATA%"
set "sourceFile=%APPDATA%\WordMat\WordMat.dotm"
set "sourceFolder=%APPDATA%\WordMat\"
set "topLevelFolder=%APPDATA%\Microsoft\Word"

if not exist "%sourceFile%" (
    set "sourceFile=C:\Program Files (x86)\WordMat\WordMat.dotm"
    set "sourceFolder=C:\Program Files (x86)\WordMat\"
)

if not exist "%sourceFile%" (
    echo WordMat.dotm could not be found
    exit /b
)

if not exist "%topLevelFolder%\STARTUP\" (
    mkdir "%topLevelFolder%\STARTUP"
)

if not exist "%topLevelFolder%\START\" (
    mkdir "%topLevelFolder%\START"
)

for /d %%i in ("%topLevelFolder%\*") do (
    copy /Y "%sourceFile%" "%%i\WordMat.dotm"
    copy /Y "%sourceFolder%WordMatP.dotm" "%%i\WordMatP.dotm"
    copy /Y "%sourceFolder%WordMatP2.dotm" "%%i\WordMatP2.dotm"
)

endlocal