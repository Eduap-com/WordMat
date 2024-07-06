@echo off
setlocal

:: Check for administrative privileges
net session >nul 2>&1
if %errorlevel% == 0 (
    set "AdminPriv=true"
) else (
    set "AdminPriv=false"
)

:: Assuming WMall and AllOK are conditions you want to check
set "AllOK=true"

set "WMUserSTARTUPfolder=%APPDATA%\Microsoft\Word\STARTUP"
set "WMUserInstallFolder=%APPDATA%\WordMat"
set "WMallSTARTUPfolder=%PROGRAMFILES%\Microsoft Office\root\Office16\STARTUP"
set "WMallInstallFolder=%PROGRAMFILES(X86)%\WordMat""


:: Check if word is running
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>NUL | find /I /N "WINWORD.EXE">NUL
if "%ERRORLEVEL%"=="0" (
    :: force close word
::    taskkill /f /im WINWORD.EXE
    echo Word is running. Please close Word before uninstalling WordMat.
    set "AllOK=false"
    goto end
)


:: Delete WordMat from Start-menu
rmdir /s /q "%APPDATA%\Microsoft\Windows\Start Menu\Programs\WordMat"  

:: Delete WordMat from appdata
rmdir /s /q "%APPDATA%\WordMat"  

:: Delete WordMat for user
if exist "%WMUserSTARTUPfolder%\WordMat.dotm" (
    del /f /q "%WMUserSTARTUPfolder%\WordMat.dotm"
    del /f /q "%WMUserSTARTUPfolder%\WordMatP.dotm"
    del /f /q "%WMUserSTARTUPfolder%\WordMatP2.dotm"
)

:: Check if file "%WMallSTARTUPfolder%\WordMat.dotm" exists
set "WMall=false"
if exist "%WMallSTARTUPfolder%\WordMat.dotm" (
    set "WMall=true"
)
:: If installed for all users and not running as admin
if "%WMall%"=="true" (
    if "%AdminPriv%"=="false" (
        echo WordMat is installed for all users.
        echo WordMat can only be uninstalled by an administrator.
        echo Please run this script again as administrator to remove WordMat for all users.
        set "AllOK=false"
        goto end
    )
    del /f /q "%WMallSTARTUPfolder%\WordMat.dotm"
    del /f /q "%WMallSTARTUPfolder%\WordMatP.dotm"
    del /f /q "%WMallSTARTUPfolder%\WordMatP2.dotm"
    rmdir /s /q "%WMallInstallFolder%"
    if %errorlevel% == 0 (
        echo Uninstall complete. WordMat is now removed from the Ribbon in Word.
    ) else (
        echo An error occurred during uninstallation.
    )
)


:end
pause