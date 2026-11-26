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
set "WMallInstallFolder=%PROGRAMFILES(X86)%\WordMat"

:: The startup folder Word uses, as the installers find it (2026-11-12):
::   WMStartupPath: Word's own setting (File, Options, Advanced, File Locations, Startup): a whole path, only there
::                  when the user has changed the folder. Word uses it first
::   WMStartupName: Office's name for the folder in %APPDATA%\Microsoft\Word (STARTUP; another name in some older
::                  translated versions of Office)
set "WMStartupPath="
for /f "tokens=2,*" %%A in ('reg query "HKCU\Software\Microsoft\Office\16.0\Word\Options" /v STARTUP-PATH 2^>nul ^| find "REG_"') do set "WMStartupPath=%%B"
set "WMStartupName="
for /f "tokens=2,*" %%A in ('reg query "HKCU\Software\Microsoft\Office\16.0\Common\General" /v Startup 2^>nul ^| find "REG_"') do set "WMStartupName=%%B"
:: (a part of a variable, %X:~0,2%, cannot be taken of a variable that is not set: the line would be an error and
:: stop the script, so the lines are jumped over then)
:: the path: environment variables in it expanded, no backslash at its end, and only used when it is a whole path
:: (X:\... or \\server\...)
if not defined WMStartupPath goto PathDone
call set "WMStartupPath=%WMStartupPath%"
if "%WMStartupPath:~-1%"=="\" set "WMStartupPath=%WMStartupPath:~0,-1%"
set "WMPathOK="
if "%WMStartupPath:~1,2%"==":\" set "WMPathOK=1"
if "%WMStartupPath:~0,2%"=="\\" set "WMPathOK=1"
if not defined WMPathOK set "WMStartupPath="
:PathDone
:: the name: only used when it is a folder name
if not defined WMStartupName goto NameDone
set "WMNameOK=1"
if not "%WMStartupName%"=="%WMStartupName:\=%" set "WMNameOK="
if not "%WMStartupName%"=="%WMStartupName:/=%" set "WMNameOK="
if not "%WMStartupName%"=="%WMStartupName::=%" set "WMNameOK="
if "%WMStartupName%"==".." set "WMNameOK="
if "%WMStartupName%"=="." set "WMNameOK="
if not defined WMNameOK set "WMStartupName="
:NameDone


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

:: Delete WordMat for user: WordMat classic (WordMat.dotm, WordMatP.dotm, WordMatP2.dotm) and WordMat+ (only WordMatP.dotm),
:: in STARTUP, in START (earlier installers also wrote there), and in the startup folder Word uses when it is another one
call :DelWordMat "%WMUserSTARTUPfolder%"
call :DelWordMat "%APPDATA%\Microsoft\Word\START"
if defined WMStartupName call :DelWordMat "%APPDATA%\Microsoft\Word\%WMStartupName%"
if defined WMStartupPath call :DelWordMat "%WMStartupPath%"

:: Check if file "%WMallSTARTUPfolder%\WordMat.dotm" exists
set "WMall=false"
if exist "%WMallSTARTUPfolder%\WordMat.dotm" (
    set "WMall=true"
)
:: WordMat+ has no WordMat.dotm
if exist "%WMallSTARTUPfolder%\WordMatP.dotm" (
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
exit /b 0

:: Deletes WordMat's templates in one folder (%1)
:DelWordMat
if "%~1"=="" exit /b 0
for %%T in (WordMat.dotm WordMatP.dotm WordMatP2.dotm) do (
    if exist "%~1\%%T" del /f /q "%~1\%%T"
)
exit /b 0
