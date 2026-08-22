@echo off
setlocal

:: ============================================================
::  Aktiverer deaktiverede tilføjelsesprogrammer i Word
::
::  VIGTIGT: Denne fil SKAL gemmes som UTF-8 UDEN BOM.
::  (Notepad++: Kodning -> UTF-8 uden BOM)
:: ============================================================

:: --- Gem nuvaerende tegntabel, og skift til UTF-8 ---
set "_cp="
for /f "tokens=2 delims=:" %%A in ('chcp') do set "_cp=%%A"
set "_cp=%_cp: =%"
set "_cp=%_cp:.=%"
chcp 65001 >nul

set "_slettet=0"
set "_fejl=0"

:: --- Office-versioner der skal ryddes: 16.0=2016/2019/2021/365, 15.0=2013, 14.0=2010, 12.0=2007 ---
:: for %%V in (16.0 15.0 14.0 12.0) do (
::    call :SletNoegle "HKCU\Software\Microsoft\Office\%%V\Outlook\Resiliency\DisabledItems"
::)
call :SletNoegle "HKCU\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems"

echo.
if "%_fejl%"=="1" (
    echo ***************************************************************
    echo  ADVARSEL: Registreringsnøglen kunne ikke
    echo  slettes. Prøv at køre denne fil som administrator, eller
    echo  luk Word helt ned først.
    echo ***************************************************************
    echo.
) else if "%_slettet%"=="0" (
    echo Der var ingen deaktiverede tilføjelsesprogrammer at aktivere.
    echo.
) else (
    echo Alle deaktiverede tilføjelsesprogrammer i Word er nu blevet aktiveret.
    echo.
)

echo Bemærk at Tilføjelsesprogrammer også bare kan være inaktive.
echo Hvis WordMat bare er inaktivt er du nødt til at ændre indstillingen manuelt under:
echo Filer / Indstillinger / Tilføjelsesprogrammer / vælg for neden 'Word-tilføjelsesprogrammer'
echo og tryk udfør. Sørg for at der er et flueben ud for WordMat.dotm. Tryk OK.
echo.
pause

:: --- Gendan oprindelig tegntabel ---
if defined _cp chcp %_cp% >nul
endlocal
exit /b %_fejl%


:: ============================================================
::  Underrutine: sletter en registreringsnoegle med fejlhaandtering
::  %1 = fuld sti til noeglen (i anfoerselstegn)
:: ============================================================
:SletNoegle
reg query %1 >nul 2>&1
if errorlevel 1 (
    echo [ -  ] Findes ikke: %~1
    exit /b 0
)

reg delete %1 /f >nul 2>&1
if errorlevel 1 (
    echo [FEJL] Kunne ikke slettes: %~1
    set "_fejl=1"
    exit /b 1
)

echo [ OK ] Slettet: %~1
set "_slettet=1"
exit /b 0