@echo off
setlocal

:: Delete registry keys
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems" /f
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Word\Resiliency\DisabledItems" /f
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Resiliency\DisabledItems" /f
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Word\Resiliency\DisabledItems" /f

:: Display message
echo Alle deaktiverede tilføjelsesprogrammer i Word er nu blevet aktiveret.
echo.
echo Bemærk at Tilføjelsesprogrammer også bare kan være inaktive.
echo Hvis WordMat bare er inaktivt er du nødt til at ændre indstillingen manuelt under:
echo Filer / Indstillinger / Tilføjelsesprogrammer / vælg for neden ’Word-tilføjelsesprogrammer’ og tryk udfør. Sørg for at der er et flueben ud for WordMat.dotm. Tryk OK.
pause

endlocal