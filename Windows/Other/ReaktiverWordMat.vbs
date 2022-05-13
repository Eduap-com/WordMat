'' Reaktiver Deaktiverede tilføjelsesprogrammer i Word
'' Mikael Sørensen, Nyborg Gymnasium
'' 6/1-2011
''

dim objShell, strRegkey

Set objShell = CreateObject("Wscript.Shell")

strRegKey="HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Resiliency\DisabledItems\"
strRegKey2="HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Word\Resiliency\DisabledItems\"
strRegKey3="HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Word\Resiliency\DisabledItems\"
strRegKey4="HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems\"

on error resume next

objshell.regdelete strRegKey4
objshell.regdelete strRegKey3
objshell.regdelete strRegKey2
objshell.regdelete strRegKey

msgbox "Alle deaktiverede tilføjelsesprogrammer i Word er nu blevet aktiveret." & vbcrlf & vbcrlf & "Bemærk at Tilføjelsesprogrammer også bare kan være inaktive." & vbcrlf & " Hvis WordMat bare er inaktivt er du nødt til at ændre indstillingen manuelt under:" & vbcrlf & " Filer / Indstillinger / Tilføjelsesprogrammer / vælg for neden ’Word-tilføjelsesprogrammer’ og tryk udfør. Sørg for at der er et flueben ud for WordMat.dotm. Tryk OK." , ,"Gennemført"

set objShell = Nothing

Wscript.Quit
