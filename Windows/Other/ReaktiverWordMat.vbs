'' Reaktiver Deaktiverede tilf�jelsesprogrammer i Word
'' Mikael S�rensen, Nyborg Gymnasium
'' 6/1-2011
''

dim objShell, strRegkey

Set objShell = CreateObject("Wscript.Shell")

strRegKey="HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Resiliency\DisabledItems\"
strRegKey2="HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Word\Resiliency\DisabledItems\"

on error resume next

objshell.regdelete strRegKey
objshell.regdelete strRegKey2

msgbox "Alle deaktiverede tilf�jelsesprogrammer i Word er nu blevet aktiveret." & vbcrlf & vbcrlf & "Bem�rk at Tilf�jelsesprogrammer ogs� bare kan v�re inaktive." & vbcrlf & " Hvis WordMat bare er inaktivt er du n�dt til at �ndre indstillingen manuelt under:" & vbcrlf & " Filer / Indstillinger / Tilf�jelsesprogrammer / v�lg for neden �Word-tilf�jelsesprogrammer� og tryk udf�r. S�rg for at der er et flueben ud for WordMat.dotm. Tryk OK." , ,"Gennemf�rt"

set objShell = Nothing

Wscript.Quit
