'' OpretWordMenu.vbs
'' Mikael S�rensen, Nyborg Gymnasium
'' 20/8-2010
'' Bruges ikke l�ngere

Dim fso, objnet, filnavn, vistadir, XPdir, usrname, dir, strPath
dim strRegkey, startupmappe

set fso = Wscript.CreateObject("Scripting.FileSystemObject")
Set objNet = WScript.CreateObject("WScript.Network")
Set objShell = CreateObject("Wscript.Shell")
Set objSysEnv = objShell.Environment("Process") 

filnavn = "WordMat.dotm"
'dir ="%appdata%\Microsoft\Word\START\"

strRegKey="HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\AppData"
strRegKey2="HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Common\General\Startup"
on error resume next

startupmappe = objshell.regread(strRegKey2)
if startupmappe="" then
	strRegKey2="HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Common\General\Startup"
	startupmappe = objshell.regread(strRegKey2)
	if startupmappe="" then
		msgbox "Word 2007 eller 2010 er ikke installeret"
		Wscript.Quit
	end if
end if
dir = objshell.regread(strRegKey) & "\Microsoft\Word\" & startupmappe

'msgbox dir
on error resume next
err.clear

if NOT(fso.FileExists(filnavn)) then
  msgbox "Installationsfilen: " & filnavn & vbcrlf & "findes ikke", , "Ingen installation"
  wscript.quit
end if

objShell.run "cmd /K copy " & filnavn & " """ & dir & """",0 ,false   ' 1, true viser commdoprompt

if err.number=0 then
	msgbox "Menuen blev oprettet i Word. Genstart Word og se efter menuen tilf�jelsesprogrammer.",vbokonly,"F�rdig"
else
	msgbox "Der skete en fejl under installationen. Problemet kan skyldes en af to ting" & vbcrlf & vbcrlf &"1. Luk Word inden du installerer." & vbcrlf & "2. Det kan ogs� v�re et rettighedsproblem, er du administrator p� denne computer?" & vbcrlf & vbcrlf & "Du kan selv pr�ve at installere MarkMenu ved at kopiere filen '" & filnavn & "' til den mappe der �bnes nu", vbokonly,"Fejl ved installation"
	strPath = "explorer.exe /e," & dir
	objShell.Run strPath
end if

set fso = Nothing 
set objnet = Nothing
set objShell = Nothing

Wscript.Quit
