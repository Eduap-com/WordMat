'' WordMat script
'' Mikael Sørensen, EDUAP
'' 10/12-2023
''

Option Explicit
Dim objFSO, objShell

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject( "WScript.Shell" )

' Define the the top level folder
Dim topLevelFolder
topLevelFolder = objShell.ExpandEnvironmentStrings("%PROGRAMFILES%") & "\Microsoft Office\root\Office16\STARTUP"

on error resume next
objFSO.DeleteFile(topLevelFolder & "\WordMat.dotm")

set objFSO = Nothing
set objShell = Nothing

'MsgBox "WordMat is now removed from Word Ribbon"

Wscript.Quit
