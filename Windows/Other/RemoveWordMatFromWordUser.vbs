'' WordMat script
'' Mikael Sørensen, EDUAP
'' 10/12-2023
''

Option Explicit
Dim objFSO, objFolder, objSubFolder, objShell

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject( "WScript.Shell" )

' Define the source file and the top level folder
Dim sourceFile, topLevelFolder
topLevelFolder = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Word"

on error resume next
objFSO.DeleteFile(topLevelFolder & "\STARTUP\WordMat.dotm")
objFSO.DeleteFile(topLevelFolder & "\START\WordMat.dotm")

Set objFolder = objFSO.GetFolder(topLevelFolder)

    For Each objSubFolder in objFolder.Subfolders
	objFSO.DeleteFile(objSubFolder.Path & "\WordMat.dotm")
    Next

set objFSO = Nothing
set objShell = Nothing

'MsgBox "WordMat is now removed from the Ribbon in Word"

Wscript.Quit
