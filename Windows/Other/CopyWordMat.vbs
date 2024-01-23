'' WordMat script
'' Mikael S�rensen, EDUAP
'' 10/12-2023
''

Option Explicit
Dim objFSO, objFolder, objSubFolder, objShell

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject( "WScript.Shell" )

' Define the source file and the top level folder
Dim sourceFile, topLevelFolder, sourceFolder
topLevelFolder = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Word"
sourceFile = "C:\Program Files (x86)\WordMat\WordMat.dotm"
sourceFolder = "C:\Program Files (x86)\WordMat\"
if Not objFSO.FileExists(sourceFile) Then
	sourceFile = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\WordMat\WordMat.dotm"
	sourceFolder = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\WordMat\"
End If
if Not objFSO.FileExists(sourceFile) Then
	MsgBox "WordMat.dotm could not be found"
	Wscript.Quit
End If

' Create the STARTUP and START folders if they don't exist
if Not objFSO.FolderExists(topLevelFolder & "\STARTUP") Then
	objFSO.CreateFolder(topLevelFolder & "\STARTUP")
End If

if Not objFSO.FolderExists(topLevelFolder & "\START") Then
	objFSO.CreateFolder(topLevelFolder & "\START")
End If

Set objFolder = objFSO.GetFolder(topLevelFolder)

on error resume next
    For Each objSubFolder in objFolder.Subfolders
		objFSO.CopyFile sourceFile, objSubFolder.Path & "\WordMat.dotm", TRUE
		objFSO.CopyFile sourceFolder & "WordMatP.dotm", objSubFolder.Path & "\WordMatP.dotm", TRUE
		objFSO.CopyFile sourceFolder & "WordMatP2.dotm", objSubFolder.Path & "\WordMatP.dotm", TRUE
    Next

set objFSO = Nothing
set objShell = Nothing

Wscript.Quit
