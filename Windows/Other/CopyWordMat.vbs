'' WordMat script
'' Mikael S�rensen, EDUAP
'' 10/12-2023
''

Option Explicit
Dim objFSO, objFolder, objShell
Dim objSubFolder, objFile

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject( "WScript.Shell" )

' Define the source file and the top level folder
Dim sourceFile, topLevelFolder
topLevelFolder = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Word"
sourceFile = "C:\Program Files (x86)\WordMat\WordMat.dotm"
if Not objFSO.FileExists(sourceFile) Then
	sourceFile = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\WordMat\WordMat.dotm"
End If
if Not objFSO.FileExists(sourceFile) Then
	MsgBox "WordMat.dotm could not be found"
	Wscript.Quit
End If

if Not objFSO.FolderExists(topLevelFolder & "\STARTUP") Then
	objFSO.CreateFolder(topLevelFolder & "\STARTUP")
End If

if Not objFSO.FolderExists(topLevelFolder & "\START") Then
	objFSO.CreateFolder(topLevelFolder & "\START")
End If

Set objFolder = objFSO.GetFolder(topLevelFolder)

    For Each objSubFolder in objFolder.Subfolders
		objFSO.CopyFile sourceFile, objSubFolder.Path & "\WordMat.dotm", TRUE
    Next

set objFSO = Nothing
set objShell = Nothing

Wscript.Quit
