'' WordMat script
'' Mikael Sørensen, EDUAP
'' 10/12-2023
''

Option Explicit
Dim objFSO, objFolder, objShell
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

CopyFileToSubfolders sourceFile, topLevelFolder

set objFSO = Nothing
set objShell = Nothing

Wscript.Quit

' Recursive subroutine to copy the file to the folder and its subfolders
Sub CopyFileToSubfolders(sourceFile, folderPath)
    Dim objFolder, objSubFolder, objFile
    Set objFolder = objFSO.GetFolder(folderPath)

    ' Copy the file to the current folder
 '   objFSO.CopyFile sourceFile, objFolder.Path & "\"

    ' Recurse through each subfolder
    For Each objSubFolder in objFolder.Subfolders
'        CopyFileToSubfolders sourceFile, objSubFolder.Path
		objFSO.CopyFile sourceFile, objSubFolder.Path & "\"
    Next
End Sub


