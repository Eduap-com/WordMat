' Dette git pusher til remote repository.

Option Explicit

'dim objFSO, objFile
'SET objFSO = CREATEOBJECT("Scripting.FileSystemObject")

Dim objShell
Set objShell = CreateObject("WScript.Shell")

' Change directory to your repository
'objShell.Run "cmd /c cd C:\path\to\your\repository", 0, True

' Git push
objShell.Run "cmd /k git push && pause && exit", 1, true
'objShell.Run "cmd /k git pull", 1, True
'objShell.Run "cmd /k git push", 1, True



Set objShell = Nothing


'set objFSO = Nothing




WScript.Quit