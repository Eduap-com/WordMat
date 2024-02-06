' Dette script renser WordMat.dotm for compileret kode mm.
' Det kræver at Ribbon Commander er installeret med licens
' Det er en hjælp hvis WordMat.dotm pludelig crasher ved åbning. 
' Der er andre løsninger, men denne er nemmest.
' Scriptet skal ligge i samme mappe som WordMat.dotm
' scriptet kopierer også de rensede filer til Mac-mappen

Option Explicit

'dim objFSO, objFile
'SET objFSO = CREATEOBJECT("Scripting.FileSystemObject")

Dim objShell
Set objShell = CreateObject("WScript.Shell")

' Change directory to your repository
'objShell.Run "cmd /c cd C:\path\to\your\repository", 0, True

' Git push
'objShell.Run "git commit -m ""Test""", 1, true ' true=wait for completion
objShell.Run "cmd /k git add . && git commit -m ""Test"" && git pull && git push && pause && Exit", 1, true
'objShell.Run "cmd /k git pull", 1, True
'objShell.Run "cmd /k git push", 1, True



Set objShell = Nothing


'set objFSO = Nothing




WScript.Quit