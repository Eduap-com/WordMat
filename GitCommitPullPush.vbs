' Dette script syncer med remote repository.

Option Explicit

Dim objShell
Set objShell = CreateObject("WScript.Shell")

' Change directory to your repository
'objShell.Run "cmd /c cd C:\path\to\your\repository", 0, True

' User inputs the message for commit message
Dim commitMessage
commitMessage = InputBox("Enter a commit message:",  "Commit","WordMat updated")

if commitMessage = "" then
    msgbox "No commit message entered. Exiting.", vbokonly, "No commit message"
    WScript.Quit
end if  

' Git push
objShell.Run "cmd /k git add . && git commit -m """ & commitMessage & """ && git pull && git push && pause && Exit", 1, true

Set objShell = Nothing

WScript.Quit