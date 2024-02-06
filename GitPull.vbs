' Dette git pull fra remote repository.

Option Explicit

Dim objShell
Set objShell = CreateObject("WScript.Shell")
objShell.Run "cmd /k git pull && pause && exit", 1, true

Set objShell = Nothing

WScript.Quit