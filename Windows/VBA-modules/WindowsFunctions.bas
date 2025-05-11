Attribute VB_Name = "WindowsFunctions"

#If Mac Then
#Else
'Public Declare PtrSafe Sub Sleep Lib "kernel32" Alias "usleep" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function SetForegroundWindow Lib "user32" _
        (ByVal hWnd As Long) As Long
#End If

Sub Sleep2(ByVal WaitTime As Long)
'waittime in seconds
On Error Resume Next
    Dim i As Long
    WaitTime = WaitTime * 100
    Do While i < WaitTime
        DoEvents
        Sleep 10
        i = i + 1
    Loop
End Sub

Sub RunDefaultProgram(FilePath As String, Optional Mappe As String = "c:\")
    On Error Resume Next
    PrepareMaxima False
    MaxProc.RunFile Mappe & "\" & FilePath, ""
End Sub

#End If

Public Sub SetExcelForeground()
#If Mac Then
#Else
    Dim hWnd As LongPtr
    hWnd = FindWindow("XLMAIN", "")
    If hWnd <> 0 Then SetForegroundWindow hWnd
#End If
End Sub
