Attribute VB_Name = "WindowsFunctions"

#If Mac Then
#Else
'Public Declare PtrSafe Sub Sleep Lib "kernel32" Alias "usleep" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

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
