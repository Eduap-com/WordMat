Attribute VB_Name = "WindowsFunctions"
Const SW_SHOW = 1
Const SW_SHOWMAXIMIZED = 3
#If Mac Then
#Else
'Public Declare PtrSafe Sub Sleep Lib "kernel32" Alias "usleep" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal hwnd As LongPtr, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   Optional ByVal lpParameters As String, _
   Optional ByVal lpDirectory As String, _
   Optional ByVal nShowCmd As LongPtr) As LongPtr

Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
 
Sub Sleep2(ByVal WaitTime As Long)
'waittime is sekunder
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
  retval = ShellExecute(0, "open", FilePath, "", Mappe, SW_SHOWMAXIMIZED)

End Sub

Sub TestShellExecute()
  Dim retval As LongPtr
  
  On Error Resume Next
'  RetVal = ShellExecute(0, "open", "<full path to program>", "<arguments>", "<run in folder>", SW_SHOWMAXIMIZED)
  retval = ShellExecute(0, "open", Environ("TEMP") & "\" & "WordMatLaTex.pdf", "", "c:\", SW_SHOWMAXIMIZED)
End Sub
#End If
