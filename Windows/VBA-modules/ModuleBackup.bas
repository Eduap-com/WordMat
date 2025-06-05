Attribute VB_Name = "ModuleBackup"
Option Explicit

Dim BackupAnswer As Integer
Dim SaveTime As Single

Sub SaveBackup()
    On Error GoTo Fejl
    Dim Path As String
    Dim UFbackup As UserFormBackup
    Dim UfWait As UserFormWaitForMaxima
    Const lCancelled_c As Long = 0
    Dim tempDoc2 As Document
    
    
    If BackupType = 2 Or BackupAnswer = 2 Then
        Exit Sub
    ElseIf BackupType = 0 And BackupAnswer = 0 Then
        Set UFbackup = New UserFormBackup
        UFbackup.Show
        '        If MsgBox(TT.A(179), vbYesNo, "Backup") = vbNo Then
        If UFbackup.Backup = False Then
            BackupAnswer = 2
            Exit Sub
        Else
            BackupAnswer = 1
        End If
    End If
        
    If timer - SaveTime < BackupTime * 60 Then Exit Sub
    SaveTime = timer
    If ActiveDocument.Path = "" Then
        MsgBox TT.A(679)
        Exit Sub
    End If
    Set UfWait = New UserFormWaitForMaxima
    UfWait.Show vbModeless
    UfWait.Label_tip.Caption = "Saving backup" ' to " & VbCrLfMac & "documents\WordMat-Backup"
    UfWait.Label_progress.Caption = "*"
    DoEvents
   
    
    '    Application.ScreenUpdating = False
    If ActiveDocument.Saved = False Then ActiveDocument.Save
    UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
    DoEvents
    BackupNo = BackupNo + 1
    If BackupNo > BackupMaxNo Then BackupNo = 1
#If Mac Then
    Path = GetTempDir & "WordMat-Backup/"
#Else
    Path = GetDocumentsDir & "\WordMat-Backup\"
#End If
    '    If Dir(path, vbDirectory) = "" Then MkDir path
    If Not fileExists(Path) Then MkDir Path
    UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
    DoEvents
    Path = Path & "WordMatBackup" & BackupNo & ".docx"
    If VBA.LenB(Path) = lCancelled_c Then Exit Sub
    
#If Mac Then
    Set tempDoc2 = Application.Documents.Add(Template:=ActiveDocument.FullName, visible:=False)
    UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
    DoEvents
    tempDoc2.ActiveWindow.Left = 2000
    tempDoc2.SaveAs Path
    UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
    DoEvents
    tempDoc2.Close
#Else
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ActiveDocument.Save
    fso.CopyFile ActiveDocument.FullName, Path
    Set fso = Nothing
    UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
    DoEvents
    UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
    DoEvents
#End If

    GoTo slut
Fejl:
    MsgBox TT.A(178), vbOKOnly, TT.A(208)
slut:
    On Error Resume Next
    If Not UfWait Is Nothing Then Unload UfWait
    Application.ScreenUpdating = True
End Sub

