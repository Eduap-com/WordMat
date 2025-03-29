VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' OBS: This module cannot be removed. It is exported using ExportAllModules at backup, but cannot be importerted using ImportAllModules.
Option Explicit
Private DontRun As Boolean

Private Sub ActivateDontRun()
    DontRun = True
End Sub

Private Sub CommandButton_backup_Click()
    BackupThisDocument
End Sub

Private Sub CommandButton_ExportAndBackup_Click()
   ExportAllModules
   BackupThisDocument
End Sub

Private Sub CommandButton_runtest_Click()
    RunTestSequenceInteractive
End Sub

Private Sub CommandButton_runtestnoninteractive_Click()
    RunTestSequenceNonInteractive
End Sub

Private Sub CommandButtonExportModules_Click()
   ExportAllModules
End Sub

Private Sub Document_Close()
' Only run if WordMat.dotm is opened directly and not if global template
#If Mac Then
#Else
    If DontRun Then Exit Sub
    If DateDiff("n", ExportDatetime, Now) > 2 Then ' Dont ask, if just exported
        If MsgBox("Do you want to Export modules and create a backup?", vbYesNo, "Export?") = vbYes Then
            ExportAllModules
            BackupThisDocument
        End If
    End If
#End If
End Sub

Private Sub Document_Open()
' Only run if WordMat.dotm is opened directly and not if global template

' A strange error occurs on Mac with this function. The first time it is opened on Mac it works fine. It is saved.
' The next time it is opened, the function reports 'Internal error' and nothing works.
' It does not occur if the document is compiled when it is opened for the first time.
' The problem also occurs even if this function only contains a msgbox "Hello"
    Dim s As String
    s = CheckKeyboardShortcutsPar(True)
    If s <> vbNullString Then
            GenerateKeyboardShortcutsWordMat
            CheckKeyboardShortcuts
            CompileMe
    End If
    
    '#If Mac Then
    '#Else ' As long as there is only one developer, this is not needed
    '    If MsgBox("Do you want to import modules?", vbYesNo, "Export?") = vbYes Then
    '        ImportAllModules
    '    End If
    '#End If

End Sub

Sub CompileMe()
#If Mac Then ' Does not work on mac
#Else
    Dim objVBECommandBar As Object, CompileMe As Variant
    Set objVBECommandBar = Application.VBE.CommandBars
    Set CompileMe = objVBECommandBar.FindControl(Type:=msoControlButton, Id:=578)
    CompileMe.Execute
#End If
End Sub
