VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' OBS: Dette modul kan ikke fjernes. Det bliver eksporteret ved ExportAllModules, men kan ikke importere ved ImportAllModules.
Option Explicit

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
    ' køres kun hvis WordMat.dotm er åbnet direkte og ikke som global skabelon
    If DateDiff("n", ExportDatetime, Now) > 2 Then ' hvis lige eksporteret, så spørg ikke
        If MsgBox("Do you want to Export modules?", vbYesNo, "Export?") = vbYes Then
            ExportAllModules
        End If
    End If
End Sub

