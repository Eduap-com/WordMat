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
' køres kun hvis WordMat.dotm er åbnet direkte og ikke som global skabelon
#If Mac Then
#Else
    If DateDiff("n", ExportDatetime, Now) > 2 Then ' hvis lige eksporteret, så spørg ikke
        If MsgBox("Do you want to Export modules and create a backup?", vbYesNo, "Export?") = vbYes Then
            ExportAllModules
            BackupThisDocument
        End If
    End If
#End If
End Sub

Private Sub Document_Open()
    ' køres kun hvis WordMat.dotm er åbnet direkte og ikke som global skabelon
    ' Der opstår mærkelig fejl på Mac med denne funktion. Første gang der åbnes på Mac fungerer det fint. Der gemmes.
    ' Næste gang der åbnes, meldes 'Internal error' på funktionen og intet virker.
    ' Det opstår ikke, hvis der compiles når dokumentet åbnes første gang.
    ' Problemet opstår også selvom denne funktion kun indeholder en msgbox "Hej"
    Dim s As String
    s = CheckKeyboardShortcutsPar(True)
    If s <> vbNullString Then
       If MsgBox("Keyboard Genveje er ikke korrekte." & vbCrLf & s & vbCrLf & "Du skal gøre følgende: " & vbCrLf & "1. Compile på mac (Alt+F11)" & vbCrLf & "2. Kør makro 'GenerateKeyboardShortcutsWordMat'" & vbCrLf & "3. Kør Macro 'Check KeyboardShortCuts' for at se om alt er korrekt" & vbCrLf & vbCrLf & "Vil du gøre det nu?", vbYesNo, "Shortcuts") = vbYes Then
       ' Hvis der ikke er compilet får man ikke ret mange makroer op med Alt+F8 på Mac (Det gør der nu. Det er et tidligere problem, måske cleaner virkede?)
       
' Denne metode fungerer ikke når WordMat ikke er compilet på Mac. Genvejene bliver ikke tilknyttet korrekt. (Det gør der nu. Det er et tidligere problem måske, cleaner virkede?)
'        If MsgBox("Keyboard Genveje er ikke korrekte." & vbCrLf & s & vbCrLf & "Vil du fixe?", vbYesNo) = vbYes Then
            GenerateKeyboardShortcutsWordMat
            CheckKeyboardShortcuts
'            MsgBox "Keyboard Genveje er repareret." & vbCrLf & vbCrLf & "Glem ikke at compile på Mac, med det samme.", vbOKOnly, "Husk compile!"
        End If
    End If
    
    '#If Mac Then
    '#Else ' så længe der dun er én udvikler er det ikke needed
    '    If MsgBox("Do you want to import modules?", vbYesNo, "Export?") = vbYes Then
    '        ImportAllModules
    '    End If
    '#End If

End Sub
