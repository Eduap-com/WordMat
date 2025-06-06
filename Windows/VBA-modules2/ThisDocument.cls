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
' k�res kun hvis WordMat.dotm er �bnet direkte og ikke som global skabelon
#If Mac Then
#Else
    If DontRun Then Exit Sub
    If DateDiff("n", ExportDatetime, Now) > 2 Then ' hvis lige eksporteret, s� sp�rg ikke
        If MsgBox("Do you want to Export modules and create a backup?", vbYesNo, "Export?") = vbYes Then
            ExportAllModules
            BackupThisDocument
        End If
    End If
#End If
End Sub

Private Sub Document_Open()
    ' k�res kun hvis WordMat.dotm er �bnet direkte og ikke som global skabelon
    ' Der opst�r m�rkelig fejl p� Mac med denne funktion. F�rste gang der �bnes p� Mac fungerer det fint. Der gemmes.
    ' N�ste gang der �bnes, meldes 'Internal error' p� funktionen og intet virker.
    ' Det opst�r ikke, hvis der compiles n�r dokumentet �bnes f�rste gang.
    ' Problemet opst�r ogs� selvom denne funktion kun indeholder en msgbox "Hej"
    Dim s As String
    s = CheckKeyboardShortcutsPar(True)
    If s <> vbNullString Then
'       If MsgBox("Keyboard Genveje er ikke korrekte." & vbCrLf & s & vbCrLf & "Du skal g�re f�lgende: " & vbCrLf & "1. Compile p� mac (Alt+F11)" & vbCrLf & "2. K�r makro 'GenerateKeyboardShortcutsWordMat'" & vbCrLf & "3. K�r Macro 'Check KeyboardShortCuts' for at se om alt er korrekt" & vbCrLf & vbCrLf & "Vil du g�re det nu?", vbYesNo, "Shortcuts") = vbYes Then
       ' Hvis der ikke er compilet f�r man ikke ret mange makroer op med Alt+F8 p� Mac (Det g�r der nu. Det er et tidligere problem, m�ske cleaner virkede?)
       
' Denne metode fungerer ikke n�r WordMat ikke er compilet p� Mac. Genvejene bliver ikke tilknyttet korrekt. (Det g�r der nu. Det er et tidligere problem m�ske, cleaner virkede?)
'        If MsgBox("Keyboard Genveje er ikke korrekte." & vbCrLf & s & vbCrLf & "Vil du fixe?", vbYesNo) = vbYes Then
            GenerateKeyboardShortcutsWordMat
            CheckKeyboardShortcuts
            CompileMe
'            MsgBox "Keyboard Genveje er repareret." & vbCrLf & vbCrLf & "Glem ikke at compile p� Mac, med det samme.", vbOKOnly, "Husk compile!"
 '       End If
    End If
    
    '#If Mac Then
    '#Else ' s� l�nge der dun er �n udvikler er det ikke needed
    '    If MsgBox("Do you want to import modules?", vbYesNo, "Export?") = vbYes Then
    '        ImportAllModules
    '    End If
    '#End If

End Sub

Sub CompileMe()
#If Mac Then ' virker ikke p� mac
#Else
    Dim objVBECommandBar As Object, CompileMe As Variant
    Set objVBECommandBar = Application.VBE.CommandBars
    Set CompileMe = objVBECommandBar.FindControl(Type:=msoControlButton, Id:=578)
    CompileMe.Execute
#End If
End Sub
