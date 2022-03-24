VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormMaximaSettings 
   Caption         =   "Indstillinger"
   ClientHeight    =   5340
   ClientLeft      =   -15
   ClientTop       =   45
   ClientWidth     =   9780.001
   OleObjectBlob   =   "UserFormMaximaSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormMaximaSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False










Option Explicit
Private MustRestart As Boolean
Private LoadUnits As Boolean
Private UserUnits As Boolean
Private LangChange As Boolean

Private Sub CheckBox_complex_Change()
    If CheckBox_complex.Value Then
        CheckBox_polaroutput.visible = True
    Else
        CheckBox_polaroutput.visible = False
    End If
End Sub

Private Sub ComboBox_language_Change()
    LangChange = True
End Sub

Private Sub CommandButton_annuller_Click()
    Me.hide
End Sub

Private Sub NulstilWordDoc(filnavn As String)
Dim appdir As String
Dim fs, f, s
On Error GoTo fejl
    
If MsgBox(Sprog.A(665) & ", " & filnavn & " ," & Sprog.A(666), vbYesNo, Sprog.A(667)) = vbNo Then GoTo slut
    
#If Mac Then
' cant
'    FileCopy "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/" & filnavn, "~/Library/Containers/com.microsoft.Word/Data/WordMat/" & Replace(filnavn, "\", "/")
#Else
Set fs = CreateObject("Scripting.FileSystemObject")
filnavn = GetProgramFilesDir & "\WordMat\" & filnavn
appdir = Environ("AppData")
If Dir(filnavn) <> "" And appdir <> "" Then
  fs.CopyFile filnavn, appdir & "\WordMat\"
  MsgBox Sprog.A(668) & " " & filnavn & " " & Sprog.A(669), vbOKOnly, Sprog.A(670)
Else
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error '"Filen " & filnavn & " kunne ikke findes eller appdata mappen kunne ikke findes"
End If
#End If

GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:


End Sub
Private Sub CommandButton_nulstilfigurer_Click()

NulstilWordDoc ("Figurer.docx")
NulstilWordDoc ("Figurer_english.docx")
NulstilWordDoc ("Figurer_spansk.docx")

End Sub

Private Sub CommandButton_nulstilfysik_Click()

NulstilWordDoc ("Formulae\FysikFormler.docx")
NulstilWordDoc ("Formulae\FysikFormler_spansk.docx")

End Sub

Private Sub CommandButton_nulstilkemiformler_Click()

NulstilWordDoc ("Formulae\KemiFormler.docx")
NulstilWordDoc ("Formulae\KemiFormler_spansk.docx")

End Sub

Private Sub CommandButton_nulstilmatformler_Click()
NulstilWordDoc ("Formulae\MatFormler.docx")
NulstilWordDoc ("Formulae\MatFormler_english.docx")
NulstilWordDoc ("Formulae\MatFormler_spansk.docx")
End Sub

Private Sub CommandButton_ok_Click()
On Error Resume Next
    Dim xact As String, sett As String, gangetegn As String, lm As String, logout As String
    Dim knap As CommandBarButton
    Dim UnitChanged As Boolean
    
    If InStr(TextBox_outunits.text, "/") > 0 Or InStr(TextBox_outunits.text, "*") > 0 Or InStr(TextBox_outunits.text, "^") > 0 Then
        MsgBox Sprog.A(343)
        MultiPage1.Value = 2
        TextBox_outunits.SetFocus
        Exit Sub
    End If
    
    If OptionButton_prik.Value Then
        MaximaGangeTegn = "prik"
    ElseIf OptionButton_kryds.Value Then
        MaximaGangeTegn = "x"
    Else
        MaximaGangeTegn = "*"
    End If
    
    
    
'    If OptionButton_punktum.value Then ' erstattet af simplere nedenfor
'        MaximaSeparator = True
'    Else
'        MaximaSeparator = False
'    End If
    
    If OptionButton_numonly.Value = True Then
        MaximaExact = 2
    ElseIf OptionButton_exactonly.Value = True Then
        MaximaExact = 1
    Else
        MaximaExact = 0
    End If
    
    LmSet = OptionButton_lmset.Value
        
    If OptionButton_logauto.Value = True Then
        MaximaLogOutput = 0
    ElseIf OptionButton_log10.Value = True Then
        MaximaLogOutput = 2
    Else
        MaximaLogOutput = 1
    End If
    
    If OptionButton_gnuplot.Value Then
        GraphApp = 0
    ElseIf OptionButton_graph.Value Then
        GraphApp = 1
    ElseIf OptionButton_geogebra.Value Then
        GraphApp = 2
    ElseIf OptionButton_excel.Value Then
        GraphApp = 3
    ElseIf OptionButton_geogebraweb.Value Then
        GraphApp = 4
    End If
    
    If OptionButton_placementright.Value = True Then
        EqNumPlacement = False
    Else
        EqNumPlacement = True
    End If
    
    If OptionButton_eqnumone.Value = True Then
        EqNumType = False
    Else
        EqNumType = True
    End If
    
    If OptionButton_casmaxima.Value = True Then
        CASengine = 0
    ElseIf OptionButton_casgeogebra.Value = True Then
        CASengine = 1
    ElseIf OptionButton_casgeogebradirect.Value = True Then
        CASengine = 2
    Else
        CASengine = 0
    End If

    
    LanguageSetting = ComboBox_language.ListIndex
    Sprog.LoadSprogArray
    
    MaximaForklaring = CheckBox_Insertforklaring.Value
    MaximaKommando = CheckBox_Insertmaximacommand.Value
    Radians = OptionButton_radianer.Value
    MaximaCifre = ComboBox_cifre.Value
    MaximaSeparator = OptionButton_punktum.Value
    MaximaComplex = CheckBox_complex.Value
    LmSet = OptionButton_lmset.Value
    MaximaVidNotation = CheckBox_vidnotation.Value
    ExcelIndlejret = CheckBox_indlejret.Value
    AllTrig = OptionButton_trigall.Value
    AutoStart = CheckBox_autostart2.Value
    SettCheckForUpdate = CheckBox_checkupdate.Value
    MaximaIndex = OptionButton_indexvar.Value
    MaximaBigFloat = CheckBox_bigfloat.Value
    ShowAssum = CheckBox_showassum.Value
    PolarOutput = CheckBox_polaroutput.Value
    dAsDiffChr = CheckBox_dasdiffchr.Value
    EqAskRef = CheckBox_askref.Value
    BackupType = ComboBox_backup.ListIndex
    BackupMaxNo = ComboBox_backupno.text
    BackupTime = ComboBox_backuptime.text
    
    If MaximaUnits <> CheckBox_units.Value Then
        MaximaUnits = CheckBox_units.Value
        UnitChanged = True
        If CheckBox_units.Value Then
            LoadUnits = True
        Else
            MustRestart = True
        End If
    End If
    
    If OutUnits <> TextBox_outunits.text Then
        OutUnits = TextBox_outunits.text
        UserUnits = True
    End If
        
    SetMathAutoCorrect
    
    UFMSettings.hide
    Sprog.CheckSetting
    RibbonSubs.RefreshRibbon
    GenerateKeyboardShortcuts
'#If Mac Then
'    If LangChange Then MsgBox Sprog.A(671)
'#End If

#If Mac Then ' fjernet v. 1.23
' håndteres nu af unitimage for windows
''        If MustRestart Then
''            TurnUnitsOff
''        ElseIf LoadUnits And MaximaUnits Then
''            TurnUnitsOn
''        ElseIf UserUnits And MaximaUnits Then
''            UpdateUnits
''        End If
'        If MustRestart Then
'            MaxProc.CloseProcess
'            If MaximaUnits Then
'                MaxProc.Units = 1
'                MaxProc.OutUnits = omax.ConvertUnits(OutUnits)
'            Else
'                MaxProc.Units = 0
'            End If
'            MaxProc.StartMaximaProcess
'        ElseIf LoadUnits And MaximaUnits Then
'            TurnUnitsOn
'        ElseIf UserUnits And MaximaUnits Then
'            MaxProc.OutUnits = omax.ConvertUnits(OutUnits)
'            MaxProc.UpdateUnits
'        End If
'    End If
#End If
End Sub


Private Sub CommandButton_openbackup_Click()
On Error Resume Next
#If Mac Then
    RunScript "OpenFinder", GetTempDir() & "/WordMat-backup"
#Else
    Shell "explorer.exe" & " " & GetDocumentsDir & "\WordMat-Backup", vbNormalFocus
#End If
End Sub

Private Sub CommandButton_restartmaxima_Click()
    RestartWordMat
End Sub

Private Sub CommandButton_shortcuts_Click()
Dim Result As VbMsgBoxResult

'    result = MsgBox(Sprog.A(63) & vbCrLf & vbCrLf & "Alt + g    Prik-gangetegn" & vbCrLf & "Alt + b    beregn" & vbCrLf & "AltGr + Enter    beregn" & vbCrLf & "Alt + L    Løs ligning(er)" & vbCrLf & "Alt + S    Slet definitioner" & vbCrLf & "Alt + D    Indsæt Definer:" & vbCrLf & "Alt + i    Indstillinger" & vbCrLf & "alt + r    Indsæt forrige resultat(er)" & vbCrLf & "alt + P    Plot graf" & vbCrLf & "alt + E    Slå enheder til / fra" & vbCrLf & "alt + M    Indsæt ny ligning" & vbCrLf & "alt + O    Omskriv" & vbCrLf & "alt + N    Skift mellem auto, eksakt, num" & vbCrLf & "alt + T   Konverter ligning til/fra LaTex" & vbCrLf & vbCrLf & "Bemærk at der også er genveje til beregn og løs ligning(er) hvis du højreklikker på en ligning.", vbYesNo, Sprog.A(64))
'    If result = vbYes Then
    DeleteNormalDotm
    GenerateKeyboardShortcuts
        
    MsgBox Sprog.A(671), vbOKOnly, ""
        
'    End If

End Sub

Private Sub CommandButton_sletenheder_Click()
    TextBox_outunits.text = ""
End Sub





Private Sub OptionButton_casmaxima_Change()
    SetCasButtons
End Sub


Private Sub UserForm_Activate()
    On Error Resume Next
    Dim sett As String
    FillComboBoxCifre
    FillComboBoxLanguage
    FillComboBoxBackup
    FillComboBoxBackupNo
    FillComboBoxBackupTime
    SetCaptions
    
#If Mac Then
    OptionButton_graph.visible = False
    OptionButton_gnuplot.visible = False
    CommandButton_nulstilfigurer.visible = False
    CommandButton_nulstilfysik.visible = False
    CommandButton_nulstilkemiformler.visible = False
    CommandButton_nulstilmatformler.visible = False
    OptionButton_casgeogebradirect.visible = True
#Else
'    OptionButton_casgeogebradirect.visible = False
#End If

    ReadAllSettingsFromRegistry
    
    CheckBox_Insertforklaring.Value = MaximaForklaring
    CheckBox_Insertmaximacommand.Value = MaximaKommando
    ComboBox_cifre.Value = MaximaCifre
    CheckBox_complex.Value = MaximaComplex
    CheckBox_units.Value = MaximaUnits
    CheckBox_vidnotation.Value = MaximaVidNotation
    CheckBox_indlejret.Value = ExcelIndlejret
    TextBox_outunits.text = OutUnits
    CheckBox_autostart2.Value = AutoStart
    CheckBox_checkupdate.Value = SettCheckForUpdate
    CheckBox_bigfloat.Value = MaximaBigFloat
    CheckBox_showassum.Value = ShowAssum
    CheckBox_polaroutput.Value = PolarOutput
    CheckBox_dasdiffchr.Value = dAsDiffChr
    CheckBox_askref.Value = EqAskRef
    Label_antalb.Caption = AntalB
    
    ComboBox_language.ListIndex = LanguageSetting
    
    ComboBox_backup.ListIndex = BackupType
    ComboBox_backupno.text = BackupMaxNo
    ComboBox_backuptime.text = BackupTime

    If CheckBox_complex.Value Then
        CheckBox_polaroutput.visible = True
    Else
        CheckBox_polaroutput.visible = False
    End If
    
    If MaximaExact = 1 Then
        OptionButton_exactonly.Value = True
    ElseIf MaximaExact = 2 Then
        OptionButton_numonly.Value = True
    Else
        OptionButton_exactandnum.Value = True
    End If

    If Radians Then
        OptionButton_radianer.Value = True
    Else
        OptionButton_grader.Value = True
    End If

    If MaximaSeparator Then
        OptionButton_punktum.Value = True
    Else
        OptionButton_komma.Value = True
    End If
    
    If AllTrig Then
        OptionButton_trigall.Value = True
    Else
        OptionButton_trigone.Value = True
    End If

    If MaximaGangeTegn = VBA.ChrW(183) Then
        OptionButton_prik.Value = True
    ElseIf MaximaGangeTegn = VBA.ChrW(215) Then
        OptionButton_kryds.Value = True
    Else
        OptionButton_stjerne.Value = True
    End If

    If LmSet Then
        OptionButton_lmset.Value = True
    Else
        OptionButton_lmbool.Value = True
    End If
    
    
    If MaximaLogOutput = 0 Then
        OptionButton_logauto.Value = True
    ElseIf MaximaLogOutput = 2 Then
        OptionButton_log10.Value = True
    Else
        OptionButton_ln.Value = True
    End If
    
    If MaximaIndex Then
        OptionButton_indexvar.Value = True
    Else
        OptionButton_indextext.Value = True
    End If
    
    If GraphApp = 0 Then
        OptionButton_gnuplot.Value = True
    ElseIf GraphApp = 1 Then
        OptionButton_graph.Value = True
    ElseIf GraphApp = 2 Then
        OptionButton_geogebra.Value = True
    ElseIf GraphApp = 3 Then
        OptionButton_excel.Value = True
    ElseIf GraphApp = 4 Then
        OptionButton_geogebraweb.Value = True
    End If
    
    If EqNumPlacement Then
        OptionButton_placementleft.Value = True
    Else
        OptionButton_placementright.Value = True
    End If
    
    If EqNumType Then
        OptionButton_eqnumtwo.Value = True
    Else
        OptionButton_eqnumone.Value = True
    End If
    
    If CASengine = 0 Then
        OptionButton_casmaxima.Value = True
    ElseIf CASengine = 1 Then
        OptionButton_casgeogebra.Value = True
    ElseIf CASengine = 2 Then
        OptionButton_casgeogebradirect.Value = True
    Else
        OptionButton_casmaxima.Value = True
    End If
    
    MustRestart = False
    LoadUnits = False
    UserUnits = False
    LangChange = False
    
    SetCasButtons

End Sub

Sub SetCasButtons()
If OptionButton_casmaxima.Value Then
    FrameLog.visible = True
    CheckBox_units.visible = True
    CheckBox_bigfloat.visible = True
    CheckBox_showassum.visible = True
Else
    FrameLog.visible = False
    CheckBox_units.visible = False
    CheckBox_bigfloat.visible = False
    CheckBox_showassum.visible = False
End If
End Sub

Sub FillComboBoxCifre()
Dim i As Integer
    For i = 2 To 16
        ComboBox_cifre.AddItem i
    Next
End Sub

Sub FillComboBoxLanguage()
    ComboBox_language.Clear
    ComboBox_language.AddItem "Auto"
    ComboBox_language.AddItem "Dansk"
    ComboBox_language.AddItem "English"
    ComboBox_language.AddItem "Espaniol"
End Sub
Sub FillComboBoxBackupNo()
    ComboBox_backupno.Clear
    ComboBox_backupno.AddItem "1"
    ComboBox_backupno.AddItem "2"
    ComboBox_backupno.AddItem "5"
    ComboBox_backupno.AddItem "10"
    ComboBox_backupno.AddItem "20"
    ComboBox_backupno.AddItem "30"
    ComboBox_backupno.AddItem "40"
    ComboBox_backupno.AddItem "50"
    ComboBox_backupno.AddItem "60"
    ComboBox_backupno.AddItem "70"
    ComboBox_backupno.AddItem "80"
    ComboBox_backupno.AddItem "90"
    ComboBox_backupno.AddItem "100"
End Sub
Sub FillComboBoxBackupTime()
    ComboBox_backuptime.Clear
    ComboBox_backuptime.AddItem "1"
    ComboBox_backuptime.AddItem "2"
    ComboBox_backuptime.AddItem "5"
    ComboBox_backuptime.AddItem "10"
    ComboBox_backuptime.AddItem "20"
    ComboBox_backuptime.AddItem "30"
End Sub
Sub FillComboBoxBackup()
    ComboBox_backup.Clear
    ComboBox_backup.AddItem Sprog.A(175)
    ComboBox_backup.AddItem Sprog.A(176)
    ComboBox_backup.AddItem Sprog.A(177)
End Sub
    
Sub SetCaptions()
    Me.Caption = Sprog.RibSettings
    
    MultiPage1.Pages("Page1").Caption = Sprog.CASSettings
    MultiPage1.Pages("Page2").Caption = Sprog.Notation
    MultiPage1.Pages("Page3").Caption = Sprog.Advanced
    MultiPage1.Pages("Page4").Caption = Sprog.Units
    MultiPage1.Pages("Page5").Caption = Sprog.Graph
    MultiPage1.Pages("Page6").Caption = Sprog.A(7)
    
    CommandButton_annuller.Caption = Sprog.Cancel
    CommandButton_ok.Caption = Sprog.OK
    Label_sigfig.Caption = Sprog.SignificantFigures
    CheckBox_vidnotation.Caption = Sprog.ScientificNotation
    CheckBox_complex.Caption = Sprog.Complex
    CheckBox_units.Caption = Sprog.Units
    CheckBox_polaroutput.Caption = Sprog.PolarNotation
    Frame1.Caption = Sprog.AngleUnit
    OptionButton_grader.Caption = Sprog.Degrees
    OptionButton_radianer.Caption = Sprog.Radians
    CheckBox_showassum.Caption = Sprog.ShowSolutionConditions
    
    Frame5.Caption = Sprog.Exact & " ?"
    OptionButton_exactandnum.Caption = Sprog.Auto & "  (x=pi=3.14)"
    OptionButton_exactonly.Caption = Sprog.Exact & "  (x=pi)"
    OptionButton_numonly.Caption = Sprog.Numeric & "  (x=3.14)"
    CheckBox_bigfloat.Caption = Sprog.HighPrecision
    CheckBox_Insertforklaring.Caption = Sprog.InsertExplanation
    CheckBox_Insertmaximacommand.Caption = Sprog.InsertMaximaCommand
    Frame3.Caption = Sprog.Separators
    Label_list.Caption = Sprog.List
    Frame2.Caption = Sprog.Multsign
    OptionButton_prik.Caption = VBA.ChrW(183) & " " & Sprog.Dot
    Frame4.Caption = Sprog.SolutionType
    OptionButton_lmbool.Caption = Sprog.SolutionBoolean & ": x=1 v x=2"
    OptionButton_lmset.Caption = Sprog.SolutionSet & ": L={1,2}"
    Frame8.Caption = Sprog.IndexSubscript
    OptionButton_indextext.Caption = Sprog.JustText
    OptionButton_indexvar.Caption = Sprog.IsIndex
    Frame9.Caption = Sprog.StandardGraph
    CheckBox_indlejret.Caption = Sprog.ExcelEmbed
    Label_outputunits.Caption = Sprog.OutputUnits
    CommandButton_sletenheder.Caption = Sprog.Clear
    Label_unithelp.Caption = Sprog.UnitHelp
    Label_unitexamples.Caption = Sprog.UnitExamples
    FrameLog.Caption = Sprog.Logarithm & " output"
    FrameTrig.Caption = Sprog.TrigEquations
    OptionButton_trigall.Caption = Sprog.AllSolutions
    OptionButton_trigone.Caption = Sprog.OnlyOneSolution
    CheckBox_autostart2.Caption = Sprog.AutoStart
    CheckBox_checkupdate.Caption = Sprog.AutoUpdate
    CommandButton_shortcuts.Caption = Sprog.GenerateShortcuts
    CommandButton_restartmaxima.Caption = Sprog.RestartWordMat
    Label_language.Caption = Sprog.Language
    Label_calculationcount.Caption = Sprog.CalculationCount & ":"
    CommandButton_nulstilfigurer.Caption = Sprog.ResetFigurDoc
    CommandButton_nulstilmatformler.Caption = Sprog.ResetMathDoc
    CommandButton_nulstilfysik.Caption = Sprog.ResetPhysicsDoc
    CommandButton_nulstilkemiformler.Caption = Sprog.ResetChemistryDoc
    CheckBox_dasdiffchr.Caption = Sprog.dAsDiffChr
    CheckBox_dasdiffchr.ControlTipText = Sprog.dAsDiffChrHelp
    Frame_side.Caption = Sprog.A(15)
    OptionButton_placementleft.Caption = Sprog.A(16)
    OptionButton_placementright.Caption = Sprog.A(17)
    Frame_eqnumtype.Caption = Sprog.A(18)
    CheckBox_askref.Caption = Sprog.A(19)
    Label_Backupexplain.Caption = Sprog.A(173)
    Label_backuptime.Caption = Sprog.A(172)
    Label_savebackup.Caption = Sprog.A(170)
    Label_backupmaxno.Caption = Sprog.A(171)
    CommandButton_openbackup.Caption = Sprog.A(174)
    
End Sub

