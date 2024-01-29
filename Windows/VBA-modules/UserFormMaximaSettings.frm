VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormMaximaSettings 
   Caption         =   "Indstillinger"
   ClientHeight    =   5955
   ClientLeft      =   -15
   ClientTop       =   45
   ClientWidth     =   10410
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

Private Sub Label_cancel_Click()
    Me.Hide
End Sub

Private Sub NulstilWordDoc(FilNavn As String)
Dim appdir As String
Dim fs As Object
On Error GoTo Fejl
    
If MsgBox(Sprog.A(665) & ", " & FilNavn & " ," & Sprog.A(666), vbYesNo, Sprog.A(667)) = vbNo Then GoTo Slut
    
#If Mac Then
' cant
'    FileCopy "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/" & filnavn, "~/Library/Containers/com.microsoft.Word/Data/WordMat/" & Replace(filnavn, "\", "/")
#Else
Set fs = CreateObject("Scripting.FileSystemObject")
FilNavn = GetProgramFilesDir & "\WordMat\" & FilNavn
appdir = Environ("AppData")
If Dir(FilNavn) <> "" And appdir <> "" Then
  fs.CopyFile FilNavn, appdir & "\WordMat\"
  MsgBox Sprog.A(668) & " " & FilNavn & " " & Sprog.A(669), vbOKOnly, Sprog.A(670)
Else
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error '"Filen " & filnavn & " kunne ikke findes eller appdata mappen kunne ikke findes"
End If
#End If

GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:


End Sub

Private Sub CommandButton_annuller_Click()

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
    Dim UnitChanged As Boolean
    
    If InStr(TextBox_outunits.Text, "/") > 0 Or InStr(TextBox_outunits.Text, "*") > 0 Or InStr(TextBox_outunits.Text, "^") > 0 Then
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

    If OptionButton_regdll.Value Then
        DllConnType = 0
    ElseIf OptionButton_directdll.Value Then
        DllConnType = 1
    ElseIf OptionButton_WSH.Value Then
        DllConnType = 2
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
    BackupMaxNo = ComboBox_backupno.Text
    BackupTime = ComboBox_backuptime.Text
    
    If MaximaUnits <> CheckBox_units.Value Then
        MaximaUnits = CheckBox_units.Value
        UnitChanged = True
        If CheckBox_units.Value Then
            LoadUnits = True
        Else
            MustRestart = True
        End If
    End If
    
    If OutUnits <> TextBox_outunits.Text Then
        OutUnits = TextBox_outunits.Text
        UserUnits = True
    End If
        
    SetMathAutoCorrect
    
    UFMSettings.Hide
    Sprog.CheckSetting
    RibbonSubs.RefreshRibbon
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
    Dim WT As Template, TemplateFundet As Boolean, KSok As Boolean, KB As KeyBinding
    
    ' Slet genveje i normal.dotm ' Det kan give fejl, specielt på mac
    DeleteKeyboardShortcutsInNormalDotm
    
    ' Find den vedhæftede globale skabelon
    For Each WT In Application.Templates
        If LCase(Left(WT, 7)) = "wordmat" And LCase(right(WT, 5)) = ".dotm" Then
            CustomizationContext = WT
            TemplateFundet = True
            Exit For
        End If
    Next
        
    ' Check om genvejene er sat i den vedhæftede skabelon. Hvis ikke så sættes de i normal.dotm
#If Mac Then
    On Error Resume Next
    If TemplateFundet Then
        If KeyBindings.Count > 10 Then
            For Each KB In KeyBindings
                Err.Clear
                If KB.KeyString = "Option+B" Then
                    KSok = True
                    Exit For
                End If
                If Err.Number > 0 Then
                    KSok = False
                    Exit For
                End If
            Next
        End If
    End If
#Else
    If TemplateFundet Then
        If KeyBindings.Count > 10 Then
            For Each KB In KeyBindings
                If KB.Command = "WordMat.Maxima.beregn" Then
                    KSok = True
                    Exit For
                End If
            Next
        End If
    End If
#End If
    
    If Not KSok Then
        MsgBox "Det ser ud til at genvejene ikke er sat korrekt i denne udgave af WordMat. Det kræver nok en Fejlmeldingen", vbOKOnly, "Fejl"
    Else
        MsgBox "Keyboard shortcuts restored", vbOKOnly, "Done"
    End If

End Sub

Private Sub CommandButton_sletenheder_Click()
    TextBox_outunits.Text = ""
End Sub

Private Sub Label_checkpartnerskab_Click()
    If QActivePartnership(True) Then
        MsgBox "Din skole har aktivt partnerskab", vbOKOnly, "Partnerskab OK"
    Else
        MsgBox "Der kunne ikke registreres et aktivt partnerskab", vbOKOnly, "Ingen Partnerskab"
    End If
End Sub

Private Sub Label3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetTabsInactive
End Sub
Private Sub Label4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetTabsInactive
End Sub
Private Sub Label5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetTabsInactive
End Sub
Private Sub Label6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetTabsInactive
End Sub
Private Sub Label7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetTabsInactive
End Sub
Private Sub Label8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetTabsInactive
End Sub
Private Sub Label9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetTabsInactive
End Sub

Private Sub OptionButton_casmaxima_Change()
    SetCasButtons
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    FillComboBoxCifre
    FillComboBoxLanguage
    FillComboBoxBackup
    FillComboBoxBackupNo
    FillComboBoxBackupTime
    SetCaptions
    Label_ok.BackColor = LBColorInactive
    Label_cancel.BackColor = LBColorInactive
    
    If Sprog.SprogNr = 1 Then
        Label_geogebraexplain.Caption = "GeoGebra Calculator Suite bliver installeret sammen med WordMat. Det kører i en browser, men kræver ikke internet. Det fungerer ens på Windows og Mac, og starter hurtigt op." & vbCrLf & "GeoGebra 5 fungerer ens på Windows og Mac, men kræver separat installation." & vbCrLf & "Windows understøtter også de andre app-versioner af GeoGebra som Geogebra 6, Graphing calculator mm. Disse vil blive anvendt, hvis GeoGebra 5 ikke er installeret."
    Else
        Label_geogebraexplain.Caption = "GeoGebra Calculator Suite is installed with WordMat. It runs in a browser, but doesn't require internet. It works the same on Windows and Mac, and starts quickly." & vbCrLf & "GeoGebra 5 Works the same on Windows and Mac, but requires installation." & vbCrLf & "Windows also supports the other app-versions of GeoGebra like Geogebra 6, Graphing Calculator etc. These will be used if GeoGebra 5 isn't installed."
    End If
#If Mac Then
    OptionButton_graph.visible = False
    OptionButton_gnuplot.visible = False
    CommandButton_nulstilfigurer.visible = False
    CommandButton_nulstilfysik.visible = False
    CommandButton_nulstilkemiformler.visible = False
    CommandButton_nulstilmatformler.visible = False
    OptionButton_casgeogebradirect.visible = True
    Frame_conntype.visible = False
    OptionButton_geogebra.Caption = "GeoGebra 5"
#Else
    OptionButton_geogebra.Caption = "GeoGebra 5"
'    OptionButton_geogebra.Caption = "GeoGebra 5 (6, App-versions ...)"
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
    TextBox_outunits.Text = OutUnits
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
    ComboBox_backupno.Text = BackupMaxNo
    ComboBox_backuptime.Text = BackupTime

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
    
    If DllConnType = 0 Then
        OptionButton_regdll.Value = True
    ElseIf DllConnType = 1 Then
        OptionButton_directdll.Value = True
    Else
        OptionButton_WSH.Value = True
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
    
    Label_TAB1.Caption = "CAS"
    Label_TAB2.Caption = Sprog.Notation
    Label_TAB3.Caption = Sprog.Graph
    Label_TAB4.Caption = Sprog.Units
    Label_TAB5.Caption = Sprog.A(7)
    Label_TAB6.Caption = "Backup"
    Label_TAB7.Caption = Sprog.Advanced
    
    Label_cancel.Caption = Sprog.Cancel
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
    OptionButton_prik.Caption = VBA.ChrW(183) & " (" & Sprog.Dot & ")"
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
    Frame_casengine.Caption = Sprog.A(686)
#If Mac Then
    Label_casexplain.visible = True
#Else
    Label_casexplain.visible = False
    Frame_casengine.Height = 75
#End If
    
End Sub

Sub ScaleForm(SF As Double)
' SF er scalefactor. Ændrer størrelsen på en formen og justerer font og position af alle elementer på formen
Dim c As control
    For Each c In Me.Controls
        c.Left = c.Left * SF
        c.Top = c.Top * SF
        c.Width = c.Width * SF
        c.Height = c.Height * SF
        c.Font.Size = CInt(c.Font.Size * SF)
    Next
    Me.Width = Me.Width * SF
    Me.Height = Me.Height * SF
    Me.Font.Size = CInt(Me.Font.Size * SF)
End Sub

Private Sub UserForm_Initialize()
    MultiPage1.Value = 0
#If Mac Then
    ScaleForm 1.5
#End If
End Sub

Private Sub Label_ok_Click()
    CommandButton_ok_Click
End Sub

Private Sub Label_TAB1_Click()
    MultiPage1.Value = 0
    SetTabsInactive
    Label_TAB1.BackColor = LBColorTABPress
End Sub
Private Sub Label_TAB2_Click()
    MultiPage1.Value = 1
    SetTabsInactive
    Label_TAB2.BackColor = LBColorTABPress
End Sub
Private Sub Label_TAB3_Click()
    MultiPage1.Value = 2
    SetTabsInactive
    Label_TAB3.BackColor = LBColorTABPress
End Sub
Private Sub Label_TAB4_Click()
    MultiPage1.Value = 3
    SetTabsInactive
    Label_TAB4.BackColor = LBColorTABPress
End Sub
Private Sub Label_TAB5_Click()
    MultiPage1.Value = 4
    SetTabsInactive
    Label_TAB5.BackColor = LBColorTABPress
End Sub
Private Sub Label_TAB6_Click()
    MultiPage1.Value = 5
    SetTabsInactive
    Label_TAB6.BackColor = LBColorTABPress
End Sub
Private Sub Label_TAB7_Click()
    MultiPage1.Value = 6
    SetTabsInactive
    Label_TAB7.BackColor = LBColorTABPress
End Sub
Private Sub Label_cancel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorPress
End Sub
Private Sub Label_cancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorHover
End Sub
Private Sub Label_ok_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorPress
End Sub
Private Sub Label_ok_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorHover
End Sub
Private Sub Label_TAB1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_TAB1.BackColor = LBColorPress
End Sub
Private Sub Label_TAB1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetTabsInactive
    If MultiPage1.Value <> 0 Then Label_TAB1.BackColor = LBColorHover
End Sub
Private Sub Label_TAB2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_TAB2.BackColor = LBColorPress
End Sub
Private Sub Label_TAB2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetTabsInactive
    If MultiPage1.Value <> 1 Then Label_TAB2.BackColor = LBColorHover
End Sub
Private Sub Label_TAB3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_TAB3.BackColor = LBColorPress
End Sub
Private Sub Label_TAB3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetTabsInactive
    If MultiPage1.Value <> 2 Then Label_TAB3.BackColor = LBColorHover
End Sub
Private Sub Label_TAB4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_TAB4.BackColor = LBColorPress
End Sub
Private Sub Label_TAB4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetTabsInactive
    If MultiPage1.Value <> 3 Then Label_TAB4.BackColor = LBColorHover
End Sub
Private Sub Label_TAB5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_TAB5.BackColor = LBColorPress
End Sub
Private Sub Label_TAB5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetTabsInactive
    If MultiPage1.Value <> 4 Then Label_TAB5.BackColor = LBColorHover
End Sub
Private Sub Label_TAB6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_TAB6.BackColor = LBColorPress
End Sub
Private Sub Label_TAB6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetTabsInactive
    If MultiPage1.Value <> 5 Then Label_TAB6.BackColor = LBColorHover
End Sub
Private Sub Label_TAB7_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_TAB7.BackColor = LBColorPress
End Sub
Private Sub Label_TAB7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetTabsInactive
    If MultiPage1.Value <> 6 Then Label_TAB7.BackColor = LBColorHover
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorInactive
    Label_cancel.BackColor = LBColorInactive
    SetTabsInactive
End Sub

Sub SetTabsInactive()
    
    If MultiPage1.Value <> 0 Then Label_TAB1.BackColor = LBColorInactive
    If MultiPage1.Value <> 1 Then Label_TAB2.BackColor = LBColorInactive
    If MultiPage1.Value <> 2 Then Label_TAB3.BackColor = LBColorInactive
    If MultiPage1.Value <> 3 Then Label_TAB4.BackColor = LBColorInactive
    If MultiPage1.Value <> 4 Then Label_TAB5.BackColor = LBColorInactive
    If MultiPage1.Value <> 5 Then Label_TAB6.BackColor = LBColorInactive
    If MultiPage1.Value <> 6 Then Label_TAB7.BackColor = LBColorInactive
    
End Sub
