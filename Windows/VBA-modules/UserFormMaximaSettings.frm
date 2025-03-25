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
' This form allows the user to set all the different settings in WordMath. Including language

Private MustRestart As Boolean
Private LoadUnits As Boolean
Private UserUnits As Boolean
Private LangChange As Boolean

Private EventsCol As New Collection
Sub SetEscEvents(ControlColl As Controls)
' SetEscEvents Me.Controls     in Initialize
    Dim CE As CEvents, c As control, TN As String, F As MSForms.Frame
    On Error Resume Next
    For Each c In ControlColl ' Me.Controls
        TN = TypeName(c)
        If TN = "CheckBox" Then
            Set CE = New CEvents: Set CE.CheckBoxControl = c: EventsCol.Add CE
        ElseIf TN = "OptionButton" Then
            Set CE = New CEvents: Set CE.OptionButtonControl = c: EventsCol.Add CE
        ElseIf TN = "ComboBox" Then
            Set CE = New CEvents: Set CE.ComboBoxControl = c: EventsCol.Add CE
        ElseIf TN = "Label" Then
            Set CE = New CEvents: Set CE.LabelControl = c: EventsCol.Add CE
        ElseIf TN = "TextBox" Then
            Set CE = New CEvents: Set CE.TextBoxControl = c: EventsCol.Add CE
        ElseIf TN = "CommandButton" Then
            Set CE = New CEvents: Set CE.CommandButtonControl = c: EventsCol.Add CE
        ElseIf TN = "ListBox" Then
            Set CE = New CEvents: Set CE.ListBoxControl = c: EventsCol.Add CE
        ElseIf TN = "Frame" Then
            Set F = c
            SetEscEvents F.Controls
        End If
    Next
End Sub

Private Sub CheckBox_complex_Change()
    If CheckBox_complex.Value Then
        CheckBox_polaroutput.visible = True
    Else
        CheckBox_polaroutput.visible = False
    End If
End Sub

Private Sub CheckBox_units_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CheckEsc KeyCode
End Sub

Private Sub ComboBox_language_Change()
    LangChange = True
End Sub

Private Sub Label_cancel_Click()
    Me.hide
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
    
    SettUseVBACAS = CheckBox_VBACAS.Value
    
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
    ExcelIndlejret = CheckBox_indlejret.Value
    AllTrig = OptionButton_trigall.Value
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
    MaximaDecOutType = ComboBox_DecType.ListIndex + 1
    
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
    
    UFMSettings.hide
    Sprog.CheckSetting
    Sprog.LoadSprogArray
    RibbonSubs.RefreshRibbon
End Sub

Private Sub CommandButton_openbackup_Click()
On Error Resume Next
#If Mac Then
    RunScript "OpenFinder", GetTempDir() & "/WordMat-backup"
#Else
    MaxProc.OpenFolder GetDocumentsDir & "\WordMat-Backup"
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
        MsgBox Sprog.A(741), vbOKOnly, Sprog.Error
    Else
        If Sprog.SprogNr = 1 Then
#If Mac Then
            MsgBox2 Sprog.A(738)
#Else
            MsgBox2 Sprog.A(739)
#End If
        Else
            MsgBox "Keyboard shortcuts restored", vbOKOnly, "Done"
        End If
    End If

End Sub

Private Sub CommandButton_sletenheder_Click()
    TextBox_outunits.Text = ""
End Sub

Private Sub Label_checkpartnerskab_Click()
    If QActivePartnership(True) Then
        MsgBox Sprog.A(120), vbOKOnly, "OK"
    Else
        MsgBox Sprog.A(121), vbOKOnly, "Sorry"
    End If
End Sub

Private Sub Label3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetTabsInactive
End Sub
Private Sub Label4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetTabsInactive
End Sub
Private Sub Label5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetTabsInactive
End Sub
Private Sub Label6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetTabsInactive
End Sub
Private Sub Label7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetTabsInactive
End Sub
Private Sub Label8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetTabsInactive
End Sub
Private Sub Label9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetTabsInactive
End Sub

Private Sub OptionButton_casmaxima_Change()
    SetCasButtons
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    
    RunFirst
    
    FillComboBoxCifre
    FillComboBoxDecType
    FillComboBoxLanguage
    FillComboBoxBackup
    FillComboBoxBackupNo
    FillComboBoxBackupTime
    SetCaptions
    Label_ok.BackColor = LBColorInactive
    Label_cancel.BackColor = LBColorInactive
    
    Label_geogebraexplain.Caption = Sprog.A(740)
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
    ComboBox_DecType.ListIndex = MaximaDecOutType - 1
    CheckBox_complex.Value = MaximaComplex
    CheckBox_units.Value = MaximaUnits
    CheckBox_indlejret.Value = ExcelIndlejret
    TextBox_outunits.Text = OutUnits
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
    
    If QActivePartnership Then
        CheckBox_VBACAS.Enabled = True
        CheckBox_VBACAS.Value = SettUseVBACAS
    Else
        CheckBox_VBACAS.Enabled = False
        CheckBox_VBACAS.Value = False
    End If

    If CASengineRegOnly = 0 Then
        OptionButton_casmaxima.Value = True
    ElseIf CASengineRegOnly = 1 Then
        OptionButton_casgeogebra.Value = True
    ElseIf CASengineRegOnly = 2 Then
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
Sub FillComboBoxDecType()
    ComboBox_DecType.Clear
    ComboBox_DecType.AddItem "Decimaler"
    ComboBox_DecType.AddItem Sprog.A(687)
    ComboBox_DecType.AddItem Sprog.A(669)
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
    Me.Caption = Sprog.A(443)
    
    MultiPage1.Pages("Page1").Caption = Sprog.A(414)
    MultiPage1.Pages("Page2").Caption = Sprog.A(666)
    MultiPage1.Pages("Page3").Caption = Sprog.A(668)
    MultiPage1.Pages("Page4").Caption = Sprog.A(693)
    MultiPage1.Pages("Page5").Caption = Sprog.A(667)
    MultiPage1.Pages("Page6").Caption = Sprog.A(7)
    
    Label_TAB1.Caption = "CAS"
    Label_TAB2.Caption = Sprog.A(666)
    Label_TAB3.Caption = Sprog.A(667)
    Label_TAB4.Caption = Sprog.A(693)
    Label_TAB5.Caption = Sprog.A(7)
    Label_TAB6.Caption = "Backup"
    Label_TAB7.Caption = Sprog.A(668)
    
    Label_cancel.Caption = Sprog.Cancel
    CommandButton_ok.Caption = Sprog.OK
    CheckBox_complex.Caption = Sprog.A(670)
    CheckBox_units.Caption = Sprog.A(693)
    CheckBox_polaroutput.Caption = Sprog.A(680)
    Frame1.Caption = Sprog.A(708)
    OptionButton_grader.Caption = Sprog.A(706)
    OptionButton_radianer.Caption = Sprog.A(707)
    CheckBox_showassum.Caption = Sprog.A(709)
    
    Frame5.Caption = Sprog.A(710) & " ?"
    OptionButton_exactandnum.Caption = Sprog.A(712) & "  (x=" & ChrW(960) & "=3.14)"
    OptionButton_exactonly.Caption = Sprog.A(710) & "  (x=" & ChrW(960) & ")"
    OptionButton_numonly.Caption = Sprog.A(711) & "  (x=3.14)"
    CheckBox_bigfloat.Caption = Sprog.A(713)
    CheckBox_Insertforklaring.Caption = Sprog.A(714)
    CheckBox_Insertmaximacommand.Caption = Sprog.A(715)
    Frame3.Caption = Sprog.A(716)
    Label_list.Caption = Sprog.A(717)
    Frame2.Caption = Sprog.A(718)
    OptionButton_prik.Caption = VBA.ChrW(183) & " (" & Sprog.A(719) & ")"
    Frame4.Caption = Sprog.A(720)
    OptionButton_lmbool.Caption = Sprog.A(721) & ": x=1 v x=2"
    OptionButton_lmset.Caption = Sprog.A(722) & ": L={1,2}"
    Frame8.Caption = Sprog.A(723)
    OptionButton_indextext.Caption = Sprog.A(724)
    OptionButton_indexvar.Caption = Sprog.A(725)
    Frame9.Caption = Sprog.A(726)
    CheckBox_indlejret.Caption = Sprog.A(727)
    Label_outputunits.Caption = Sprog.A(728)
    CommandButton_sletenheder.Caption = Sprog.A(815)
    Label_unithelp.Caption = Sprog.A(729)
    Label_unitexamples.Caption = Sprog.A(730)
    FrameLog.Caption = Sprog.A(816) & " output"
    FrameTrig.Caption = Sprog.A(731)
    OptionButton_trigall.Caption = Sprog.A(732)
    OptionButton_trigone.Caption = Sprog.A(733)
    CheckBox_checkupdate.Caption = Sprog.A(734)
    CommandButton_shortcuts.Caption = Sprog.A(735)
    CommandButton_restartmaxima.Caption = Sprog.A(736)
    Label_language.Caption = Sprog.A(817)
    Label_calculationcount.Caption = Sprog.A(737) & ":"
    CheckBox_dasdiffchr.Caption = Sprog.A(840)
    CheckBox_dasdiffchr.ControlTipText = Sprog.A(841)
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
    Frame_casengine.Height = 95
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
    SetEscEvents Me.Controls
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
Private Sub Label_cancel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_cancel.BackColor = LBColorPress
End Sub
Private Sub Label_cancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_cancel.BackColor = LBColorHover
End Sub
Private Sub Label_ok_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_ok.BackColor = LBColorPress
End Sub
Private Sub Label_ok_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_ok.BackColor = LBColorHover
End Sub
Private Sub Label_TAB1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_TAB1.BackColor = LBColorPress
End Sub
Private Sub Label_TAB1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetTabsInactive
    If MultiPage1.Value <> 0 Then Label_TAB1.BackColor = LBColorHover
End Sub
Private Sub Label_TAB2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_TAB2.BackColor = LBColorPress
End Sub
Private Sub Label_TAB2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetTabsInactive
    If MultiPage1.Value <> 1 Then Label_TAB2.BackColor = LBColorHover
End Sub
Private Sub Label_TAB3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_TAB3.BackColor = LBColorPress
End Sub
Private Sub Label_TAB3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetTabsInactive
    If MultiPage1.Value <> 2 Then Label_TAB3.BackColor = LBColorHover
End Sub
Private Sub Label_TAB4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_TAB4.BackColor = LBColorPress
End Sub
Private Sub Label_TAB4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetTabsInactive
    If MultiPage1.Value <> 3 Then Label_TAB4.BackColor = LBColorHover
End Sub
Private Sub Label_TAB5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_TAB5.BackColor = LBColorPress
End Sub
Private Sub Label_TAB5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetTabsInactive
    If MultiPage1.Value <> 4 Then Label_TAB5.BackColor = LBColorHover
End Sub
Private Sub Label_TAB6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_TAB6.BackColor = LBColorPress
End Sub
Private Sub Label_TAB6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetTabsInactive
    If MultiPage1.Value <> 5 Then Label_TAB6.BackColor = LBColorHover
End Sub
Private Sub Label_TAB7_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_TAB7.BackColor = LBColorPress
End Sub
Private Sub Label_TAB7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetTabsInactive
    If MultiPage1.Value <> 6 Then Label_TAB7.BackColor = LBColorHover
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    MsgBox KeyCode
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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

Sub CheckEsc(ByVal KeyCode As Integer)
    If KeyCode = 27 Then
        Label_cancel_Click
    End If
End Sub

