Attribute VB_Name = "Maxima"
Option Explicit
'Public UFWait As UserFormWaitForMaxima
Public UFSelectVar As UserFormSelectVar
'Public UFSelectVars As UserFormSelectVars
Public omax As CMaxima
Public tid As Double
Private DeVarList As String
Private TempCas As Integer

Public Function PrepareMaxima() As Boolean 'Optional Unit As Boolean = False
    '    Dim UFwait2 As UserFormWaitForMaxima

    On Error GoTo fejl
    Dim op As Boolean
    If DebugWM Then
        UserFormDebug.Label_time.Caption = ""
        tid = Timer
    End If
#If Mac Then
    Dim D As Document
#End If
    
    SaveBackup
    
'    Dim UfWait2 As UserFormWaitStartup
    op = False
    If Not SettingsRead Then ReadAllSettingsFromRegistry

    If omax Is Nothing Then
        '        LavRCMenu    ' højreklikmenu på ligninger
        SetMathAutoCorrect
        On Error Resume Next
        Application.Run macroname:="Popstart"
        On Error GoTo fejl
'        If UfWait2 Is Nothing Then Set UfWait2 = New UserFormWaitStartup
'        UfWait2.Show vbModeless
        op = True
#If Mac Then
        Set D = ActiveDocument
#End If
        DoEvents
        '        Wait (0.1)
        Set omax = New CMaxima
    End If

    If MaxProc Is Nothing And CASengine = 0 Then
        If Not op Then
#If Mac Then
            Set D = ActiveDocument
#Else
'            If UfWait2 Is Nothing Then Set UfWait2 = New UserFormWaitStartup
 '           UfWait2.Show vbModeless
            op = True
#End If
            DoEvents
        End If
        On Error Resume Next
        
#If Mac Then
#Else
getproc:
        Err.Clear
        Set MaxProc = GetMaxProc() 'CreateObject("MaximaProcessClass")
        If Err.Number <> 0 Then
            Err.Clear
            If QActivePartnership(False, True) Then
                If DllConnType = 0 Then
                    If MsgBox2("Kan ikke forbinde til Maxima. Vil du anvende metoden 'dll direct' i stedet?" & VbCrLfMac & VbCrLfMac & "(Denne indstilling findes under avanceret i Indstillinger)", vbYesNo, Sprog.Error) = vbYes Then
                        DllConnType = 1
                        GoTo getproc
                    End If
                ElseIf DllConnType = 1 Then
                    If MsgBox2("Kan ikke forbinde til Maxima. Vil du anvende metoden 'WSH' i stedet?" & VbCrLfMac & VbCrLfMac & "(Denne indstilling findes under avanceret i Indstillinger)", vbYesNo, Sprog.Error) = vbYes Then
                        DllConnType = 2
                    End If
                Else ' wsh har ikke brug for
                End If
            Else
                MsgBox2 Sprog.A(54), vbOKOnly, Sprog.Error
            End If
            GoTo slut
        End If
        If DllConnType = 2 Then
            If SettCheckForUpdate Then CheckForUpdateSilent
            GoTo finish
        End If
        On Error GoTo fejl
        MaxProc.Units = 0
        If CASengine = 0 Then MaxProc.StartMaximaProcess
#End If
        If SettCheckForUpdate Then CheckForUpdateSilent

#If Mac Then
        If MaximaUnits Then
            TurnUnitsOn
        End If
#Else
        WaitForMaximaUntil
        If MaxProc.ErrCode > 0 Then
            MsgBox Sprog.A(55) & vbCrLf & "ErrCode: " & MaxProc.ErrCode & vbCrLf & vbCrLf & MaxProc.LastMaximaOutput, vbOKOnly, Sprog.Error
            GoTo fejl
        End If
#End If
    End If
#If Mac Then
#Else
    If MaximaUnits Then
        If MaxProcUnit Is Nothing Then
            If Not op Then
'                If UfWait2 Is Nothing Then Set UfWait2 = New UserFormWaitStartup
'                UfWait2.Show vbModeless
                op = True
#If Mac Then
                Set D = ActiveDocument
#End If
                DoEvents
            End If
            On Error Resume Next
            Set MaxProcUnit = GetMaxProc() 'CreateObject("MaximaProcessClass")
            If Err.Number <> 0 Then
                MsgBox Sprog.A(54), vbOKOnly, Sprog.Error
                GoTo slut
            End If
            On Error GoTo fejl
            MaxProcUnit.Units = 1
            MaxProcUnit.OutUnits = omax.ConvertUnits(OutUnits)
            MaxProcUnit.StartMaximaProcess
            WaitForMaximaUnitUntil
        End If
    End If
#End If

finish:
    omax.ConvertLnLog = True ' andre funktioner kan ændre denne. den nulstilles
    If Not omax.PrepareNewCommand Then    ' nulstiller og finder definitioner
        GoTo fejl
    End If
    On Error Resume Next
    If op Then
'        Unload UfWait2
#If Mac Then
        D.Activate
#End If
    End If
    PrepareMaxima = True
    GoTo slut
fejl:
    On Error Resume Next
'    If Not UfWait2 Is Nothing Then Unload UfWait2
    PrepareMaxima = False
slut:
    On Error Resume Next
'    If Not UfWait2 Is Nothing Then Unload UfWait2
End Function
#If Mac Then
Function GetMaxProc() As MaximaProcess
        Set GetMaxProc = New MaximaProcess
End Function
#Else
Function GetMaxProc() As Object
    If DllConnType = 0 Then
        Set GetMaxProc = CreateObject("MaximaProcessClass")
    ElseIf DllConnType = 1 Then
        Set GetMaxProc = PGetMaxProc()
    End If
    If Not GetMaxProc Is Nothing Then GetMaxProc.SetMaximaPath GetMaximaPath()
End Function
#End If
Sub WaitForMaximaUntil(Optional StopTime As Integer = 500)
    Dim i As Integer

    If StopTime = 0 Then StopTime = 300

    Do While MaxProc.Finished = 0 And MaxProc.ErrCode = 0 And Not (omax.StopNow) And i < StopTime
        Wait (0.1)
        i = i + 1
    Loop
slut:
End Sub
Sub WaitForMaximaUnitUntil(Optional StopTime As Integer = 500)
    Dim i As Integer

    If StopTime = 0 Then StopTime = 300
#If Mac Then
    Do While MaxProc.Finished = 0 And MaxProc.ErrCode = 0 And Not (omax.StopNow) And i < StopTime
#Else
    Do While MaxProcUnit.Finished = 0 And MaxProcUnit.ErrCode = 0 And Not (omax.StopNow) And i < StopTime
#End If
        Wait (0.1)
        i = i + 1
    Loop
slut:
End Sub

Public Sub PrepareMaximaNoSplash()
    On Error GoTo slut
    If Not SettingsRead Then ReadAllSettingsFromRegistry
    If omax Is Nothing Then
'        LavRCMenu    ' højreklikmenu på ligninger
        Set omax = New CMaxima
        If MaxProc Is Nothing Then
            '        Set MaxProc = New MathMenu.MaximaProcessClass
            Set MaxProc = GetMaxProc() 'CreateObject("MaximaProcessClass")
                On Error Resume Next
                If Err.Number <> 0 Then
                    MsgBox Sprog.A(54), vbOKOnly, Sprog.Error
                    GoTo slut
                End If
                On Error GoTo fejl
            MaxProc.Units = 0
            MaxProc.StartMaximaProcess
            If SettCheckForUpdate Then CheckForUpdateSilent
            '            MsgBox "ok", vbOKOnly, ""
        End If
        
#If Mac Then
#Else
        If MaximaUnits Then
            If MaxProcUnit Is Nothing Then
                Set MaxProcUnit = GetMaxProc() 'CreateObject("MaximaProcessClass")
                On Error Resume Next
                If Err.Number <> 0 Then
                    MsgBox Sprog.A(54), vbOKOnly, Sprog.Error
                    GoTo slut
                End If
                On Error GoTo fejl
                MaxProcUnit.Units = 1
                MaxProcUnit.OutUnits = omax.ConvertUnits(OutUnits)
                MaxProcUnit.StartMaximaProcess
            End If
        End If
#End If
        
    End If
    omax.PrepareNewCommand    ' nulstiller og finder definitioner
    '    WaitForMaximaUntil
fejl:
slut:
End Sub
Sub RestartMaxima()

    Set omax = New CMaxima
    
#If Mac Then
#Else
    If Not CASengine = 0 Then Exit Sub
    
    
    If Not MaxProc Is Nothing Then
        MaxProc.CloseProcess
        MaxProc.StartMaximaProcess
    Else
        '        Set MaxProc = New MathMenu.MaximaProcessClass
        On Error Resume Next
'        Shell "cmd.exe /c taskkill /IM sbcl.exe /F" ' Denne slår alt maxima ihjel
        Set MaxProc = GetMaxProc() 'CreateObject("MaximaProcessClass")
        If Err.Number <> 0 Then
            MsgBox Sprog.A(54), vbOKOnly, Sprog.Error
            GoTo slut
        End If
        On Error GoTo fejl
        If omax Is Nothing Then
            Set omax = New CMaxima
        End If
    End If
    If Not MaxProcUnit Is Nothing Then
        MaxProcUnit.CloseProcess
        MaxProcUnit.StartMaximaProcess
    End If
#End If

    GoTo slut
fejl:
     MsgBox Sprog.A(55) & vbCrLf & "ErrCode: " & MaxProc.ErrCode & vbCrLf & vbCrLf & MaxProc.LastMaximaOutput, vbOKOnly, Sprog.Error
slut:
End Sub
Sub MaximaCommand()
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long, DontGoBack As Boolean
    On Error GoTo fejl
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled

    '    Set UFWait = New UserFormWaitForMaxima
    PrepareMaxima
    omax.prevspr = ""
    If Not omax.MaximaInstalled Then GoTo slut
    
    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    ElseIf Selection.OMaths.Count > 0 Then
        omax.ReadSelection
    Else
        omax.Kommando = Selection.Range.text
        DontGoBack = True
    End If
    
    
    
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    '    UFWait.Show vbModeless
    If CASengine = 0 Then
        omax.ExecuteMaximaCommand
    ElseIf CASengine = 1 Then
        OpenGeoGebraWeb omax.Kommando, "CAS", True, True
        GoTo slut
    ElseIf CASengine = 2 Then
        omax.MaximaOutput = RunGeoGebraDirect(omax.Kommando)
    End If


    If CheckForError Then GoTo slut

    omax.GoToEndOfSelectedMaths DontGoBack
    Selection.TypeParagraph
    omax.InsertMaximaOutput
    '   UFWait.Hide
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos

End Sub
Sub MaximaSolveInequality(Optional variabel As String)
' løser een ulighed
    On Error GoTo fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim eqs As Boolean
    Dim scrollpos As Double
    Dim sstart As Long
    Dim sslut As Long
    Dim s As String
    Dim ea As New ExpressionAnalyser
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo slut
    '    Set UFWait = New UserFormWaitForMaxima
    
    Set UFSelectVar = New UserFormSelectVar

    If Selection.OMaths.Count < 2 Then

        omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If
        eqs = False
        If InStr(omax.Kommando, VBA.ChrW(8804)) > 0 Or InStr(omax.Kommando, VBA.ChrW(8805)) > 0 Then
            omax.Kommando = Replace(omax.Kommando, VBA.ChrW(8804), "<")
            omax.Kommando = Replace(omax.Kommando, VBA.ChrW(8805), ">")
            eqs = True
        End If
        omax.FindVariable
        If variabel = vbNullString Then
            UFSelectVar.vars = omax.vars
            UFSelectVar.DefS = omax.DefString
            UFSelectVar.Show
            variabel = UFSelectVar.SelectedVar
        End If
        
        If variabel = "" Then
            GoTo slut
        End If
        If variabel = "" Then GoTo slut
        
        If CASengine = 1 Or CASengine = 2 Then
            s = Trim(omax.Kommando)
            s = Replace(s, vbCrLf, "")
            s = Replace(s, vbCr, "")
            s = Replace(s, vbLf, "")
            If MaximaComplex Then
                s = "Csolve(" & Replace(s, ",", ".") & "," & variabel & ")"
            Else
                s = "solve(" & Replace(s, ",", ".") & "," & variabel & ")"
            End If
            If MaximaDecOutType = 3 Then
                s = "ScientificText(" & s & " , " & MaximaCifre & ")"
            ElseIf MaximaExact = 2 Then
                s = "numeric(" & s & " , " & MaximaCifre & ")"
            End If

        End If
        
        If CASengine = 0 Then
            omax.MaximaSolveInequality (variabel)
        ElseIf CASengine = 1 Then
            If MaximaForklaring Then
                omax.GoToEndOfSelectedMaths
                If MaximaForklaring Then
                    Selection.TypeParagraph
                    InsertForklaring Sprog.EquationSolvedFor & variabel & Sprog.ByCAS, True
                    Selection.TypeParagraph
                End If
            End If
            OpenGeoGebraWeb s, "CAS", True, True
            GoTo slut
        ElseIf CASengine = 2 Then
            Call RunGeoGebraDirect(s)
            If omax.MaximaOutput = "{}" Then
                omax.MaximaOutput = variabel & VBA.ChrW(8712) & VBA.ChrW(8709)
            ElseIf omax.MaximaOutput = "{" & variabel & "=" & variabel & "}" Or omax.MaximaOutput = "{x=x}" Then
                omax.MaximaOutput = variabel & VBA.ChrW(8712) & VBA.ChrW(8477)
            Else
                If Left(omax.MaximaOutput, 1) = "{" Then omax.MaximaOutput = Mid(omax.MaximaOutput, 2, Len(omax.MaximaOutput) - 2)
                ea.text = omax.MaximaOutput
                omax.MaximaOutput = ""
                ea.Pos = 1
                Do
                    s = ea.GetNextListItem(ea.Pos, ";")
                    If s <> "" Then omax.MaximaOutput = omax.MaximaOutput & s & "    " & VBA.ChrW(8744) & "    "
                Loop While s <> ""
                If omax.MaximaOutput <> "" Then omax.MaximaOutput = Left(omax.MaximaOutput, Len(omax.MaximaOutput) - 9)
            End If
            Application.Activate
        End If
        
        
        If CheckForError Then GoTo slut
        
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord

        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph


        If eqs Then
            omax.MaximaOutput = Replace(omax.MaximaOutput, "<", VBA.ChrW(8804))
            omax.MaximaOutput = Replace(omax.MaximaOutput, ">", VBA.ChrW(8805))
        End If

        If MaximaForklaring And omax.MaximaOutput <> "" Then
            InsertForklaring vbTab & Sprog.A(56) & variabel & Sprog.A(57)
        End If


        If omax.MaximaOutput = "universalset" Then
            omax.MaximaOutput = "x" & VBA.ChrW(8712) & VBA.ChrW(8477)
            omax.InsertMaximaOutput
            Selection.TypeParagraph
            Selection.TypeText Sprog.A(58)
        ElseIf omax.MaximaOutput = "emptyset" Then
            omax.MaximaOutput = "L=" & VBA.ChrW(8709)
            omax.InsertMaximaOutput
            Selection.TypeParagraph
            If MaximaComplex Then
                Selection.TypeText Sprog.A(59) & variabel
            Else
                Selection.TypeText Sprog.A(60) & variabel
            End If
        Else
            omax.InsertMaximaOutput
        End If

    Else    ' ligningssystem
        MsgBox (Sprog.A(376))
    End If
    
    Oundo.EndCustomRecord

    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos

End Sub
Sub MaximaSolve()
' Denne konstruktion er nødvendig for at kunne kalde MaximaSolve med parametre fra TestTabel
' Man kan ikke tilknytte genveje subs der tager parametre selvom de optional
    MaximaSolvePar
End Sub
'Sub MaximaSolve(Optional variabel As String)
Sub MaximaSolvePar(Optional variabel As String)
    Dim Arr As Variant, s As String, t As String, V As String
    Dim fejlm As String
    On Error GoTo fejl
    Application.ScreenUpdating = False
    Dim IsSolved As Boolean
    Dim scrollpos As Double
    Dim UFSolvenumeric As New UserFormNumericQuestion
    Dim ea As New ExpressionAnalyser, SaveKommando As String
    Dim sstart As Long, sslut As Long, p As Long, p2 As Long
    Dim SaveSettingsCifre As Integer
    Dim SaveSettingsExact As Integer
    Dim SaveSettingsOutunits As String
    Dim SaveSettingsLog As Integer
    Dim SaveSettingsDecOutType As Integer
    
    scrollpos = ActiveWindow.VerticalPercentScrolled
    
    TempCas = CASengine
'    On Error Resume Next: oData.GetFromClipboard:   ClipText = oData.GetText: On Error GoTo fejl ' Skulle sikre at clipboard ikke ændres, men virker ikke
    PrepareMaxima
    omax.prevspr = ""

    Set UFSelectVar = New UserFormSelectVar
    UFSelectVar.NoEq = 1
    
    SaveSettingsCifre = MaximaCifre
    SaveSettingsExact = MaximaExact
    SaveSettingsOutunits = OutUnits
    SaveSettingsLog = MaximaLogOutput
    SaveSettingsDecOutType = MaximaDecOutType
    
    
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo slut
    sstart = Selection.start
    sslut = Selection.End

    If omax.DefFejl Then
        MsgBox2 Sprog.DefError & vbCrLf & VisDef & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox2 Sprog.EquationMissingError, vbOKOnly, Sprog.EquationMissingError2
        GoTo slut
    End If
    If sstart = sslut Then
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    If InStr(Selection.OMaths(1).Range.text, "<") > 1 Or InStr(Selection.OMaths(1).Range.text, ">") > 1 Or InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8804)) > 1 Or InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8805)) > 1 Then
        MaximaSolveInequality variabel
        GoTo slut
    End If
    p = InStr(Selection.OMaths(1).Range.text, "=")
    If p < 1 Then
        MsgBox Sprog.A(141), vbOKOnly, Sprog.Error
        GoTo slut
    Else
        p2 = InStr(p + 1, Selection.OMaths(1).Range.text, "=")
        If p2 > 0 Then
            If Sprog.SprogNr = 1 Then
                MsgBox "Der kan ikke være to ligmedtegn i en ligning", vbOKOnly, "Fejl"
            Else
                MsgBox "You cant place two equal sign in an equation", vbOKOnly, "Fejl"
            End If
            GoTo slut
        End If
    End If

    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    If Selection.OMaths.Count < 2 And InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8743)) < 1 Then
        ' kun 1 ligning

        UFSolvenumeric.Ligning = omax.Kommando
        
        omax.FindVariable
        If Not ValiderVariable Then GoTo slut
        SaveKommando = omax.Kommando
newcas:
        omax.StopNow = False
        omax.Kommando = SaveKommando
        If CASengine > 0 And Not AllTrig Then ' På geogebra skal der via vba genkendes om det er trigonometrisk ligning
            If Not InStr(omax.vars, ";") > 0 Then ' metoden virker kun med 1 variabel
                ea.SetNormalBrackets
                ea.text = omax.Kommando
                ea.text = Replace(ea.text, VBA.ChrW(8289), "")
                s = ""
                Do
                    V = ea.GetNextVar()
                    If V = "sin" Or V = "cos" Or V = "tan" Then
                        t = ea.GetNextBracketContent()
                        If InStr(t, omax.vars) > 0 Then
                            If V = "cos" Then
                                s = "180"
                            Else
                                s = "90"
                            End If
                        End If
                    End If
                    ea.Pos = ea.Pos + 1
                Loop While V <> ""
                If s <> "" And Radians Then
                    If s = "90" Then
                        s = ChrW(960) & "/2" ' pi
                    Else
                        s = ChrW(960) '"pi"
                    End If
                End If
                If s <> "" Then
                    UFSelectVar.TextBox_def.text = "0<=" & omax.vars & "<=" & s & VbCrLfMac
                    UFSelectVar.TempDefs = "0<=" & omax.vars & "<=" & s
                End If
            End If
        End If
        If variabel = vbNullString Then
            UFSelectVar.vars = omax.vars
            UFSelectVar.DefS = omax.DefString
            UFSelectVar.Show
            variabel = UFSelectVar.SelectedVar
        End If
        If variabel = "" Then GoTo slut
        omax.TempDefs = Replace(UFSelectVar.TempDefs, "%pi", "pi")
        If CASengine = 1 Or CASengine = 2 Then
            s = Trim(omax.Kommando)
            s = Replace(s, vbCrLf, "")
            s = Replace(s, vbCr, "")
            s = Replace(s, vbLf, "")
            If MaximaComplex Then
                s = "Csolve(" & Replace(s, ",", ".") & "," & variabel & ")"
            Else
                s = "solve(" & Replace(s, ",", ".") & "," & variabel & ")"
            End If
            If MaximaDecOutType = 3 Then
                s = "ScientificText(" & s & " , " & MaximaCifre & ")"
            ElseIf MaximaExact = 2 Then
                s = "numeric(" & s & " , " & MaximaCifre & ")"
            End If

        End If
        
        If CASengine = 0 Then
            omax.MaximaSolve (variabel)
        ElseIf CASengine = 1 Then
            If MaximaForklaring Then
                omax.GoToEndOfSelectedMaths
                If MaximaForklaring Then
                    Selection.TypeParagraph
                    InsertForklaring Sprog.EquationSolvedFor & variabel & Sprog.ByCAS, True
                    Selection.TypeParagraph
                End If
            End If
            OpenGeoGebraWeb s, "CAS", True, True
            GoTo slut
        ElseIf CASengine = 2 Then
            Call RunGeoGebraDirect(s)
            If omax.MaximaOutput = "{}" Then
                omax.MaximaOutput = variabel & VBA.ChrW(8712) & VBA.ChrW(8709)
            ElseIf omax.MaximaOutput = "{" & variabel & "=" & variabel & "}" Or omax.MaximaOutput = "{x=x}" Then
                omax.MaximaOutput = variabel & VBA.ChrW(8712) & VBA.ChrW(8477)
            Else
                If Left(omax.MaximaOutput, 1) = "{" Then omax.MaximaOutput = Mid(omax.MaximaOutput, 2, Len(omax.MaximaOutput) - 2)
                ea.text = omax.MaximaOutput
                omax.MaximaOutput = ""
                ea.Pos = 1
                Do
                    s = Trim(ea.GetNextListItem(ea.Pos, ";"))
                    If s <> "" Then omax.MaximaOutput = omax.MaximaOutput & s & "    " & VBA.ChrW(8744) & "    "
                Loop While s <> ""
                If omax.MaximaOutput <> "" Then omax.MaximaOutput = Left(omax.MaximaOutput, Len(omax.MaximaOutput) - 9)
'                omax.MaximaOutput = Replace(omax.MaximaOutput, ";", "    " & VBA.ChrW(8744) & "    ")    ' komma erstattes med eller
            End If
            Application.Activate
        End If
                
        
        If omax.StopNow Then
            If omax.AntalVars > 1 Then
                GoTo slut
            End If
        End If
        If CheckForError Then GoTo slut

        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
        
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph
        If Len(omax.MaximaOutput) > 250 Then
            Dim resultat As VbMsgBoxResult
            resultat = MsgBox(Sprog.A(127) & vbCrLf & vbCrLf & omax.MaximaOutput, vbOKCancel, Sprog.Warning)
            If resultat = vbCancel Then GoTo slut
        End If

        variabel = omax.ConvertToWordSymbols(variabel)
        
        Dim Sep As String
'        If CASengine = 0 Then
            Sep = VBA.ChrW(8744)
'        Else
'            sep = ";"
'        End If
        If omax.StopNow Or (omax.IsAllSolved(omax.MaximaOutput, variabel, Sep) = "false" And Not (InStr(variabel, "+") > 0)) Then
            IsSolved = False
        Else
            IsSolved = True
        End If

        ' indsæt forklaring hvis ønsket
        If MaximaForklaring And (IsSolved Or InStr(omax.KommentarOutput, "solving system of equations")) Then
            '            InsertForklaring "Ligningen løses for " & variabel & " vha. CAS-værktøjet " & AppNavn & ". "
            InsertForklaring Sprog.EquationSolvedFor & variabel & Sprog.ByCAS
        End If

        fejlm = TranslateComment(omax.KommentarOutput)

        If Len(fejlm) > 0 Then
            fejlm = Sprog.A(128) & vbCrLf & fejlm
            MsgBox fejlm & vbCrLf & vbCrLf & omax.KommentarOutput & vbCrLf & MaxProc.LastMaximaOutput, vbOKOnly, Sprog.Error
            RestartMaxima
        ElseIf InStr(omax.MaximaOutput, VBA.ChrW(8709)) Then    ' ingen løsning
            omax.InsertMaximaOutput
            Selection.TypeParagraph
            If MaximaComplex Then
                Selection.TypeText Sprog.A(129) & variabel
            Else
                Selection.TypeText Sprog.A(130) & variabel
            End If
        ElseIf InStr(omax.MaximaOutput, VBA.ChrW(8477)) > 0 Then  ' alle er løsning
            omax.InsertMaximaOutput
            Selection.TypeParagraph
            Selection.TypeText Sprog.A(132)
        ElseIf omax.MaximaOutput = "?" Or omax.MaximaOutput = "" Or InStr(omax.KommentarOutput, "Lisp error") > 0 Or (Not LmSet And Not IsSolved) Then
            UserFormChooseCAS.Show
            If UserFormChooseCAS.ChosenCAS = 2 Then ' maxima num
                CASengine = 0
                GoTo stophop
            ElseIf UserFormChooseCAS.ChosenCAS = -1 Then
                GoTo slut
            ElseIf UserFormChooseCAS.ChosenCAS = 1 Then ' Maxima sym
                CASengine = 0
                GoTo newcas
            ElseIf UserFormChooseCAS.ChosenCAS = 3 Then ' geogebra sym
                CASengine = 2
                GoTo newcas
            ElseIf UserFormChooseCAS.ChosenCAS = 4 Then ' geogebra num
                CASengine = 2
                Selection.MoveLeft wdCharacter, 1
                MaximaNsolve variabel
                GoTo slut
            ElseIf UserFormChooseCAS.ChosenCAS = 5 Then ' geogebra browser sym
                CASengine = 1
                GoTo newcas
            Else ' grafisk geogebra
                CASengine = 1
                Selection.MoveLeft wdCharacter, 1
                MaximaNsolve variabel
                GoTo slut
            End If
        ElseIf False Then
stophop:     If omax.AntalVars > 1 Then
                UFSolvenumeric.result = ""
            Else
                UFSolvenumeric.Label_omskrevet.Caption = omax.MaximaOutput
                omax.PrepareNewCommand
                omax.TempDefs = UFSelectVar.TempDefs
                UFSolvenumeric.FejlMeld = omax.KommentarOutput
                UFSolvenumeric.variabel = variabel
                UFSolvenumeric.Show
            End If
            If UFSolvenumeric.result = "num" Then ' grafisk løsning valgt
                Selection.End = sslut    ' slut skal være først eller går det galt
                Selection.start = sstart
                ActiveWindow.VerticalPercentScrolled = scrollpos
                If GraphApp = 0 And CASengine = 0 Then
                    MaximaSolveNumeric UFSelectVar.ListBox_vars.text
                Else
                    CASengine = 1
                    Selection.MoveLeft wdCharacter, 1
                    MaximaNsolve variabel
                    GoTo slut
                End If
            Else
                If UFSolvenumeric.result = "nsolve" Then
                    InsertForklaring Sprog.EquationSolvedNumFor & variabel & Sprog.ByCAS, False
                    omax.MaximaOutput = UFSolvenumeric.Label_nsolve.Caption
                ElseIf UFSolvenumeric.result = "omskriv" Then
                    InsertForklaring "", True
                    omax.MaximaOutput = UFSolvenumeric.Label_omskrevet.Caption
                End If
                omax.InsertMaximaOutput
                Selection.TypeParagraph
                Selection.TypeText UFSolvenumeric.Kommentar
            End If

        Else    ' hvis der er løsning
'            If MaximaUnits Then
'                omax.Kommando = omax.MaximaOutput
'                omax.beregn
'                omax.MaximaOutput = omax.MaximaOutput
'            End If
            omax.InsertMaximaOutput
        End If

        Oundo.EndCustomRecord

    Else    '--------------- ligningssystem ----------------------

        omax.FindVariable
        If Not ValiderVariable Then GoTo slut
        UFSelectVar.NoEq = omax.AntalKom
        UFSelectVar.vars = omax.vars
        UFSelectVar.DefS = omax.DefString
        UFSelectVar.Show
        variabel = UFSelectVar.SelectedVar

        '    variabel = InputBox("Indtast variable som ligningssystemet skal løses for adskilt af komma", "Variable", "x,y")

        If variabel = "" Then GoTo slut
        omax.TempDefs = UFSelectVar.TempDefs
        '    UFWait.ActionToPerform = "solvesystem"
        '    UFWait.VarParam = variabel

        '    Set UFWait.omax = omax
        '    UFWait.Show
        
newcassys:
        If CASengine = 1 Or CASengine = 2 Then
            If MaximaComplex Then
                s = "Csolve({" & Replace(Replace(omax.KommandoerStreng, ",", "."), ";", " , ") & "},{" & Replace(variabel, ";", " , ") & "})"
            Else
                s = "solve({" & Replace(Replace(omax.KommandoerStreng, ",", "."), ";", " , ") & "},{" & Replace(variabel, ";", " , ") & "})"
            End If
            If MaximaDecOutType = 3 Then
                s = "ScientificText(" & s & " , " & MaximaCifre & ")"
            ElseIf MaximaExact = 2 Then
                s = "numeric(" & s & " , " & MaximaCifre & ")"
            End If
        End If
        
        If CASengine = 0 Then
            omax.SolveSystem (variabel)
        ElseIf CASengine = 1 Then
            If MaximaForklaring Then
                omax.GoToEndOfSelectedMaths
                If MaximaForklaring Then
                    Selection.TypeParagraph
                    InsertForklaring Sprog.A(134) & variabel & Sprog.A(135), False
                    Selection.TypeParagraph
                End If
            End If
            OpenGeoGebraWeb s, "CAS", True, True
            GoTo slut
        ElseIf CASengine = 2 Then
            Call RunGeoGebraDirect(s)
            If omax.MaximaOutput = "{}" Then
                omax.MaximaOutput = variabel & VBA.ChrW(8712) & VBA.ChrW(8709)
            ElseIf omax.MaximaOutput = "{" & variabel & "=" & variabel & "}" Or omax.MaximaOutput = "{x=x}" Then
                omax.MaximaOutput = variabel & VBA.ChrW(8712) & VBA.ChrW(8477)
            ElseIf omax.MaximaOutput <> "?" And omax.MaximaOutput <> "" Then
                omax.ConvertOutputToSolvedGGB ' burde måske forbedres med ea.getnextlistitem ligesom solve
            End If
            Application.Activate
        End If
        
        If omax.StopNow Then GoTo slut
        If CheckForError Then GoTo slut
        If omax.StopNow Then GoTo slut
        Application.ScreenUpdating = False
        
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
        
        '    omax.KommentarOutput = TranslateReplaceComment(omax.KommentarOutput)

        If omax.MaximaOutput = "?merror(""Anumberwasfoundwhereavariablewasexpected-`solve'"")" Then
            MsgBox Sprog.A(133), vbOKOnly, Sprog.Error
            GoTo slut
        End If

        If (omax.DefFejl = True) Then
            MsgBox Sprog.DefError & vbCrLf & VisDef & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
            GoTo slut
        End If

        variabel = Replace(omax.ConvertToWordSymbols(variabel), ";", ",")
        
        If omax.MaximaOutput = "?" Or omax.MaximaOutput = "" Or InStr(omax.KommentarOutput, "Lisp error") > 0 Then
            UserFormChooseCAS.Show
            If UserFormChooseCAS.ChosenCAS = 2 Then ' maxima num
                GoTo stophop
            ElseIf UserFormChooseCAS.ChosenCAS = -1 Then
                GoTo slut
            ElseIf UserFormChooseCAS.ChosenCAS = 1 Then ' Maxima sym
                CASengine = 0
                GoTo newcassys
            ElseIf UserFormChooseCAS.ChosenCAS = 3 Then ' geogebra sym
                CASengine = 2
                GoTo newcassys
            ElseIf UserFormChooseCAS.ChosenCAS = 4 Then ' geogebra num
                TempCas = CASengine
                CASengine = 2
'                Selection.MoveLeft wdCharacter, 1
                MaximaNsolve variabel
                GoTo slut
                CASengine = TempCas
            ElseIf UserFormChooseCAS.ChosenCAS = 5 Then ' geogebra browser sym
                CASengine = 1
                GoTo newcassys
            Else ' grafisk geogebra
                TempCas = CASengine
                CASengine = 1
                Selection.MoveLeft wdCharacter, 1
                MaximaNsolve variabel
                CASengine = TempCas
                GoTo slut
            End If
        ElseIf Len(omax.MaximaOutput) > 1 Then
            omax.GoToEndOfSelectedMaths
            Selection.TypeParagraph
            If MaximaForklaring Then
                InsertForklaring Sprog.A(134) & variabel & Sprog.A(135)
            End If
            omax.InsertMaximaOutput
            Arr = Split(omax.MaximaOutput, "=")
            If UBound(Arr) = 1 Then
                If InStr(Arr(0), variabel) > 0 And InStr(Arr(1), variabel) > 0 Then
                    '                    Result = MsgBox("Maxima kunne ikke løse ligningssystemet. Den var for kompleks." & vbCrLf & vbCrLf & omax.KommentarOutput & vbCrLf & vbCrLf & "Tryk OK hvis du vil forsøge at løse ligningen numerisk.", vbOKCancel, "Fejl")
                    UFSolvenumeric.FejlMeld = omax.KommentarOutput
                    UFSolvenumeric.Show
                    If UFSolvenumeric.result = "num" Then
                        Selection.End = sslut    ' slut skal være først eller går det galt
                        Selection.start = sstart
                        ActiveWindow.VerticalPercentScrolled = scrollpos
                        MaximaSolveNumeric
                    Else
                        Selection.TypeParagraph
                        Selection.TypeText Sprog.A(136)
                    End If
                End If
            ElseIf InStr(omax.MaximaOutput, VBA.ChrW(8709)) > 0 And CASengine > 0 Then
                Selection.TypeParagraph
                Selection.TypeText "GeoGebra har været anvendt til at løse ligningssystemet. Det er usikkert om der kan være løsninger. Det anbefales at forsøge med anden metode. Fx Maxima, eller numerisk/grafisk"
            End If
        Else    ' ingen løsninger
            omax.GoToEndOfSelectedMaths
            Selection.TypeParagraph
            If Len(omax.KommentarOutput) <= 1 Then
                omax.MaximaOutput = "L=" & VBA.ChrW(8709)
                '               omax.GoToEndOfSelectedMaths
                omax.InsertMaximaOutput
                Selection.TypeParagraph
                If MaximaComplex Then
                    Selection.TypeText Sprog.A(137) & variabel
                Else
                    Selection.TypeText Sprog.A(138) & variabel
                End If
            Else
                If (omax.DefFejl = True) Then
                    MsgBox Sprog.DefError & vbCrLf & VisDef & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
                Else
                    fejlm = Sprog.A(131) & vbCrLf
                End If
                If InStr(omax.Kommando, "=") < 1 Then
                    fejlm = fejlm & Sprog.A(139) & vbCrLf
                End If
                '                fejlm = fejlm & TranslateComment(omax.KommentarOutput)
                UFSolvenumeric.FejlMeld = omax.KommentarOutput
                UFSolvenumeric.Show
                If UFSolvenumeric.result = "num" Then
                    Selection.End = sslut    ' slut skal være først eller går det galt
                    Selection.start = sstart
                    ActiveWindow.VerticalPercentScrolled = scrollpos
                    MaximaSolveNumeric
                Else
                    Selection.TypeParagraph
                    Selection.TypeText Sprog.A(140)
                End If
                '                MsgBox fejlm & "Angav du de rigtige variable?" & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, "Fejl"
            End If
        End If
        Oundo.EndCustomRecord
    End If
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    '    omax.Luk
    On Error Resume Next
    MaximaCifre = SaveSettingsCifre
    MaximaExact = SaveSettingsExact
    OutUnits = SaveSettingsOutunits
    MaximaLogOutput = SaveSettingsLog
    MaximaDecOutType = SaveSettingsDecOutType
    CASengine = TempCas
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
    '   UnLockWindow
End Sub
Sub InsertForklaring(ForklarTekst As String, Optional biimp As Boolean = True)
    Dim tdefs As String
    Dim gemfontsize As Integer
    Dim gemitalic As Boolean
    Dim gemfontcolor As Integer
    Dim gemsb As Integer
    Dim gemsa As Integer
    Dim mo As Range
#If Mac Then
    Selection.TypeText " " ' ensures that the preceding math-box does not change font
#End If
    gemfontsize = Selection.Font.Size
    gemitalic = Selection.Font.Italic
    gemfontcolor = Selection.Font.ColorIndex
    gemsb = Selection.ParagraphFormat.SpaceBefore
    gemsa = Selection.ParagraphFormat.SpaceAfter

    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 2
        .SpaceAfterAuto = False
        '        .LineUnitBefore = 0
        '        .LineUnitAfter = 0
    End With

    If biimp Then
        Selection.TypeText text:=vbTab
        Selection.OMaths.Add Range:=Selection.Range
        Selection.InsertSymbol CharacterNumber:=8661, Unicode:=True, Bias:=0
        Selection.OMaths(1).Range.Font.Size = 16
        Selection.MoveRight unit:=wdCharacter, Count:=1
        Selection.TypeText text:=vbTab
    End If
    Selection.Font.Size = 8
    Selection.Font.ColorIndex = wdGray50
    Selection.Font.Italic = True

    Selection.TypeText ForklarTekst
    If omax.TempDefs <> "" Then
        '                If Not MaximaSeparator Then
        '                    tdefs = Replace(omax.TempDefs, ",", ";")
        '                    tdefs = Replace(tdefs, ".", ",")
        '                Else
        tdefs = omax.TempDefs
        tdefs = Replace(tdefs, "pi", ChrW(960))
        '                End If
        Selection.TypeText Sprog.A(61)
        Set mo = Selection.OMaths.Add(Selection.Range)
        Selection.TypeText omax.ConvertToWordSymbols(tdefs)
        mo.OMaths.BuildUp
    End If

    Selection.TypeParagraph
    Selection.Font.Size = gemfontsize
    Selection.Font.Italic = gemitalic
    Selection.Font.ColorIndex = gemfontcolor
    With Selection.ParagraphFormat
        .SpaceBefore = gemsb
        '        .SpaceBeforeAuto = False
        .SpaceAfter = gemsa
        '        .SpaceAfterAuto = False
    End With

End Sub
Sub MaximaEliminate()
    Dim ForklarTekst As String
    Dim tdefs As String
    Dim mo As Range
    Dim fejlm As String
    On Error GoTo fejl
    Application.ScreenUpdating = False
    '    LockWindow
    Dim scrollpos As Double
    scrollpos = ActiveWindow.VerticalPercentScrolled

    PrepareMaxima
    omax.prevspr = ""
    '    Set UFSelectVar = New UserFormSelectVar
    Set UFSelectVar = New UserFormSelectVar
    Dim variabel As String
    Dim sstart As Long, sslut As Long
    If Not omax.MaximaInstalled Then GoTo slut
    '    If UFWait Is Nothing Then Set UFWait = New UserFormWaitForMaxima
    '    Set UFWait = New UserFormWaitForMaxima
    sstart = Selection.start
    sslut = Selection.End

    If omax.DefFejl Then
        MsgBox Sprog.DefError & vbCrLf & VisDef & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    If sstart = sslut Then
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    If InStr(Selection.OMaths(1).Range.text, "=") < 1 Then
        Dim result As VbMsgBoxResult
        result = MsgBox(Sprog.A(141), vbYesNo, Sprog.Warning)
        If result = vbNo Then GoTo slut
    End If

    If Selection.OMaths.Count < 2 And InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8743)) < 1 Then
        GoTo slut
    Else    ' ligningssystem

        omax.ReadSelection
        omax.FindVariable
        If Not ValiderVariable Then GoTo slut
        UFSelectVar.Eliminate = True
        UFSelectVar.NoEq = omax.AntalKom
        UFSelectVar.vars = omax.vars
        UFSelectVar.DefS = omax.DefString
        UFSelectVar.Caption = Sprog.A(377)
        UFSelectVar.Show
        variabel = UFSelectVar.SelectedVar

        '    variabel = InputBox("Indtast variable som ligningssystemet skal løses for adskilt af komma", "Variable", "x,y")

        If variabel = "" Then GoTo slut
        omax.TempDefs = UFSelectVar.TempDefs
        '    UFWait.ActionToPerform = "solvesystem"
        '    UFWait.VarParam = variabel
        '    Set UFWait.omax = omax
        '    UFWait.Show
        
        
        
    If CASengine = 0 Then
        omax.Eliminate (variabel)
    Else
        MsgBox "This function is not supported by GeoGebra. Try equation solver", vbOKOnly, "No support"
        GoTo slut
    End If
        
        
        If omax.StopNow Then GoTo slut
        If CheckForError Then GoTo slut
        
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
        
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph
        If omax.StopNow Then GoTo slut
        Application.ScreenUpdating = False
        '    omax.KommentarOutput = TranslateReplaceComment(omax.KommentarOutput)

        If omax.MaximaOutput = "?merror(""Anumberwasfoundwhereavariablewasexpected-`solve'"")" Then
            MsgBox Sprog.A(133), vbOKOnly, Sprog.Error
            GoTo slut
        End If

        If (omax.DefFejl = True) Then
            MsgBox Sprog.DefError & vbCrLf & VisDef & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
            GoTo slut
        End If

        variabel = Replace(omax.ConvertToWordSymbols(variabel), ";", ",")

        If MaximaForklaring Then
            '            Selection.TypeParagraph
            ForklarTekst = Sprog.A(142) & variabel & Sprog.A(143)
            InsertForklaring ForklarTekst, False
            '            Selection.TypeText ForklarTekst
            If omax.TempDefs <> "" Then
                If Not MaximaSeparator Then
                    tdefs = Replace(omax.TempDefs, ",", ";")
                    tdefs = Replace(tdefs, ".", ",")
                Else
                    tdefs = omax.TempDefs
                End If
                Selection.TypeText Sprog.A(378) & ": "
                Set mo = Selection.OMaths.Add(Selection.Range)
                Selection.TypeText tdefs
                mo.OMaths.BuildUp
            End If
            Selection.TypeParagraph
        End If

        If Len(omax.MaximaOutput) > 1 Then
            omax.InsertMaximaOutput
        Else    ' ingen løsninger
            If Len(omax.KommentarOutput) <= 1 Then
                omax.MaximaOutput = "L=" & VBA.ChrW(8709)
                '               omax.GoToEndOfSelectedMaths
                omax.InsertMaximaOutput
                Selection.TypeParagraph
                If MaximaComplex Then
                    Selection.TypeText Sprog.A(137) & variabel
                Else
                    Selection.TypeText Sprog.A(138) & variabel
                End If
            Else
                If (omax.DefFejl = True) Then
                    MsgBox Sprog.DefError & vbCrLf & VisDef & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
                Else
                    fejlm = Sprog.A(131) & vbCrLf
                End If
                If InStr(omax.Kommando, "=") < 1 Then
                    fejlm = fejlm & Sprog.A(139) & vbCrLf
                End If
                '                fejlm = fejlm & TranslateComment(omax.KommentarOutput)
                UserFormNumericQuestion.FejlMeld = omax.KommentarOutput
                UserFormNumericQuestion.Show
                If UserFormNumericQuestion.result = "num" Then
                    Selection.End = sslut    ' slut skal være først eller går det galt
                    Selection.start = sstart
                    ActiveWindow.VerticalPercentScrolled = scrollpos
                    MaximaSolveNumeric
                Else
                    Selection.TypeParagraph
                    Selection.TypeText Sprog.A(140)
                End If
                '                MsgBox fejlm & "Angav du de rigtige variable?" & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, "Fejl"
            End If
        End If
    End If
    
    Oundo.EndCustomRecord

    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    '    omax.Luk
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
    '   UnLockWindow
End Sub
Sub MaximaNsolve(Optional ByVal variabel As String)
    Dim Arr As Variant
    Dim fejlm As String
    Dim UFnsolve As New UserFormNumericQuestion
        On Error GoTo fejl
    Application.ScreenUpdating = False
    '    LockWindow
    Dim IsSolved As Boolean
    Dim scrollpos As Double
    Dim ea As New ExpressionAnalyser, s As String, V As String, t As String
    scrollpos = ActiveWindow.VerticalPercentScrolled

    '    PrepareMaximaNoSplash
    PrepareMaxima
    omax.prevspr = ""
    Set UFSelectVar = New UserFormSelectVar
    UFSelectVar.NoEq = 1
'    Set UFSelectVars = New UserFormSelectVars
'    Dim variabel As String
    Dim sstart As Long, sslut As Long
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo slut
    '    If UFWait Is Nothing Then Set UFWait = New UserFormWaitForMaxima
    '    Set UFWait = New UserFormWaitForMaxima
    sstart = Selection.start
    sslut = Selection.End

    If omax.DefFejl Then
        MsgBox Sprog.DefError & vbCrLf & VisDef & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.EquationMissingError, vbOKOnly, Sprog.EquationMissingError2
        GoTo slut
    End If
    If sstart = sslut Then
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    If InStr(Selection.OMaths(1).Range.text, "<") > 1 Or InStr(Selection.OMaths(1).Range.text, ">") > 1 Or InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8804)) > 1 Or InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8805)) > 1 Then
        MaximaSolveInequality
        GoTo slut
    End If
    If InStr(Selection.OMaths(1).Range.text, "=") < 1 Then
        Dim result As VbMsgBoxResult
        result = MsgBox(Sprog.A(141), vbYesNo, Sprog.Warning)
        If result = vbNo Then GoTo slut
    End If

    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Selection.OMaths.Count < 2 And InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8743)) < 1 Then
        ' kun 1 ligning

        UFnsolve.Ligning = omax.Kommando
        omax.FindVariable
        
        If CASengine > 0 And Not AllTrig Then ' På geogebra skal der via vba genkendes om det er trigonometrisk ligning
            If Not InStr(omax.vars, ";") > 0 Then ' metoden virker kun med 1 variabel
                ea.SetNormalBrackets
                ea.text = omax.Kommando
                ea.text = Replace(ea.text, VBA.ChrW(8289), "")
                s = ""
                Do
                    V = ea.GetNextVar()
                    If V = "sin" Or V = "cos" Or V = "tan" Then
                        t = ea.GetNextBracketContent()
                        If InStr(t, omax.vars) > 0 Then s = "90"
                    End If
                    ea.Pos = ea.Pos + 1
                Loop While V <> ""
                If s <> "" And Radians Then s = "pi/2"
                If s <> "" Then
                    UFSelectVar.TextBox_def.text = "0<=" & omax.vars & "<=" & s & VbCrLfMac
                    UFSelectVar.TempDefs = "0<=" & omax.vars & "<=" & s
                End If
            End If
        End If
        If variabel = vbNullString Then
            UFSelectVar.vars = omax.vars
            UFSelectVar.DefS = omax.DefString
            UFSelectVar.Show
            variabel = UFSelectVar.SelectedVar
        End If
        
        If variabel = "" Then GoTo slut
        omax.TempDefs = UFSelectVar.TempDefs
        
        Dim LHS As String, rhs As String
        If CASengine = 1 Or CASengine = 2 Then
            s = Trim(omax.Kommando)
            s = Replace(s, vbCrLf, "")
            s = Replace(s, vbCr, "")
            s = Replace(s, vbLf, "")
            s = "nsolve(" & Replace(s, ",", ".") & "," & variabel & ")"
        End If
        
        If CASengine = 1 Then
            If MaximaForklaring Then
                omax.GoToEndOfSelectedMaths
                Selection.TypeParagraph
                If Sprog.SprogNr = 1 Then
                    InsertForklaring "Ligningen blev løst grafisk med GeoGebra:", False
                Else
                    InsertForklaring "Equation was solved graphically using GeoGebra:", False
                End If
                Selection.TypeParagraph
            End If
            s = Replace(omax.Kommando, ",", ".")
            Arr = Split(s, "=")
            LHS = Arr(0)
            rhs = Arr(1)
            If variabel <> "x" Then
                ea.text = LHS
                ea.ReplaceVar variabel, "x"
                LHS = ea.text
                ea.text = rhs
                ea.ReplaceVar variabel, "x"
                rhs = ea.text
            End If
            OpenGeoGebraWeb "y=" & LHS & ";y=" & rhs & ";intersect(" & LHS & "," & rhs & ");" & "Nsolve(" & s & "," & variabel & ")", "CAS", True, True
            GoTo slut
        ElseIf CASengine = 2 Then
            If MaximaDecOutType = 3 Then
                s = "ScientificText(" & s & " , " & MaximaCifre & ")"
            Else
                s = "Numeric(" & s & " , " & MaximaCifre & ")"
            End If
            Call RunGeoGebraDirect(s)
            If Left(omax.MaximaOutput, 1) = "{" Then omax.MaximaOutput = Mid(omax.MaximaOutput, 2, Len(omax.MaximaOutput) - 2)
            omax.MaximaOutput = Replace(omax.MaximaOutput, ";", "    " & VBA.ChrW(8744) & "    ")    ' komma erstattes med eller
            UFnsolve.Label_nsolve.Caption = omax.MaximaOutput
            Application.Activate
'            omax.InsertMaximaOutput
'            GoTo ghop
        End If

        UFnsolve.Label_overskrift.Caption = ""
        UFnsolve.Label_omskrevet.Caption = ""
        UFnsolve.FejlMeld = ""
        UFnsolve.variabel = variabel
        UFnsolve.Show

        If omax.StopNow Then GoTo slut
        If CheckForError Then GoTo slut
        
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord

        If UFnsolve.result = "afbryd" Then GoTo slut

        If UFnsolve.result = "num" Then
            Selection.End = sslut    ' slut skal være først ellers går det galt
            Selection.start = sstart
            ActiveWindow.VerticalPercentScrolled = scrollpos
            Unload UFnsolve
            If CASengine = 0 Then
                MaximaSolveNumeric UFSelectVar.ListBox_vars.text
            Else
      s = Replace(omax.Kommando, ",", ".")
            Arr = Split(s, "=")
            LHS = Arr(0)
            rhs = Arr(1)
            If variabel <> "x" Then
                ea.text = LHS
                ea.ReplaceVar variabel, "x"
                LHS = ea.text
                ea.text = rhs
                ea.ReplaceVar variabel, "x"
                rhs = ea.text
            End If
            OpenGeoGebraWeb "y=" & LHS & ";y=" & rhs & ";intersect(" & LHS & "," & rhs & ");" & "Nsolve(" & s & "," & variabel & ")", "CAS", True, True
            End If
            GoTo slut
        End If
ghop:
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph

        If UFnsolve.result = "nsolve" Then
            InsertForklaring Sprog.EquationSolvedNumFor & variabel & Sprog.A(57), False
            omax.MaximaOutput = UFnsolve.Label_nsolve.Caption
        ElseIf UFnsolve.result = "omskriv" Then
            omax.MaximaOutput = UFnsolve.Label_omskrevet.Caption
        End If



        '    omax.KommentarOutput = TranslateReplaceComment(omax.KommentarOutput)

        If Len(omax.MaximaOutput) > 150 Then
            Dim resultat As VbMsgBoxResult
            resultat = MsgBox(Sprog.A(127) & omax.MaximaOutput, vbOKCancel, Sprog.Warning)
            If resultat = vbCancel Then GoTo slut
        End If

        variabel = omax.ConvertToWordSymbols(variabel)

        If omax.IsAllSolved(omax.MaximaOutput, variabel, VBA.ChrW(8744)) = "false" And Not (InStr(variabel, "+") > 0) Then
            IsSolved = False
        Else
            IsSolved = True
        End If

        ' indsæt forklaring hvis ønsket
        If MaximaForklaring And IsSolved Then
            InsertForklaring Sprog.EquationSolvedNumFor & variabel & Sprog.A(57)
        End If

        omax.InsertMaximaOutput

    Else    ' ligningssystem
        Dim Variable As String, guess As String, inp As String, j As Integer
        Dim Arr2 As Variant

        omax.ReadSelection
        omax.FindVariable
        UFSelectVar.NoEq = Selection.OMaths.Count
        UFSelectVar.vars = omax.vars
        UFSelectVar.Show
        Variable = UFSelectVar.SelectedVar
        Variable = Replace(Variable, ",", "=1" & ListSeparator)
        Variable = Variable & "=1"
        inp = InputBox(Sprog.A(379), Sprog.A(380), Variable)
        If inp = "" Then GoTo slut    ' trykket cancel
        Arr = Split(inp, ListSeparator)

        Variable = ""
        For j = 0 To UBound(Arr)
            Arr2 = Split(Arr(j), "=")
            Variable = Variable & Trim(Arr2(0)) & ","
            If UBound(Arr2) = 0 Or Trim(Arr2(1)) = "" Then
                guess = guess & ",1"
            Else
                guess = guess & Replace(Trim(Arr2(1)), ",", ".") & ","
            End If
        Next
        Variable = Left(Variable, Len(Variable) - 1)
        guess = Left(guess, Len(guess) - 1)

        If Variable = "" Then GoTo slut

        If CASengine = 1 Or CASengine = 2 Then
            s = "nsolve({" & Replace(omax.KommandoerStreng, ";", " , ") & "},{" & Replace(inp, ";", " , ") & "})"
        End If
        
        If CASengine = 0 Then
            omax.SolveSystemNumeric Variable, guess
        ElseIf CASengine = 1 Then
            If MaximaForklaring Then
                omax.GoToEndOfSelectedMaths
                Selection.TypeParagraph
                InsertForklaring "Ligningsystemet blev løst med GeoGebra:", False
                Selection.TypeParagraph
            End If
            If InStr(Variable, "x") > 0 And InStr(Variable, "y") > 0 Then s = omax.KommandoerStreng & ";intersect(" & Replace(omax.KommandoerStreng, ";", " , ") & ");" & s
            OpenGeoGebraWeb s, "CAS", True, True
            GoTo slut
        ElseIf CASengine = 2 Then
            If MaximaDecOutType = 3 Then
                s = "ScientificText(" & s & " , " & MaximaCifre & ")"
            Else
                s = "Numeric(" & s & " , " & MaximaCifre & ")"
            End If
            Call RunGeoGebraDirect(s)
            If Left(omax.MaximaOutput, 1) = "{" Then omax.MaximaOutput = Mid(omax.MaximaOutput, 2, Len(omax.MaximaOutput) - 2)
            omax.MaximaOutput = Replace(omax.MaximaOutput, ";", "    " & VBA.ChrW(8744) & "    ")    ' komma erstattes med eller
            UFnsolve.Label_nsolve.Caption = omax.MaximaOutput
            Application.Activate
        End If
        
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph

        If Len(omax.MaximaOutput) > 1 Then
            If MaximaForklaring Then
                If Not MaximaSeparator Then guess = Replace(guess, ",", ";")
                InsertForklaring Sprog.A(381) & " " & Variable & " " & Sprog.A(382) & " " & guess & " " & Sprog.A(34) & ". ", False
            End If
            omax.InsertMaximaOutput
        Else
            If (omax.DefFejl = True) Then
                MsgBox Sprog.DefError & vbCrLf & VisDef & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
            Else
                fejlm = Sprog.A(383) & "." & vbCrLf
            End If
            If InStr(omax.Kommando, "=") < 1 Then
                fejlm = fejlm & Sprog.A(139) & "." & vbCrLf
            End If
            fejlm = fejlm & TranslateComment(omax.KommentarOutput)
            MsgBox fejlm & Sprog.A(384) & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
        End If
    End If
    
    Oundo.EndCustomRecord

    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    '    omax.Luk
    On Error Resume Next
    Unload UFnsolve
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
    '   UnLockWindow
End Sub
Sub MaximaSolveNumeric(Optional Var As String)
' grafisk løsning + newton + intervalmetode
' var indsættes som den variabel der løses for
    On Error GoTo fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim variabel As String
    Dim t As String
    Dim scrollpos As Double
    Dim VarGuess As String
    scrollpos = ActiveWindow.VerticalPercentScrolled
    Dim sstart As Long
    Dim sslut As Long
    sstart = Selection.start
    sslut = Selection.End

    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo slut
    '    Set UFWait = New UserFormWaitForMaxima
    Dim UFSolvenumeric As New UserFormSolveNumeric

'    Set UFSelectVars = New UserFormSelectVars

    '    MsgBox WordWindowNavn
    Dim Sep As String
    If MaximaSeparator Then
        Sep = ","
    Else
        Sep = ";"
    End If

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.EquationMissingError, vbOKOnly, Sprog.EquationMissingError2
        GoTo slut
    End If
    If sstart = sslut Then
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    If InStr(Selection.OMaths(1).Range.text, "=") < 1 Then
        Dim result As VbMsgBoxResult
        result = MsgBox(Sprog.A(141), vbYesNo, Sprog.Warning)
        If result = vbNo Then GoTo slut
    End If

    If Selection.OMaths.Count < 2 Then

        '    ActivateWord WordWindowNavn & " - Microsoft Word"

        omax.ReadSelection
        omax.FindVariable
        If Var <> vbNullString Then
            variabel = Var
            UFSolvenumeric.TextBox_variabel.text = Var
        End If
        UFSolvenumeric.Udtryk = omax.Kommando
        UFSolvenumeric.Show
        '    Set omax = New CMaxima
        omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If
        omax.FindVariable
        omax.TempDefs = UFSelectVar.TempDefs

        If UFSolvenumeric.SelectedVar = "" Then GoTo slut
        If UFSolvenumeric.Method = "newton" Then
            '        UFWait.ActionToPerform = "solvenumeric"
            '        varguess = UFSolvenumeric.SelectedVar & "=" & UFSolvenumeric.TextBox_guess.text
            VarGuess = UFSolvenumeric.TextBox_guess.text
            '        UFWait.VarParam = varguess
            '        arr = Split(varguess, "=")
            '        variabel = arr(0)
            variabel = UFSolvenumeric.SelectedVar
            omax.MaximaSolveNumeric variabel, VarGuess
        ElseIf UFSolvenumeric.Method = "findroot" Then
            '        UFWait.ActionToPerform = "findroot"
            VarGuess = UFSolvenumeric.SelectedVar & ListSeparator & UFSolvenumeric.TextBox_lval.text & ListSeparator & UFSolvenumeric.TextBox_hval.text
            '        UFWait.VarParam = varguess
            variabel = UFSolvenumeric.SelectedVar
            omax.MaximaFindRoot variabel, UFSolvenumeric.TextBox_lval.text, UFSolvenumeric.TextBox_hval.text
        ElseIf UFSolvenumeric.Method = "poly" Then
            '        UFWait.ActionToPerform = "poly"
            '        UFWait.VarParam = UFSolvenumeric.SelectedVar
            variabel = UFSolvenumeric.SelectedVar
            omax.PolySolve (variabel)
        End If

        '    Set UFWait.omax = omax
        '    UFWait.Show
        If omax.StopNow Then GoTo slut
        If CheckForError Then GoTo slut
        '    omax.KommentarOutput = TranslateReplaceComment(omax.KommentarOutput)
        If InStr(omax.MaximaOutput, VBA.ChrW(8776) & "false") > 0 Then
            omax.MaximaOutput = ""
            omax.KommentarOutput = Sprog.A(385)
        End If

        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph

        If Len(omax.MaximaOutput) > 150 Then
            Dim resultat As VbMsgBoxResult
            resultat = MsgBox(Sprog.A(127) & omax.MaximaOutput, vbOKCancel, Sprog.Warning)
            If resultat = vbCancel Then GoTo slut
        End If

        If Len(omax.MaximaOutput) > 1 And InStr(omax.MaximaOutput, "Lisperror") < 1 Then
            t = Sprog.EquationSolvedNumFor & variabel & Sprog.A(147)
            If MaximaForklaring Then
                If UFSolvenumeric.Method = "newton" Then
                    t = t & " " & Sprog.A(386) & " " & variabel & "=" & VarGuess
                Else
                    t = t & " " & Sprog.A(387) & " [" & Replace(Replace(VarGuess, variabel & ";", "", 1, 1), ";", " ; ", 1, 1) & "]"
                End If
                InsertForklaring t, False
                '            Selection.TypeText t
                '            Selection.TypeParagraph
            End If
            omax.InsertMaximaOutput
            omax.GoToEndOfSelectedMaths
            Selection.TypeParagraph
            Selection.TypeText Sprog.A(148)

        Else
            Dim fejlm As String
            If (omax.DefFejl = True) Then
                MsgBox Sprog.DefError & vbCrLf & VisDef & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
            Else
                fejlm = Sprog.A(388) & "." & vbCrLf
            End If
            fejlm = fejlm & TranslateComment(omax.KommentarOutput)
            MsgBox fejlm & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.A(389)
        End If

    Else
        ' ligningssystem


    End If

    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub
Sub beregn()
    '    MsgBox WordWindowNavn
    '    Dim omax As New CMaxima
    Dim fejlm As String
        On Error GoTo fejl
   ' Application.ScreenUpdating = False
    '   LockWindow
    Dim tid As Single
    tid = Timer
#If Mac Then
    Dim D As Document
    Set D = ActiveDocument
#Else
    system.Cursor = wdCursorWait
#End If
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    '    TimeText = ""
    '    Dim st As Double
    '    st = Timer
    scrollpos = ActiveWindow.VerticalPercentScrolled
    '    Set UFWait = New UserFormWaitForMaxima
    
    If SettUseVBACAS And (MaximaExact = 2 And Not MaximaUnits And Not MaximaComplex) Then
        On Error Resume Next
        Err.Clear
        Application.Run macroname:="CASCALC"
        If Err.Number = 513 Then
            MsgBox2 Err.Description, vbOKOnly, Sprog.Error
            GoTo slut
        ElseIf Err.Number = 0 Then ' hvis ingen fejl, så er beregningen done og sat ind i Word
            GoTo slut
        End If
        Err.Clear
        On Error GoTo fejl
    End If
    
    
    If Not PrepareMaxima Then GoTo slut
    omax.prevspr = ""

    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo slut
    If Selection.OMaths.Count = 0 Then  'And Len(Selection.Range.text) < 2
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    If Selection.OMaths.Count > 1 Then
        MsgBox Sprog.A(149), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    omax.ReadSelection
    
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    If omax.Kommando = "" Then GoTo slut
    If Not ValidateInput(omax.Kommando) Then GoTo slut
    '    Set UFWait.omax = omax
    '    UFWait.ActionToPerform = "beregn"
    '    UFWait.Show
    '    If omax.StopNow Then GoTo slut
    
    Dim s As String, t As String, fo As String
    
    If CASengine > 0 Then
        s = Trim(omax.Kommando)
        '        If Left(s, 1) = "=" Then s = Left(s, Len(s) - 1)
        s = GetCmdAfterEqualSign(s)
        If MaximaDecOutType = 3 Then
            s = "ScientificText(" & s & " , " & MaximaCifre & ")"
        ElseIf MaximaExact = 2 Then
            s = "numeric(" & s & " , " & MaximaCifre & ")"
        End If
    End If
    
    If CASengine = 0 Then
        omax.beregn
    ElseIf CASengine = 1 Then
        If MaximaForklaring Then
            omax.GoToEndOfSelectedMaths
            If MaximaForklaring Then
                Selection.TypeParagraph
                InsertForklaring Sprog.A(682), False
                Selection.TypeParagraph
            End If
        End If
        OpenGeoGebraWeb s, "CAS", True, True
        GoTo slut
    ElseIf CASengine = 2 Then
        fo = RunGeoGebraDirect(s)
        If MaximaExact = 0 And MaximaDecOutType < 3 Then
            If fo = "?" Or fo = "null" Or fo = "" Then
                s = "numeric(" & s & " , " & MaximaCifre & ")"
            Else
                ' det første resultat kan ikke bare fødes ind i GeoGebra igen. Det giver problemer i særlige tilfælde. Eksempel: '\cbrt(79/138)^(2)' Her burde være parentes. Den rigtige fortolkning er cbrt((79/138)^2), som kommer frem hvis den tastes og læses i Word. Hvis den køres direkte i WordMat oversætteren indsættes ikke korrekt parentes. Normalt ikke et problem, da alt normalt læses fra Word
                '                fo = omax.ReadFromWord(fo) ' forsøg på at omgå problem med at føde resulkat direkte ind i geogebra igen
                '                s = "numeric(" & fo & " , " & MaximaCifre & ")"
                s = "numeric(" & s & " , " & MaximaCifre & ")" ' Der er eksempler, hvor det er bedre at beregne numerisk videre på det eksakte resultat istedet for direkte på det originale, men pga ovenstående problematik
            End If
            MaximaExact = 2
            t = RunGeoGebraDirect(s)
            MaximaExact = 0
            If (fo = "?" Or fo = "null" Or fo = "") And (t = "?" Or t = "null" Or t = "") Then
                omax.MaximaOutput = fo
            ElseIf (fo = "?" Or fo = "null" Or fo = "") Then
                omax.MaximaOutput = t
            ElseIf (t = "?" Or t = "null" Or t = "") Or fo = t Or t = "" Then
                omax.MaximaOutput = fo
            Else
                omax.MaximaOutput = fo & VBA.ChrW(&H2248) & t
            End If
        End If
        omax.MaximaOutput = "=" & omax.MaximaOutput
    End If
    
    If omax.StopNow Then GoTo slut
    If CheckForError Then GoTo slut
    '    TimeText = TimeText & vbCrLf & "beregn: " & Timer - st

    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord
    
    If Len(omax.MaximaOutput) > 0 And Trim(omax.MaximaOutput) <> "=res1" Then
        InsertOutput omax.MaximaOutput
    Else
        fejlm = Sprog.A(128) & vbCrLf
        fejlm = fejlm & TranslateComment(omax.KommentarOutput)
        '        If InStr(omax.KommentarOutput, "infix") > 0 Then
        '            fejlm = fejlm & "Husk at alle gangetegn skal laves. 2*x ikke 2x" & vbCrLf
        '        End If
#If Mac Then
        fejlm = fejlm & vbCrLf & vbCrLf & omax.KommentarOutput & vbCrLf
        MsgBox2 fejlm, vbOKOnly, Sprog.Error
#Else
        fejlm = fejlm & vbCrLf & vbCrLf & omax.KommentarOutput & vbCrLf & MaxProc.LastMaximaOutput
        MsgBox2 fejlm, vbOKOnly, Sprog.Error
        RestartMaxima
#End If
    End If
    Oundo.EndCustomRecord

    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
#If Mac Then
    D.Activate
#End If
    On Error Resume Next
    ActiveWindow.VerticalPercentScrolled = scrollpos
    Application.ScreenUpdating = True
    '    UnLockWindow
    '    TimeText = TimeText & vbCrLf & "beregn ialt: " & Timer - st
    '    MsgBox TimeText
'    MsgBox Timer - tid
End Sub

Function GetCmdAfterEqualSign(Kommando As String) As String
    Dim posligmed As Integer, possumtegn As Integer, posca As Integer
    Do    ' gå tilbage til nærmeste ligmed
        posligmed = InStr(Kommando, "=")
        possumtegn = InStr(Kommando, VBA.ChrW(8721))
        '    posprodtegn = InStr(Kommando, VBA.ChrW(8719))
        If possumtegn = 0 Then possumtegn = InStr(Kommando, VBA.ChrW(8719))    ' produkttegn
        If possumtegn > 0 And possumtegn < posligmed Then    ' hvis sumtegn er der =tegn som del deraf
            posligmed = 0
        End If
        posca = InStr(Kommando, VBA.ChrW(8776))
        If posca > posligmed Then posligmed = posca
        If posligmed > 0 Then
            Kommando = right(Kommando, Len(Kommando) - posligmed)
        End If
    Loop While posligmed > 0
    GetCmdAfterEqualSign = Kommando
End Function
Sub Omskriv()
    On Error GoTo fejl
    Dim s As String
    Dim UFomskriv As New UserFormOmskriv
    PrepareMaxima
    omax.prevspr = ""
    Dim scrollpos As Double
    Dim SaveSettingsCifre As Integer
    Dim SaveSettingsExact As Integer
    Dim SaveSettingsOutunits As String
    Dim SaveSettingsLog As Integer
    Dim SaveSettingsDecOutType As Integer
    
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    SaveSettingsCifre = MaximaCifre
    SaveSettingsExact = MaximaExact
    SaveSettingsOutunits = OutUnits
    SaveSettingsLog = MaximaLogOutput
    SaveSettingsDecOutType = MaximaDecOutType
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo slut
    '    MsgBox WordWindowNavn

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    If sstart = sslut Then    ' sikrer at gotoendofselectedmaths virker
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    '    omax.OpenCmd
    omax.FindVariable
    UFomskriv.vars = omax.vars
    UFomskriv.Show
    If UFomskriv.annuller Then GoTo slut

    omax.TempDefs = UFomskriv.TempDefs
    
    
    If CASengine > 0 Then
        s = omax.Kommando
        If UFomskriv.CheckBox_trigreduce.Value Then s = "trigsimplify(" & s & ")"
        If UFomskriv.CheckBox_factor.Value Then s = "factor(" & s & ")"
        If UFomskriv.CheckBox_expand.Value Then s = "expand(" & s & ")"
        If UFomskriv.CheckBox_auto.Value Then s = "simplify(" & s & ")"
        If MaximaDecOutType = 3 Then
            s = "ScientificText(" & s & " , " & MaximaCifre & ")"
        ElseIf MaximaExact = 2 Then
            s = "numeric(" & s & " , " & MaximaCifre & ")"
        End If
    End If
    
    If CASengine = 0 Then
        omax.Omskriv False, UFomskriv.CheckBox_auto.Value, UFomskriv.CheckBox_factor.Value, UFomskriv.CheckBox_expand.Value, UFomskriv.CheckBox_rationaliser.Value, UFomskriv.CheckBox_trigreduce.Value
    ElseIf CASengine = 1 Then
        If MaximaForklaring Then
            omax.GoToEndOfSelectedMaths
            Selection.TypeParagraph
            InsertForklaring Sprog.A(50), False
            Selection.TypeParagraph
        End If
        OpenGeoGebraWeb s, "CAS", True, True
        GoTo slut
    ElseIf CASengine = 2 Then
        omax.MaximaOutput = RunGeoGebraDirect(s)
    End If
    
    
    If omax.StopNow Then GoTo slut
    If CheckForError Then GoTo slut

    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord

    If Not (UFomskriv.SammeLinje) Then
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph

        If Len(omax.MaximaOutput) > 0 Then
            If MaximaForklaring Then
                s = Sprog.A(150)
                If UFomskriv.CheckBox_auto.Value Then
                    s = s & Sprog.A(151)
                End If
                If UFomskriv.CheckBox_expand.Value Then
                    s = s & Sprog.Expand & ", "
                End If
                If UFomskriv.CheckBox_factor.Value Then
                    s = s & Sprog.Factor & ", "
                End If
                If UFomskriv.CheckBox_rationaliser.Value Then
                    s = s & Sprog.Rationalize & ", "
                End If
                If UFomskriv.CheckBox_trigreduce.Value Then
                    s = s & Sprog.A(152)
                End If
                InsertForklaring s, False
            End If
            omax.InsertMaximaOutput
        Else
            Dim fejlm As String
            fejlm = Sprog.A(153) & vbCrLf

            MsgBox fejlm & vbCrLf & omax.KommentarOutput & vbCrLf & MaxProc.LastMaximaOutput, vbOKOnly, Sprog.Error
        End If
    Else    ' not newline
        InsertOutput "=" & omax.MaximaOutput
    End If
    
    Oundo.EndCustomRecord

    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    MaximaCifre = SaveSettingsCifre
    MaximaExact = SaveSettingsExact
    OutUnits = SaveSettingsOutunits
    MaximaLogOutput = SaveSettingsLog
    MaximaDecOutType = SaveSettingsDecOutType
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos

End Sub
Sub reducer()
    On Error GoTo fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim scrollpos As Double, s As String
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo slut
    '    MsgBox WordWindowNavn

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    If sstart = sslut Then    ' sikrer at gotoendofselectedmaths virker
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    If CASengine > 0 Then
        s = "simplify(" & omax.Kommando & ")"
        If MaximaDecOutType = 3 Then
            s = "ScientificText(" & s & " , " & MaximaCifre & ")"
        ElseIf MaximaExact = 2 Then
            s = "numeric(" & s & " , " & MaximaCifre & ")"
        End If
    End If
    
    If CASengine = 0 Then
        omax.Reduce
    ElseIf CASengine = 1 Then
        If MaximaForklaring Then
            omax.GoToEndOfSelectedMaths
            Selection.TypeParagraph
            InsertForklaring Sprog.A(50), False
            Selection.TypeParagraph
        End If
        OpenGeoGebraWeb s, "CAS", True, True
        GoTo slut
    ElseIf CASengine = 2 Then
        omax.MaximaOutput = RunGeoGebraDirect(s)
    End If
    
    
    If omax.StopNow Then GoTo slut
    If CheckForError Then GoTo slut
    
    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord

    omax.GoToEndOfSelectedMaths
    Selection.TypeParagraph

    '    omax.CloseCmd
    If Len(omax.MaximaOutput) > 0 Then
        If MaximaForklaring Then
            InsertForklaring Sprog.A(154), False
        End If
        omax.InsertMaximaOutput
    Else
        Dim fejlm As String
        fejlm = Sprog.A(153) & vbCrLf

        MsgBox fejlm & vbCrLf & omax.KommentarOutput & vbCrLf & MaxProc.LastMaximaOutput, vbOKOnly, Sprog.Error
    End If

    Oundo.EndCustomRecord

    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos

End Sub
Sub CompareTest()
    On Error GoTo fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    If Not omax.MaximaInstalled Then GoTo slut
    '    MsgBox WordWindowNavn

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    If sstart = sslut Then    ' sikrer at gotoendofselectedmaths virker
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    '    omax.OpenCmd
    omax.CompareTest
    If omax.StopNow Then GoTo slut
    If CheckForError Then GoTo slut
    omax.GoToEndOfSelectedMaths
    Selection.TypeParagraph
    '    omax.CloseCmd
    omax.MaximaOutput = Replace(omax.MaximaOutput, "true", Sprog.A(144))
    omax.MaximaOutput = Replace(omax.MaximaOutput, "false", Sprog.A(145))
    omax.MaximaOutput = Replace(omax.MaximaOutput, "unknown", Sprog.A(146))
    omax.MaximaOutput = Replace(omax.MaximaOutput, "dontknow", Sprog.A(146))
    If Len(omax.MaximaOutput) > 0 Then
        If MaximaForklaring Then
            InsertForklaring Sprog.A(163), False
        End If
        omax.InsertMaximaOutput
        If InStr(omax.KommentarOutput, "Numerically tested") > 0 And InStr(omax.MaximaOutput, "sand") > 0 Then
            Selection.TypeText Sprog.A(164)
            Selection.TypeParagraph
        End If
    Else
        Dim fejlm As String
        fejlm = Sprog.A(165) & vbCrLf

        MsgBox fejlm & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
    End If

    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos

End Sub
Sub faktoriser()
    On Error GoTo fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim scrollpos As Double, s As String
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    If Not omax.MaximaInstalled Then GoTo slut
    '    MsgBox WordWindowNavn

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    If sstart = sslut Then    ' sikrer at gotoendofselectedmaths virker
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut
    
    
    If CASengine > 0 Then
        s = "factor(" & omax.Kommando & ")"
        If MaximaDecOutType = 3 Then
            s = "ScientificText(" & s & " , " & MaximaCifre & ")"
        ElseIf MaximaExact = 2 Then
            s = "numeric(" & s & " , " & MaximaCifre & ")"
        End If
    End If
    
    
    If CASengine = 0 Then
        omax.Factor
    ElseIf CASengine = 1 Then
        If MaximaForklaring Then
            omax.GoToEndOfSelectedMaths
            Selection.TypeParagraph
            InsertForklaring Sprog.A(155), False
            Selection.TypeParagraph
        End If
        OpenGeoGebraWeb s, "CAS", True, True
        GoTo slut
    ElseIf CASengine = 2 Then
        omax.MaximaOutput = RunGeoGebraDirect(s)
    End If
    

    '    omax.OpenCmd
    If omax.StopNow Then GoTo slut
    If CheckForError Then GoTo slut
    
    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord

    omax.GoToEndOfSelectedMaths
    Selection.TypeParagraph
    '    omax.CloseCmd
    If Len(omax.MaximaOutput) > 0 Then
        If MaximaForklaring Then
            InsertForklaring Sprog.A(155), False
        End If
        omax.InsertMaximaOutput
    Else
        Dim fejlm As String
        fejlm = Sprog.A(156) & vbCrLf

        MsgBox fejlm & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
    End If
    
    Oundo.EndCustomRecord
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos

End Sub
Sub udvid()
    On Error GoTo fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    Dim scrollpos As Double, s As String
    scrollpos = ActiveWindow.VerticalPercentScrolled
    If Not omax.MaximaInstalled Then GoTo slut
    '    MsgBox WordWindowNavn
    '    Set UFWait = New UserFormWaitForMaxima

    If sstart = sslut Then    ' sikrer at gotoendofselectedmaths virker
        Selection.OMaths(1).ParentOMath.Range.Select
    End If

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    If CASengine > 0 Then
        s = "expand(" & omax.Kommando & ")"
        If MaximaDecOutType = 3 Then
            s = "ScientificText(" & s & " , " & MaximaCifre & ")"
        ElseIf MaximaExact = 2 Then
            s = "numeric(" & s & " , " & MaximaCifre & ")"
        End If
    End If
    
    If CASengine = 0 Then
        omax.Expand
    ElseIf CASengine = 1 Then
        If MaximaForklaring Then
            omax.GoToEndOfSelectedMaths
            Selection.TypeParagraph
            InsertForklaring Sprog.A(50), False
            Selection.TypeParagraph
        End If
        OpenGeoGebraWeb s, "CAS", True, True
        GoTo slut
    ElseIf CASengine = 2 Then
        omax.MaximaOutput = RunGeoGebraDirect(s)
    End If

    If omax.StopNow Then GoTo slut
    If CheckForError Then GoTo slut
    
    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord
    
    omax.GoToEndOfSelectedMaths
    Selection.TypeParagraph
    
    If Len(omax.MaximaOutput) > 0 Then
        If MaximaForklaring Then
            InsertForklaring Sprog.A(157), False
        End If
        omax.InsertMaximaOutput
    Else
        Dim fejlm As String
        fejlm = Sprog.A(158) & vbCrLf

        MsgBox fejlm & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
    End If
    
    Oundo.EndCustomRecord
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub
Sub Differentier()
    On Error GoTo fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim variabel As String, s As String
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    If Not omax.MaximaInstalled Then GoTo slut
    '    MsgBox WordWindowNavn
    If sstart = sslut Then    ' sikrer at gotoendofselectedmaths virker
        Selection.OMaths(1).ParentOMath.Range.Select
    End If

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    variabel = InputBox(Sprog.A(51), Sprog.Variable, "x")
    If variabel = "" Then GoTo slut
    '    omax.OpenCmd
    
    If CASengine > 0 Then
        s = "derivative(" & omax.Kommando & " , " & variabel & ")"
        If MaximaDecOutType = 3 Then
            s = "ScientificText(" & s & " , " & MaximaCifre & ")"
        ElseIf MaximaExact = 2 Then
            s = "numeric(" & s & " , " & MaximaCifre & ")"
        End If
    End If
    
    
    If CASengine = 0 Then
        omax.Differentier variabel
    ElseIf CASengine = 1 Then
        If MaximaForklaring Then
            omax.GoToEndOfSelectedMaths
            Selection.TypeParagraph
            InsertForklaring Sprog.A(50), False
            Selection.TypeParagraph
        End If
        OpenGeoGebraWeb s, "CAS", True, True
        GoTo slut
    ElseIf CASengine = 2 Then
        omax.MaximaOutput = RunGeoGebraDirect(s)
    End If
    
    
    If omax.StopNow Then GoTo slut
    If CheckForError Then GoTo slut
    
    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord
    
    omax.GoToEndOfSelectedMaths
    Selection.TypeParagraph
    '    Selection.TypeParagraph
    '    omax.CloseCmd
    If Trim(Len(omax.MaximaOutput)) > 0 Then
        If MaximaForklaring Then
            InsertForklaring Sprog.A(49), False
        End If
        omax.InsertMaximaOutput
    Else
        Dim fejlm As String
        fejlm = Sprog.A(50) & vbCrLf

        MsgBox fejlm & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
    End If
    
    Oundo.EndCustomRecord
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub
Sub Integrer()
    On Error GoTo fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim variabel As String, s As String
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    '    MsgBox WordWindowNavn
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo slut

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    If sstart = sslut Then    ' sikrer at gotoendofselectedmaths virker
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    variabel = InputBox(Sprog.A(51), Sprog.Variable, "x")
    If variabel = "" Then GoTo slut
    '    omax.OpenCmd
    
    If CASengine > 0 Then
        s = "integral(" & omax.Kommando & " , " & variabel & ")"
        If MaximaDecOutType = 3 Then
            s = "ScientificText(" & s & " , " & MaximaCifre & ")"
        ElseIf MaximaExact = 2 Then
            s = "numeric(" & s & " , " & MaximaCifre & ")"
        End If
    End If
    
    If CASengine = 0 Then
        omax.Integrer variabel
    ElseIf CASengine = 1 Then
        If MaximaForklaring Then
            omax.GoToEndOfSelectedMaths
            Selection.TypeParagraph
            InsertForklaring Sprog.A(50), False
            Selection.TypeParagraph
        End If
        OpenGeoGebraWeb s, "CAS", True, True
        GoTo slut
    ElseIf CASengine = 2 Then
        omax.MaximaOutput = RunGeoGebraDirect(s)
    End If
    
    
    If omax.StopNow Then GoTo slut
    If CheckForError Then GoTo slut
    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord

    omax.GoToEndOfSelectedMaths
    Selection.TypeParagraph
    '    Selection.TypeParagraph
    '    omax.CloseCmd
    If Trim(Len(omax.MaximaOutput)) > 0 Then

        If MaximaForklaring Then
            InsertForklaring Sprog.A(52), False
        End If
        omax.InsertMaximaOutput
    Else
        Dim fejlm As String
        fejlm = Sprog.A(53) & vbCrLf

        MsgBox fejlm & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
    End If
    
    Oundo.EndCustomRecord
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub
Sub SolveDENumeric()
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    Dim variabel As String
    Dim ea As New ExpressionAnalyser
    Dim UFdiffeq As New UserFormDeSolveNumeric
    On Error GoTo fejl

    ea.SetNormalBrackets

    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled

    PrepareMaxima

    If Not omax.MaximaInstalled Then GoTo slut

    If Selection.OMaths.Count > 0 And sstart = sslut Then    ' sikrer at gotoendofselectedmaths virker
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    omax.FindVariable
    If InStr(omax.vars, "t") > 0 Then
        variabel = "t"
    ElseIf InStr(omax.vars, "x") > 0 Then
        variabel = "x"
    Else
        variabel = "x"
    End If
    If omax.Kommando <> vbNullString Or omax.KommandoArrayLength > 0 Then
        UFdiffeq.TextBox_eq1.text = GetRHS(omax.KommandoArray(0))
        UFdiffeq.TextBox_eq2.text = GetRHS(omax.KommandoArray(1))
        UFdiffeq.TextBox_eq3.text = GetRHS(omax.KommandoArray(2))
        UFdiffeq.TextBox_eq4.text = GetRHS(omax.KommandoArray(3))
        UFdiffeq.TextBox_eq5.text = GetRHS(omax.KommandoArray(4))
        UFdiffeq.TextBox_eq6.text = GetRHS(omax.KommandoArray(5))
        UFdiffeq.TextBox_eq7.text = GetRHS(omax.KommandoArray(6))
        UFdiffeq.TextBox_eq8.text = GetRHS(omax.KommandoArray(7))
        UFdiffeq.TextBox_eq9.text = GetRHS(omax.KommandoArray(8))
        DeVarList = ""
        UFdiffeq.TextBox_var1.text = GetLHSvar(omax.KommandoArray(0))
        UFdiffeq.TextBox_var2.text = GetLHSvar(omax.KommandoArray(1))
        UFdiffeq.TextBox_var3.text = GetLHSvar(omax.KommandoArray(2))
        UFdiffeq.TextBox_var4.text = GetLHSvar(omax.KommandoArray(3))
        UFdiffeq.TextBox_var5.text = GetLHSvar(omax.KommandoArray(4))
        UFdiffeq.TextBox_var6.text = GetLHSvar(omax.KommandoArray(5))
        UFdiffeq.TextBox_var7.text = GetLHSvar(omax.KommandoArray(6))
        UFdiffeq.TextBox_var8.text = GetLHSvar(omax.KommandoArray(7))
        UFdiffeq.TextBox_var9.text = GetLHSvar(omax.KommandoArray(8))
        If InStr(DeVarList, variabel) > 0 Then
            If Not InStr(DeVarList, "x") > 0 Then
                variabel = "x"
            ElseIf Not InStr(DeVarList, "t") > 0 Then
                variabel = "t"
            ElseIf Not InStr(DeVarList, "y") > 0 Then
                variabel = "y"
            ElseIf Not InStr(DeVarList, "z") > 0 Then
                variabel = "z"
            End If
        End If
    End If
    If UFdiffeq.TextBox_eq1.text <> vbNullString And UFdiffeq.TextBox_init1.text = vbNullString Then
        UFdiffeq.TextBox_init1.text = "1"
    End If
    If UFdiffeq.TextBox_eq2.text <> vbNullString And UFdiffeq.TextBox_init2.text = vbNullString Then
        UFdiffeq.TextBox_init2.text = "1"
    End If
    If UFdiffeq.TextBox_eq3.text <> vbNullString And UFdiffeq.TextBox_init3.text = vbNullString Then
        UFdiffeq.TextBox_init3.text = "1"
    End If
    If UFdiffeq.TextBox_eq4.text <> vbNullString And UFdiffeq.TextBox_init4.text = vbNullString Then
        UFdiffeq.TextBox_init4.text = "1"
    End If
    If UFdiffeq.TextBox_eq5.text <> vbNullString And UFdiffeq.TextBox_init5.text = vbNullString Then
        UFdiffeq.TextBox_init5.text = "1"
    End If
    If UFdiffeq.TextBox_eq6.text <> vbNullString And UFdiffeq.TextBox_init6.text = vbNullString Then
        UFdiffeq.TextBox_init6.text = "1"
    End If
    If UFdiffeq.TextBox_eq7.text <> vbNullString And UFdiffeq.TextBox_init7.text = vbNullString Then
        UFdiffeq.TextBox_init7.text = "1"
    End If
    If UFdiffeq.TextBox_eq8.text <> vbNullString And UFdiffeq.TextBox_init8.text = vbNullString Then
        UFdiffeq.TextBox_init8.text = "1"
    End If
    If UFdiffeq.TextBox_eq9.text <> vbNullString And UFdiffeq.TextBox_init9.text = vbNullString Then
        UFdiffeq.TextBox_init9.text = "1"
    End If


    UFdiffeq.TextBox_varx.text = variabel

    UFdiffeq.Show
    If UFdiffeq.luk Then GoTo slut


    '    MsgBox omax.MaximaOutput

    If UFdiffeq.InsertType = 3 Then
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph
        omax.MaximaOutput = UFdiffeq.ListOutput
        omax.InsertMaximaOutput
    End If


    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub
Function GetRHS(s As String) As String
    Dim Arr As Variant
    s = omax.ConvertToAscii(s)
    Arr = Split(s, "=")
    If UBound(Arr) > 0 Then
        GetRHS = Arr(UBound(Arr))
    Else
        GetRHS = s
    End If
End Function
Function GetLHSvar(s As String) As String
    Dim Arr As Variant
    Dim ea As New ExpressionAnalyser
    Dim Var As String, i As Integer
    If s = vbNullString Then
        GetLHSvar = vbNullString
        Exit Function
    End If
    s = omax.ConvertToAscii(s)
    Arr = Split(s, "=")
    s = Arr(0)
    ea.text = s
    Do
        Var = ea.GetNextVar()
        If InStr(DeVarList, Var) <= 0 Or DeVarList = vbNullString Then
            DeVarList = DeVarList & Var & ","
            GetLHSvar = Var
            Exit Function
        End If
        i = i + 1
    Loop While Var <> vbNullString And i < 10
End Function
Sub SolveDE()
    SolveDEpar
End Sub
Sub SolveDEpar(Optional funktion As String, Optional variabel As String)
    On Error GoTo fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    Dim t As String
    Dim UFdiffeq As New UserFormDiffEq
    Dim ea As New ExpressionAnalyser
    ea.SetNormalBrackets
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    '    MsgBox WordWindowNavn
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo slut

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
'    If MaximaUnits Then ' ikke et problem for unitimage
'        MsgBox Sprog.A(159), vbOKOnly, Sprog.Error
'        GoTo slut
'    End If
    If sstart = sslut Then    ' sikrer at gotoendofselectedmaths virker
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    '    funktion = InputBox("Indtast funktion som udtrykket skal løses for", "Variabel", "y")
    '    variabel = InputBox("Indtast variabel som funktionen afhænger af", "Variabel", "x")
    If funktion = vbNullString And variabel = vbNullString Then
        omax.FindVariable
        If InStr(omax.vars, "t") > 0 Then
            variabel = "t"
        ElseIf InStr(omax.vars, "x") > 0 Then
            variabel = "x"
        Else
            variabel = "x"
        End If
        If InStr(omax.Kommando, "f(") > 0 Then
'            funktion = "f(" & variabel & ")"
            omax.Kommando = Replace(omax.Kommando, "f(" & variabel & ")", "y")
            omax.Kommando = Replace(omax.Kommando, "f^' (" & variabel & ")", "y^'")
            funktion = "y"
'        ElseIf InStr(omax.Kommando, "^' (") > 0 Then
'            ea.Text = omax.Kommando
'            funktion = ea.GetPrevVar(InStr(omax.Kommando, "^' (")) & "(" & variabel & ")"
        ElseIf InStr(omax.Kommando, "g(") > 0 Then
'            funktion = "g(" & variabel & ")"
            omax.Kommando = Replace(omax.Kommando, "g(" & variabel & ")", "y")
            omax.Kommando = Replace(omax.Kommando, "g^' (" & variabel & ")", "y^'")
            funktion = "y"
        ElseIf InStr(omax.Kommando, "g^' (") > 0 Then
'            funktion = "g(" & variabel & ")"
            omax.Kommando = Replace(omax.Kommando, "g(" & variabel & ")", "y")
            omax.Kommando = Replace(omax.Kommando, "g^' (" & variabel & ")", "y^'")
            funktion = "y"
        ElseIf InStr(omax.Kommando, "N(") > 0 Then
            funktion = "N(" & variabel & ")"
            omax.Kommando = Replace(omax.Kommando, "N(" & variabel & ")", "N")
            omax.Kommando = Replace(omax.Kommando, "N^' (" & variabel & ")", "N^'")
            funktion = "N"
        ElseIf InStr(omax.Kommando, "N") > 0 Then
            funktion = "N"
        ElseIf InStr(omax.Kommando, "y(" & variabel & ")") > 0 Then
'            funktion = "y(" & variabel & ")"
            omax.Kommando = Replace(omax.Kommando, "y(" & variabel & ")", "y")
            omax.Kommando = Replace(omax.Kommando, "y^' (" & variabel & ")", "y^'")
            funktion = "y"
        ElseIf InStr(omax.vars, "y") > 0 Then
            funktion = "y"
        Else
            ea.text = omax.Kommando
            funktion = ea.GetNextVar(1)
        End If
        UFdiffeq.vars = omax.vars
        UFdiffeq.DefS = omax.DefString
        UFdiffeq.TextBox_funktion.text = funktion
        UFdiffeq.TextBox_variabel.text = variabel
        UFdiffeq.Label_ligning.Caption = Replace(omax.ConvertToAscii(omax.Kommando), " ", "")
        UFdiffeq.Show
        If UFdiffeq.luk Then GoTo slut
        variabel = UFdiffeq.TextBox_variabel.text
        funktion = UFdiffeq.TextBox_funktion.text
    Else
        Dim Arr() As String
        Arr = Split(variabel, "=")
        If UBound(Arr) > 0 Then
            UFdiffeq.TextBox_startx.text = Arr(1)
            variabel = Arr(0)
        End If
        Arr = Split(funktion, "=")
        If UBound(Arr) > 0 Then
            UFdiffeq.TextBox_starty.text = Arr(1)
            funktion = Arr(0)
        End If
    End If
    If variabel = "" Then GoTo slut
    '    omax.OpenCmd
    omax.TempDefs = UFdiffeq.TempDefs

    If CASengine > 0 Then
        Dim s As String
        s = Replace(omax.Kommando, "^'", "'")
        s = Replace(s, vbCrLf, "")
        s = Replace(s, vbCr, "")
        s = Replace(s, vbLf, "")
        s = Replace(s, ",", ".")
        t = "," & funktion & "," & variabel
        If UFdiffeq.TextBox_startx.text <> "" And UFdiffeq.TextBox_starty.text <> "" Then
            t = t & ",{(" & Replace(UFdiffeq.TextBox_startx.text, ",", ".") & " , " & Replace(UFdiffeq.TextBox_starty.text, ",", ".") & ")"
        End If
        If UFdiffeq.TextBox_bcx.text <> "" And UFdiffeq.TextBox_bcy.text <> "" Then
            t = t & ",(" & Replace(UFdiffeq.TextBox_bcx.text, ",", ".") & " , " & Replace(UFdiffeq.TextBox_bcy.text, ",", ".") & ")}"
        ElseIf UFdiffeq.TextBox_startx.text <> "" And UFdiffeq.TextBox_starty.text <> "" Then
            t = t & "}"
        End If
        If UFdiffeq.TextBox_starty2.text <> "" Then
        t = t & ",(0," & Replace(UFdiffeq.TextBox_starty2.text, ",", ".") & ")"
        End If
        s = "solveODE(" & s & t & ")"
        If MaximaDecOutType = 3 Then
            s = "ScientificText(" & s & " , " & MaximaCifre & ")"
        ElseIf MaximaExact = 2 Then
            s = "numeric(" & s & " , " & MaximaCifre & ")"
        End If
    End If

    If CASengine = 0 Then
        omax.SolveDE funktion, variabel, UFdiffeq.TextBox_startx.text, UFdiffeq.TextBox_starty.text, UFdiffeq.TextBox_starty2.text, UFdiffeq.TextBox_bcx.text, UFdiffeq.TextBox_bcy.text
    ElseIf CASengine = 1 Then
        If MaximaForklaring Then
            omax.GoToEndOfSelectedMaths
            If MaximaForklaring Then
                Selection.TypeParagraph
                InsertForklaring Sprog.EquationSolvedFor & variabel & Sprog.ByCAS, True
                Selection.TypeParagraph
            End If
        End If
        OpenGeoGebraWeb s, "CAS", True, True
        GoTo slut
    ElseIf CASengine = 2 Then
            Call RunGeoGebraDirect(s)
            omax.MaximaOutput = Replace(omax.MaximaOutput, "{", "")
            omax.MaximaOutput = Replace(omax.MaximaOutput, "}", " ") ' mellemrummet sikrer afstand så det bliver gange efter c_1
'            If Left(omax.MaximaOutput, 1) = "{" Then omax.MaximaOutput = Mid(omax.MaximaOutput, 2, Len(omax.MaximaOutput) - 2)
            omax.MaximaOutput = Replace(omax.MaximaOutput, ";", "    " & VBA.ChrW(8744) & "    ")    ' komma erstattes med eller
            Application.Activate
    End If
    
        
    If omax.StopNow Then GoTo slut
    If CheckForError Then GoTo slut
    
    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord
    
    omax.GoToEndOfSelectedMaths
    
    
    Selection.TypeParagraph
    '    omax.CloseCmd
    If Len(omax.MaximaOutput) > 1 Then
        If MaximaForklaring Then
            funktion = Split(funktion, "(")(0)
            t = Sprog.A(390)
            If Len(UFdiffeq.TextBox_startx.text) > 0 And Len(UFdiffeq.TextBox_starty.text) > 0 And Len(UFdiffeq.TextBox_bcx.text) > 0 And Len(UFdiffeq.TextBox_bcy.text) > 0 Then
                t = t & " " & Sprog.A(391) & " " & funktion & "(" & UFdiffeq.TextBox_startx.text & ")=" & UFdiffeq.TextBox_starty.text & " og " & funktion & "(" & UFdiffeq.TextBox_starty.text & ")=" & UFdiffeq.TextBox_bcy.text
            ElseIf Len(UFdiffeq.TextBox_startx.text) > 0 And Len(UFdiffeq.TextBox_starty.text) > 0 And Len(UFdiffeq.TextBox_starty2.text) > 0 Then
                t = t & " " & Sprog.A(392) & " " & funktion & "(" & UFdiffeq.TextBox_startx.text & ")=" & UFdiffeq.TextBox_starty.text & " og " & funktion & "'(" & UFdiffeq.TextBox_startx.text & ")=" & UFdiffeq.TextBox_starty2.text
            ElseIf Len(UFdiffeq.TextBox_startx.text) > 0 And Len(UFdiffeq.TextBox_starty.text) > 0 Then
                t = t & " " & Sprog.A(393) & " " & funktion & "(" & UFdiffeq.TextBox_startx.text & ")=" & UFdiffeq.TextBox_starty.text
            End If
            InsertForklaring t, False
        End If
        '        InsertOutput omax.MaximaOutput
        omax.InsertMaximaOutput
    Else
        Dim fejlm As String
        fejlm = Sprog.A(160) & vbCrLf

        MsgBox fejlm & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
    End If

    Oundo.EndCustomRecord

    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først ellers går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub
Function TranslateComment(comm As String) As String
    If InStr(comm, "infix") > 0 Then
        TranslateComment = Sprog.A(161)
    ElseIf InStr(comm, "prefix") > 0 Then
        TranslateComment = ""
    ElseIf InStr(comm, "function was expected") > 0 Then
        TranslateComment = Sprog.A(161)
    ElseIf InStr(comm, "limit of integration must be real") > 0 Then
        TranslateComment = Sprog.A(162)
    Else
        TranslateComment = ""
    End If
End Function
Function TranslateReplaceOutput(comm As String) As String

    comm = Replace(comm, "?merror(""Anumberwasfoundwhereavariablewasexpected-`solve'"")", "Fejl: Du har bedt om at løse en ligning for en variabel der allerede er defineret." & vbCrLf & "Sæt en 'slet def:' kommando ind før ligningen.")

    TranslateReplaceOutput = comm
End Function
Sub InsertOutput(text As String)
' indsætter text efter den mathobj cursoren er i på sikker måde
    Dim IsBold As Boolean
    If Selection.start = Selection.End Then
        omax.GoToEndOfSelectedMaths
    Else
        Selection.Collapse wdCollapseEnd
    End If
    If Selection.OMaths.Count > 0 Then
        Selection.OMaths(1).Range.Select
        Selection.Collapse wdCollapseEnd
    Else
    End If
    Selection.TypeText text    ' giver problemer med at =tegn kommer under brøkstreg
    Selection.Move wdCharacter, -1
    If Selection.OMaths(1).Range.Font.Bold Then
        IsBold = True
        Selection.OMaths(1).Range.Font.Bold = False
    End If
    Selection.OMaths.BuildUp
    If IsBold Then Selection.OMaths(1).Range.Font.Bold = True
    omax.GoToEndOfSelectedMaths
End Sub

#If Mac Then
#Else
Sub LukMaximaProcess()
'    MaxProc.CloseProcess
    MaxProc.ConsoleInterrupt
End Sub
#End If

Function ValidateInput(Expr) As Boolean
    Dim ED As ErrorDefinition
   
    ValidateInput = True
    '    ' validate brackets
   
    If GetCountOfChar(Expr, "(") <> GetCountOfChar(Expr, ")") Then
        If Sprog.SprogNr = 1 Then
'            MsgBox "Antallet af parenteser passer ikke i udtrykket" & vbCrLf & vbCrLf & Expr, vbOKOnly, "Syntaks fejl"
            ED.Title = "Syntaks fejl"
            ED.Description = "Antallet af parenteser passer ikke i udtrykket"
            ED.MaximaOutput = Expr
        Else
'            MsgBox "The number of brackets do not match in" & vbCrLf & vbCrLf & Expr, vbOKOnly, "Syntax error"
            ED.Title = "Syntax error"
            ED.Description = "The number of brackets do not match"
            ED.MaximaOutput = Expr
        End If
    ElseIf InStr(Expr, "\left(") > 0 Or InStr(Expr, "\ast") > 0 Then
        If Sprog.SprogNr = 1 Then
'            MsgBox "Du har en forkert indstilling i Word." & vbCrLf & "I Ligningsmenuen skal du skifte fra Latex til Unicode for at WordMat virker", vbOKOnly, "Forkert indstilling"
            ED.Title = "Indstillingsfejl"
            ED.Description = "Du har en forkert indstilling i Word." & vbCrLf & "I Ligningsmenuen skal du skifte fra Latex til Unicode for at WordMat virker"
        Else
'            MsgBox "You have a wrong setting in Words equation menu." & vbCrLf & "Change from Latex to Unicode.", vbOKOnly, "Wrong setting"
            ED.Title = "Setting error"
            ED.Description = "You have a wrong setting in Words equation menu." & vbCrLf & "Change from Latex to Unicode."
        End If
    End If
   
    If ED.Title <> vbNullString Then
        UserFormError.SetErrorDefinition ED
        UserFormError.Show
        ValidateInput = False
    End If
    
End Function

Public Function GetVersion(n As String) As Single
' finder versionsnr ud fra mappenavn på maxima
Dim V As String, p As Integer
    p = InStr(n, "-")
    If p > 0 Then
        V = right(n, Len(n) - p)
    End If
    GetVersion = val(V)
End Function

#If Mac Then
#Else
Function GetMaximaPath() As String
' Finder Maximastien. søger i Appdata og programfiles. Hvis der er Maxima begge steder så bruges nyeste version
' Hvis der er samme version begge steder, så returneres stien til appdata.

    Dim FN As String, DN As String, V As Single ' Til AppData
    Dim FN1 As String, DN1 As String, V1 As Single ' Til Program files
    Dim s As String
    On Error Resume Next

    DN = Environ("AppData") & "\WordMat\"
    FN = Dir(DN & "Maxima*", vbDirectory)
    V = GetVersion(FN)
    Do
        s = Dir()
        If GetVersion(s) > V Then
            FN = s
            V = GetVersion(FN)
        End If
    Loop While s <> ""
    
    DN1 = GetProgramFilesDir() & "\WordMat\"
    FN1 = Dir(DN1 & "Maxima*", vbDirectory)
    If FN = vbNullString Then
        DN = DN1
        FN = FN1
        V = GetVersion(FN)
    Else
        V1 = GetVersion(FN1)
        If V1 > V Then
            DN = DN1
            FN = FN1
            V = V1
        End If
    End If
            
    GetMaximaPath = DN & FN
End Function
#End If

Function AskSignFromForm(Udtryk As String) As Integer
' 0=zero   1=neg   2=pos    3= nochoice
    Dim UF As New UserFormAskSign
    UF.Label_udtryk.Caption = Udtryk
    UF.Show
    
    If UF.OptionButton_nul.Value Then
        AskSignFromForm = 0
    ElseIf UF.OptionButton_negativ.Value Then
        AskSignFromForm = 1
    ElseIf UF.OptionButton_positiv.Value Then
        AskSignFromForm = 2
    Else
        AskSignFromForm = 3
    End If
    Unload UF
End Function

Function ValiderVariable() As Boolean
    Dim Arr() As String, i As Integer, s As String, ED As ErrorDefinition
'    On Error Resume Next
    ValiderVariable = True
    If omax.AntalVars > 0 Then
        Arr = Split(omax.vars, ";")
        For i = 0 To UBound(Arr)
            If Left(Arr(i), 1) = "_" Then
                ValiderVariable = False
                If Sprog.SprogNr = 1 Then
                    ED.Title = "Syntaksfejl - Ikke korrekt matematik."
                    ED.DefFejl = False
                    ED.Description = "Du har indtastes noget som ikke er korrekt matematisk notation." & vbCrLf & "Fejlen er relateret til en variabel: " & Arr(i) & vbCrLf & vbCrLf & "Det kan fx skyldes, at du har sænket skrift på et tal. " & vbCrLf & "K" & VBA.ChrW(8345) & " skal erstattes med et tal og ikke med fx 5000" & VBA.ChrW(8345) & ". Der skal bare stå 5000."
                Else
                    ED.Title = "Syntax error - Not correct math."
                    ED.DefFejl = False
                    ED.Description = "You have typed something that is not correct mathematical notation" & vbCrLf & "The error is related to a variable: " & Arr(i) & vbCrLf & "This can be caused by a subscript on a number." & vbCrLf & "K" & VBA.ChrW(8345) & " must be replaced by fx 5000" & VBA.ChrW(8345) & ". Just write 5000."
                End If
                UserFormError.SetErrorDefinition ED
                UserFormError.Show
                Exit Function
            End If
        Next
    End If
End Function
