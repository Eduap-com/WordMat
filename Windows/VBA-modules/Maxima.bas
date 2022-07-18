Attribute VB_Name = "Maxima"
Option Explicit
'Public UFWait As UserFormWaitForMaxima
Public UFSelectVar As UserFormSelectVar
Public UFSelectVars As UserFormSelectVars
Public omax As CMaxima
Public tid As Double
Private DeVarList As String
Private TempCas As Integer

Public Sub PrepareMaxima(Optional Unit As Boolean = False)
    '    Dim UFwait2 As UserFormWaitForMaxima

    On Error GoTo fejl
    Dim op As Boolean
    If DebugWM Then
        UserFormDebug.Label_time.Caption = ""
        tid = Timer
    End If
#If Mac Then
    Dim d As Document
#End If
    
    SaveBackup
    
    Dim UfWait2 As UserFormWaitStartup
    op = False
    If Not SettingsRead Then ReadAllSettingsFromRegistry

    If omax Is Nothing Then
        LavRCMenu    ' h*oe*jreklikmenu p*aa* ligninger
        SetMathAutoCorrect
        If UfWait2 Is Nothing Then Set UfWait2 = New UserFormWaitStartup
        UfWait2.Show vbModeless
        op = True
#If Mac Then
        Set d = ActiveDocument
#End If
        DoEvents
        '        Wait (0.1)
        Set omax = New CMaxima
    End If

    If MaxProc Is Nothing And CASengine = 0 Then
        If Not op Then
#If Mac Then
            Set d = ActiveDocument
#Else
            If UfWait2 Is Nothing Then Set UfWait2 = New UserFormWaitStartup
            UfWait2.Show vbModeless
            op = True
#End If
            DoEvents
        End If
        On Error Resume Next
        
#If Mac Then
#Else
        Set MaxProc = GetMaxProc() 'CreateObject("MaximaProcessClass")
        If Err.Number <> 0 Then
            MsgBox Sprog.A(54), vbOKOnly, Sprog.Error
            GoTo slut
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
                If UfWait2 Is Nothing Then Set UfWait2 = New UserFormWaitStartup
                UfWait2.Show vbModeless
                op = True
#If Mac Then
                Set d = ActiveDocument
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
    omax.ConvertLnLog = True ' andre funktioner kan *ae*ndre denne. den nulstilles
    omax.PrepareNewCommand    ' nulstiller og finder definitioner

    On Error Resume Next
    If op Then
        Unload UfWait2
#If Mac Then
        d.Activate
#End If
    End If
    GoTo slut
fejl:
    On Error Resume Next
    Unload UfWait2
slut:
End Sub
#If Mac Then
Function GetMaxProc() As MaximaProcess
        Set GetMaxProc = New MaximaProcess
End Function
#Else
Function GetMaxProc() As Object
        Set GetMaxProc = CreateObject("MaximaProcessClass")
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
        LavRCMenu    ' h*oe*jreklikmenu p*aa* ligninger
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
    
    If Not CASengine = 0 Then Exit Sub
    
    If Not MaxProc Is Nothing Then
        MaxProc.CloseProcess
        MaxProc.StartMaximaProcess
    Else
        '        Set MaxProc = New MathMenu.MaximaProcessClass
        On Error Resume Next
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
#If Mac Then
#Else
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
    Dim sstart As Long, sslut As Long
    On Error GoTo fejl
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled

    '    Set UFWait = New UserFormWaitForMaxima
    PrepareMaxima
    omax.prevspr = ""
    If Not omax.MaximaInstalled Then GoTo slut
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se p*aa* felter med := defligmed og := symbol
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


    If omax.CheckForError Then GoTo slut

    omax.GoToEndOfSelectedMaths
    Selection.TypeParagraph
    omax.InsertMaximaOutput
    '   UFWait.Hide
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal v*ae*re f*oe*rst eller g*aa*r det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos

End Sub
Sub MaximaSolveInequality(Optional variabel As String)
' l*oe*ser een ulighed
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
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se p*aa* felter med := defligmed og := symbol
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
        
        Dim Res As String, d As String, td As Variant
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
            If MaximaVidNotation Then
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
        
        
        If omax.CheckForError Then GoTo slut

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
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal v*ae*re f*oe*rst eller g*aa*r det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos

End Sub
Sub MaximaSolve()
' Denne konstruktion er n*oe*dvendig for at kunne kalde MaximaSolve med parametre fra TestTabel
' Man kan ikke tilknytte genveje subs der tager parametre selvom de optional
    
'    If CASengine = 1 Then
'        GeoGebraWeb "CAS", "solve"
'        Exit Sub
'    End If
    
    MaximaSolvePar

End Sub
'Sub MaximaSolve(Optional variabel As String)
Sub MaximaSolvePar(Optional variabel As String)
    Dim Arr As Variant, s As String, t As String, v As String
    Dim fejlm As String
    On Error GoTo fejl
    Application.ScreenUpdating = False
    '    LockWindow
    Dim IsSolved As Boolean
    Dim scrollpos As Double
    Dim UFSolvenumeric As New UserFormNumericQuestion
    Dim ea As New ExpressionAnalyser, SaveKommando As String
    scrollpos = ActiveWindow.VerticalPercentScrolled
    TempCas = CASengine
    '    PrepareMaximaNoSplash  ' ved ikke hvorfor det var nosplash, erstattet
    PrepareMaxima
    omax.prevspr = ""

    Set UFSelectVar = New UserFormSelectVar
    Set UFSelectVars = New UserFormSelectVars
    Dim sstart As Long, sslut As Long
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo slut
    '    If UFWait Is Nothing Then Set UFWait = New UserFormWaitForMaxima
    '    Set UFWait = New UserFormWaitForMaxima
    sstart = Selection.start
    sslut = Selection.End

    If omax.deffejl Then
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
        MaximaSolveInequality variabel
        GoTo slut
    End If
    If InStr(Selection.OMaths(1).Range.text, "=") < 1 Then
        Dim Result As VbMsgBoxResult
        Result = MsgBox(Sprog.A(141), vbYesNo, Sprog.Warning)
        If Result = vbNo Then GoTo slut
    End If

    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se p*aa* felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    If Selection.OMaths.Count < 2 And InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8743)) < 1 Then
        ' kun 1 ligning

        UFSolvenumeric.Ligning = omax.Kommando

        omax.FindVariable
        SaveKommando = omax.Kommando
newcas:
        omax.StopNow = False
        omax.Kommando = SaveKommando
        If CASengine > 0 And Not AllTrig Then ' P*aa* geogebra skal der via vba genkendes om det er trigonometrisk ligning
            If Not InStr(omax.vars, ";") > 0 Then ' metoden virker kun med 1 variabel
                ea.SetNormalBrackets
                ea.text = omax.Kommando
                ea.text = Replace(ea.text, VBA.ChrW(8289), "")
                s = ""
                Do
                    v = ea.GetNextVar()
                    If v = "sin" Or v = "cos" Or v = "tan" Then
                        t = ea.GetNextBracketContent()
                        If InStr(t, omax.vars) > 0 Then
                            If v = "cos" Then
                                s = "180"
                            Else
                                s = "90"
                            End If
                        End If
                    End If
                    ea.Pos = ea.Pos + 1
                Loop While v <> ""
                If s <> "" And Radians Then
                    If s = "90" Then
                        s = "pi/2"
                    Else
                        s = "pi"
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
        omax.TempDefs = UFSelectVar.TempDefs
        
        Dim Res As String, d As String, td As Variant
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
            If MaximaVidNotation Then
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
        If omax.CheckForError Then GoTo slut

#If Mac Then
#Else
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
#End If
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph
        If Len(omax.MaximaOutput) > 150 Then
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

        ' inds*ae*t forklaring hvis *oe*nsket
        If MaximaForklaring And (IsSolved Or InStr(omax.KommentarOutput, "solving system of equations")) Then
            '            InsertForklaring "Ligningen l*oe*ses for " & variabel & " vha. CAS-v*ae*rkt*oe*jet " & AppNavn & ". "
            InsertForklaring Sprog.EquationSolvedFor & variabel & Sprog.ByCAS
        End If

        fejlm = TranslateComment(omax.KommentarOutput)

        If Len(fejlm) > 0 Then
            fejlm = Sprog.A(128) & vbCrLf & fejlm
            MsgBox fejlm & vbCrLf & vbCrLf & omax.KommentarOutput & vbCrLf & MaxProc.LastMaximaOutput, vbOKOnly, Sprog.Error
            RestartMaxima
        ElseIf InStr(omax.MaximaOutput, VBA.ChrW(8709)) Then    ' ingen l*oe*sning
            omax.InsertMaximaOutput
            Selection.TypeParagraph
            If MaximaComplex Then
                Selection.TypeText Sprog.A(129) & variabel
            Else
                Selection.TypeText Sprog.A(130) & variabel
            End If
        ElseIf InStr(omax.MaximaOutput, VBA.ChrW(8477)) > 0 Then  ' alle er l*oe*sning
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
                UFSolvenumeric.Result = ""
            Else
                UFSolvenumeric.Label_omskrevet.Caption = omax.MaximaOutput
                omax.PrepareNewCommand
                omax.TempDefs = UFSelectVar.TempDefs
                UFSolvenumeric.FejlMeld = omax.KommentarOutput
                UFSolvenumeric.variabel = variabel
                UFSolvenumeric.Show
            End If
            If UFSolvenumeric.Result = "num" Then
                Selection.End = sslut    ' slut skal v*ae*re f*oe*rst eller g*aa*r det galt
                Selection.start = sstart
                ActiveWindow.VerticalPercentScrolled = scrollpos
                MaximaSolveNumeric UFSelectVar.ListBox_vars.text
            Else
                If UFSolvenumeric.Result = "nsolve" Then
                    InsertForklaring Sprog.EquationSolvedNumFor & variabel & Sprog.ByCAS, False
                    omax.MaximaOutput = UFSolvenumeric.Label_nsolve.Caption
                ElseIf UFSolvenumeric.Result = "omskriv" Then
                    InsertForklaring "", True
                    omax.MaximaOutput = UFSolvenumeric.Label_omskrevet.Caption
                End If
                omax.InsertMaximaOutput
                Selection.TypeParagraph
                Selection.TypeText UFSolvenumeric.Kommentar
            End If

        Else    ' hvis der er l*oe*sning
            omax.InsertMaximaOutput
        End If
#If Mac Then
#Else
        Oundo.EndCustomRecord
#End If

    Else    ' ligningssystem

        omax.FindVariable
        UFSelectVars.NoEq = omax.AntalKom
        UFSelectVars.vars = omax.vars
        UFSelectVars.DefS = omax.DefString
        UFSelectVars.Show
        variabel = UFSelectVars.SelectedVar

        '    variabel = InputBox("Indtast variable som ligningssystemet skal l*oe*ses for adskilt af komma", "Variable", "x,y")

        If variabel = "" Then GoTo slut
        omax.TempDefs = UFSelectVars.TempDefs
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
            If MaximaVidNotation Then
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
                omax.ConvertOutputToSolvedGGB ' burde m*aa*ske forbedres med ea.getnextlistitem ligesom solve
            End If
            Application.Activate
        End If
        
        If omax.StopNow Then GoTo slut
        If omax.CheckForError Then GoTo slut
        If omax.StopNow Then GoTo slut
        Application.ScreenUpdating = False
        '    omax.KommentarOutput = TranslateReplaceComment(omax.KommentarOutput)

        If omax.MaximaOutput = "?merror(""Anumberwasfoundwhereavariablewasexpected-`solve'"")" Then
            MsgBox Sprog.A(133), vbOKOnly, Sprog.Error
            GoTo slut
        End If

        If (omax.deffejl = True) Then
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
                    '                    Result = MsgBox("Maxima kunne ikke l*oe*se ligningssystemet. Den var for kompleks." & vbCrLf & vbCrLf & omax.KommentarOutput & vbCrLf & vbCrLf & "Tryk OK hvis du vil fors*oe*ge at l*oe*se ligningen numerisk.", vbOKCancel, "Fejl")
                    UFSolvenumeric.FejlMeld = omax.KommentarOutput
                    UFSolvenumeric.Show
                    If UFSolvenumeric.Result = "num" Then
                        Selection.End = sslut    ' slut skal v*ae*re f*oe*rst eller g*aa*r det galt
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
                Selection.TypeText "GeoGebra har v*ae*ret anvendt til at l*oe*se ligningssystemet. Det er usikkert om der kan v*ae*re l*oe*sninger. Det anbefales at fors*oe*ge med anden metode. Fx Maxima, eller numerisk/grafisk"
            End If
        Else    ' ingen l*oe*sninger
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
                If (omax.deffejl = True) Then
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
                If UFSolvenumeric.Result = "num" Then
                    Selection.End = sslut    ' slut skal v*ae*re f*oe*rst eller g*aa*r det galt
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
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    '    omax.Luk
    On Error Resume Next
    CASengine = TempCas
    Selection.End = sslut    ' slut skal v*ae*re f*oe*rst eller g*aa*r det galt
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
        Selection.MoveRight Unit:=wdCharacter, Count:=1
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
    Set UFSelectVars = New UserFormSelectVars
    Dim variabel As String
    Dim sstart As Long, sslut As Long
    If Not omax.MaximaInstalled Then GoTo slut
    '    If UFWait Is Nothing Then Set UFWait = New UserFormWaitForMaxima
    '    Set UFWait = New UserFormWaitForMaxima
    sstart = Selection.start
    sslut = Selection.End

    If omax.deffejl Then
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
        Dim Result As VbMsgBoxResult
        Result = MsgBox(Sprog.A(141), vbYesNo, Sprog.Warning)
        If Result = vbNo Then GoTo slut
    End If

    If Selection.OMaths.Count < 2 And InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8743)) < 1 Then
        GoTo slut
    Else    ' ligningssystem

        omax.ReadSelection
        omax.FindVariable
        UFSelectVars.Eliminate = True
        UFSelectVars.NoEq = omax.AntalKom
        UFSelectVars.vars = omax.vars
        UFSelectVars.DefS = omax.DefString
        UFSelectVars.Caption = Sprog.A(377)
        UFSelectVars.Show
        variabel = UFSelectVars.SelectedVar

        '    variabel = InputBox("Indtast variable som ligningssystemet skal l*oe*ses for adskilt af komma", "Variable", "x,y")

        If variabel = "" Then GoTo slut
        omax.TempDefs = UFSelectVars.TempDefs
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
        If omax.CheckForError Then GoTo slut
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph
        If omax.StopNow Then GoTo slut
        Application.ScreenUpdating = False
        '    omax.KommentarOutput = TranslateReplaceComment(omax.KommentarOutput)

        If omax.MaximaOutput = "?merror(""Anumberwasfoundwhereavariablewasexpected-`solve'"")" Then
            MsgBox Sprog.A(133), vbOKOnly, Sprog.Error
            GoTo slut
        End If

        If (omax.deffejl = True) Then
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
        Else    ' ingen l*oe*sninger
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
                If (omax.deffejl = True) Then
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
                If UserFormNumericQuestion.Result = "num" Then
                    Selection.End = sslut    ' slut skal v*ae*re f*oe*rst eller g*aa*r det galt
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
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    '    omax.Luk
    On Error Resume Next
    Selection.End = sslut    ' slut skal v*ae*re f*oe*rst eller g*aa*r det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
    '   UnLockWindow
End Sub
Sub MaximaNsolve(Optional ByVal variabel As String)
    Dim Arr As Variant
    Dim fejlm As String
    Dim solutions As String
    Dim UFnsolve As New UserFormNumericQuestion
    '    On Error GoTo fejl
    Application.ScreenUpdating = False
    '    LockWindow
    Dim IsSolved As Boolean
    Dim scrollpos As Double
    Dim ea As New ExpressionAnalyser, s As String, v As String, t As String
    scrollpos = ActiveWindow.VerticalPercentScrolled

    '    PrepareMaximaNoSplash
    PrepareMaxima
    omax.prevspr = ""
    Set UFSelectVar = New UserFormSelectVar
    Set UFSelectVars = New UserFormSelectVars
'    Dim variabel As String
    Dim sstart As Long, sslut As Long
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo slut
    '    If UFWait Is Nothing Then Set UFWait = New UserFormWaitForMaxima
    '    Set UFWait = New UserFormWaitForMaxima
    sstart = Selection.start
    sslut = Selection.End

    If omax.deffejl Then
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
        Dim Result As VbMsgBoxResult
        Result = MsgBox(Sprog.A(141), vbYesNo, Sprog.Warning)
        If Result = vbNo Then GoTo slut
    End If

    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se p*aa* felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Selection.OMaths.Count < 2 And InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8743)) < 1 Then
        ' kun 1 ligning

        UFnsolve.Ligning = omax.Kommando
        omax.FindVariable
        
        If CASengine > 0 And Not AllTrig Then ' P*aa* geogebra skal der via vba genkendes om det er trigonometrisk ligning
            If Not InStr(omax.vars, ";") > 0 Then ' metoden virker kun med 1 variabel
                ea.SetNormalBrackets
                ea.text = omax.Kommando
                ea.text = Replace(ea.text, VBA.ChrW(8289), "")
                s = ""
                Do
                    v = ea.GetNextVar()
                    If v = "sin" Or v = "cos" Or v = "tan" Then
                        t = ea.GetNextBracketContent()
                        If InStr(t, omax.vars) > 0 Then s = "90"
                    End If
                    ea.Pos = ea.Pos + 1
                Loop While v <> ""
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
        
        Dim lhs As String, rhs As String
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
                InsertForklaring "Ligningen blev l*oe*st med GeoGebra:", False
                Selection.TypeParagraph
            End If
            s = Replace(omax.Kommando, ",", ".")
            Arr = Split(s, "=")
            lhs = Arr(0)
            rhs = Arr(1)
            If variabel <> "x" Then
                ea.text = lhs
                ea.ReplaceVar variabel, "x"
                lhs = ea.text
                ea.text = rhs
                ea.ReplaceVar variabel, "x"
                rhs = ea.text
            End If
            OpenGeoGebraWeb "y=" & lhs & ";y=" & rhs & ";intersect(" & lhs & "," & rhs & ");" & "Nsolve(" & s & "," & variabel & ")", "CAS", True, True
            GoTo slut
        ElseIf CASengine = 2 Then
            If MaximaVidNotation Then
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
        If omax.CheckForError Then GoTo slut
        If UFnsolve.Result = "afbryd" Then GoTo slut

        If UFnsolve.Result = "num" Then
            Selection.End = sslut    ' slut skal v*ae*re f*oe*rst ellers g*aa*r det galt
            Selection.start = sstart
            ActiveWindow.VerticalPercentScrolled = scrollpos
            Unload UFnsolve
            If CASengine = 0 Then
                MaximaSolveNumeric UFSelectVar.ListBox_vars.text
            Else
      s = Replace(omax.Kommando, ",", ".")
            Arr = Split(s, "=")
            lhs = Arr(0)
            rhs = Arr(1)
            If variabel <> "x" Then
                ea.text = lhs
                ea.ReplaceVar variabel, "x"
                lhs = ea.text
                ea.text = rhs
                ea.ReplaceVar variabel, "x"
                rhs = ea.text
            End If
            OpenGeoGebraWeb "y=" & lhs & ";y=" & rhs & ";intersect(" & lhs & "," & rhs & ");" & "Nsolve(" & s & "," & variabel & ")", "CAS", True, True
            End If
            GoTo slut
        End If
ghop:
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph

        If UFnsolve.Result = "nsolve" Then
            InsertForklaring Sprog.EquationSolvedNumFor & variabel & Sprog.A(57), False
            omax.MaximaOutput = UFnsolve.Label_nsolve.Caption
        ElseIf UFnsolve.Result = "omskriv" Then
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

        ' inds*ae*t forklaring hvis *oe*nsket
        If MaximaForklaring And IsSolved Then
            InsertForklaring Sprog.EquationSolvedNumFor & variabel & Sprog.A(57)
        End If

        omax.InsertMaximaOutput

    Else    ' ligningssystem
        Dim Variable As String, guess As String, inp As String, j As Integer
        Dim arr2 As Variant

        omax.ReadSelection
        omax.FindVariable
        UFSelectVars.NoEq = Selection.OMaths.Count
        UFSelectVars.vars = omax.vars
        UFSelectVars.Show
        Variable = UFSelectVars.SelectedVar
        Variable = Replace(Variable, ",", "=1" & ListSeparator)
        Variable = Variable & "=1"
        inp = InputBox(Sprog.A(379), Sprog.A(380), Variable)
        If inp = "" Then GoTo slut    ' trykket cancel
        Arr = Split(inp, ListSeparator)

        Variable = ""
        For j = 0 To UBound(Arr)
            arr2 = Split(Arr(j), "=")
            Variable = Variable & Trim(arr2(0)) & ","
            If UBound(arr2) = 0 Or Trim(arr2(1)) = "" Then
                guess = guess & ",1"
            Else
                guess = guess & Replace(Trim(arr2(1)), ",", ".") & ","
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
                InsertForklaring "Ligningsystemet blev l*oe*st med GeoGebra:", False
                Selection.TypeParagraph
            End If
            If InStr(Variable, "x") > 0 And InStr(Variable, "y") > 0 Then s = omax.KommandoerStreng & ";intersect(" & Replace(omax.KommandoerStreng, ";", " , ") & ");" & s
            OpenGeoGebraWeb s, "CAS", True, True
            GoTo slut
        ElseIf CASengine = 2 Then
            If MaximaVidNotation Then
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
            If (omax.deffejl = True) Then
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
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    '    omax.Luk
    On Error Resume Next
    Unload UFnsolve
    On Error Resume Next
    Selection.End = sslut    ' slut skal v*ae*re f*oe*rst eller g*aa*r det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
    '   UnLockWindow
End Sub
Sub MaximaSolveNumeric(Optional var As String)
' grafisk l*oe*sning + newton + intervalmetode
' var inds*ae*ttes som den variabel der l*oe*ses for
    On Error GoTo fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim variabel As String
    Dim t As String
    Dim scrollpos As Double
    Dim VarGuess As String
    Dim Arr As Variant
    Dim arr2 As Variant
    Dim j As Integer
    scrollpos = ActiveWindow.VerticalPercentScrolled
    Dim sstart As Long
    Dim sslut As Long
    sstart = Selection.start
    sslut = Selection.End

    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo slut
    '    Set UFWait = New UserFormWaitForMaxima
    Dim UFSolvenumeric As New UserFormSolveNumeric

    Set UFSelectVars = New UserFormSelectVars

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
        Dim Result As VbMsgBoxResult
        Result = MsgBox(Sprog.A(141), vbYesNo, Sprog.Warning)
        If Result = vbNo Then GoTo slut
    End If

    If Selection.OMaths.Count < 2 Then

        '    ActivateWord WordWindowNavn & " - Microsoft Word"

        omax.ReadSelection
        omax.FindVariable
        If var <> vbNullString Then
            variabel = var
            UFSolvenumeric.TextBox_variabel.text = var
        End If
        UFSolvenumeric.udtryk = omax.Kommando
        UFSolvenumeric.Show
        '    Set omax = New CMaxima
        omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se p*aa* felter med := defligmed og := symbol
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
        If omax.CheckForError Then GoTo slut
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
            If (omax.deffejl = True) Then
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
    Selection.End = sslut    ' slut skal v*ae*re f*oe*rst eller g*aa*r det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub
Sub beregn()
'    MsgBox WordWindowNavn
'    Dim omax As New CMaxima
    Dim fejlm As String
'    On Error GoTo fejl
    '    Application.ScreenUpdating = False
    '   LockWindow
    
#If Mac Then
    Dim d As Document
    Set d = ActiveDocument
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
    PrepareMaxima
    omax.prevspr = ""

    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo slut
    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    If Selection.OMaths.Count > 1 Then
        MsgBox Sprog.A(149), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se p*aa* felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    If omax.Kommando = "" Then GoTo slut
    If Not ValidateInput(omax.Kommando) Then GoTo slut
    '    Set UFWait.omax = omax
    '    UFWait.ActionToPerform = "beregn"
    '    UFWait.Show
    '    If omax.StopNow Then GoTo slut
    
    Dim s As String, Res As String, def As String, Arr() As String, i As Integer, ms As String, t As String, fo As String
    
    If CASengine > 0 Then
        s = Trim(omax.Kommando)
'        If Left(s, 1) = "=" Then s = Left(s, Len(s) - 1)
        s = GetCmdAfterEqualSign(s)
        If MaximaVidNotation Then
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
        If MaximaExact = 0 And Not MaximaVidNotation Then
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
    If omax.CheckForError Then GoTo slut
    '    TimeText = TimeText & vbCrLf & "beregn: " & Timer - st

#If Mac Then
#Else
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
#End If
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
        MsgBox fejlm, vbOKOnly, Sprog.Error
#Else
        fejlm = fejlm & vbCrLf & vbCrLf & omax.KommentarOutput & vbCrLf & MaxProc.LastMaximaOutput
        MsgBox fejlm, vbOKOnly, Sprog.Error
        RestartMaxima
#End If
    End If
#If Mac Then
#Else
        Oundo.EndCustomRecord
#End If

    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
#If Mac Then
    d.Activate
#End If
    On Error Resume Next
    ActiveWindow.VerticalPercentScrolled = scrollpos
    Application.ScreenUpdating = True
    '    UnLockWindow
    '    TimeText = TimeText & vbCrLf & "beregn ialt: " & Timer - st
    '    MsgBox TimeText

End Sub

Function GetCmdAfterEqualSign(Kommando As String) As String
    Dim posligmed As Integer, possumtegn As Integer, posca As Integer
    Do    ' g*aa* tilbage til n*ae*rmeste ligmed
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
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    SaveSettingsCifre = MaximaCifre
    SaveSettingsExact = MaximaExact
    SaveSettingsOutunits = OutUnits
    SaveSettingsLog = MaximaLogOutput
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
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se p*aa* felter med := defligmed og := symbol
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
        If MaximaVidNotation Then
            s = "ScientificText(" & s & " , " & MaximaCifre & ")"
        ElseIf MaximaExact = 2 Then
            s = "numeric(" & s & " , " & MaximaCifre & ")"
        End If
    End If
    
    If CASengine = 0 Then
        omax.Omskriv UFomskriv.CheckBox_vidnotation.Value, UFomskriv.CheckBox_auto.Value, UFomskriv.CheckBox_factor.Value, UFomskriv.CheckBox_expand.Value, UFomskriv.CheckBox_rationaliser.Value, UFomskriv.CheckBox_trigreduce.Value
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
    If omax.CheckForError Then GoTo slut

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

    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    MaximaCifre = SaveSettingsCifre
    MaximaExact = SaveSettingsExact
    OutUnits = SaveSettingsOutunits
    MaximaLogOutput = SaveSettingsLog
    On Error Resume Next
    Selection.End = sslut    ' slut skal v*ae*re f*oe*rst eller g*aa*r det galt
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
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se p*aa* felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    If CASengine > 0 Then
        s = "simplify(" & omax.Kommando & ")"
        If MaximaVidNotation Then
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
    If omax.CheckForError Then GoTo slut
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

    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal v*ae*re f*oe*rst eller g*aa*r det galt
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
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se p*aa* felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    '    omax.OpenCmd
    omax.CompareTest
    If omax.StopNow Then GoTo slut
    If omax.CheckForError Then GoTo slut
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
    Selection.End = sslut    ' slut skal v*ae*re f*oe*rst eller g*aa*r det galt
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
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se p*aa* felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut
    
    
    If CASengine > 0 Then
        s = "factor(" & omax.Kommando & ")"
        If MaximaVidNotation Then
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
    If omax.CheckForError Then GoTo slut
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

    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal v*ae*re f*oe*rst eller g*aa*r det galt
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
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se p*aa* felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    If CASengine > 0 Then
        s = "expand(" & omax.Kommando & ")"
        If MaximaVidNotation Then
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
    If omax.CheckForError Then GoTo slut
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
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal v*ae*re f*oe*rst eller g*aa*r det galt
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
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se p*aa* felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    variabel = InputBox(Sprog.A(51), Sprog.Variable, "x")
    If variabel = "" Then GoTo slut
    '    omax.OpenCmd
    
    If CASengine > 0 Then
        s = "derivative(" & omax.Kommando & " , " & variabel & ")"
        If MaximaVidNotation Then
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
    If omax.CheckForError Then GoTo slut
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

    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal v*ae*re f*oe*rst eller g*aa*r det galt
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
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se p*aa* felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    variabel = InputBox(Sprog.A(51), Sprog.Variable, "x")
    If variabel = "" Then GoTo slut
    '    omax.OpenCmd
    
    If CASengine > 0 Then
        s = "integral(" & omax.Kommando & " , " & variabel & ")"
        If MaximaVidNotation Then
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
    If omax.CheckForError Then GoTo slut
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

    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal v*ae*re f*oe*rst eller g*aa*r det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos

End Sub
Sub SolveDENumeric()
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    Dim variabel As String, i As Integer
    Dim xmin As String, xmax As String, xstep As String, DElist As String, varlist As String, guesslist As String
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
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se p*aa* felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    omax.FindVariable
    If InStr(omax.vars, "t") > 0 Then
        variabel = "t"
    ElseIf InStr(omax.vars, "x") > 0 Then
        variabel = "x"
    ElseIf UFdiffeq.TextBox_varx.text = vbNullString Then
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
    Selection.End = sslut    ' slut skal v*ae*re f*oe*rst eller g*aa*r det galt
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
    Dim var As String, i As Integer
    If s = vbNullString Then
        GetLHSvar = vbNullString
        Exit Function
    End If
    s = omax.ConvertToAscii(s)
    Arr = Split(s, "=")
    s = Arr(0)
    ea.text = s
    Do
        var = ea.GetNextVar()
        If InStr(DeVarList, var) <= 0 Or DeVarList = vbNullString Then
            DeVarList = DeVarList & var & ","
            GetLHSvar = var
            Exit Function
        End If
        i = i + 1
    Loop While var <> vbNullString And i < 10
End Function
Sub SolveDE()
    SolveDEpar
End Sub
Sub SolveDEpar(Optional funktion As String, Optional variabel As String)
    On Error GoTo fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim vars As String
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    Dim t As String
    Dim Arr As Variant
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
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se p*aa* felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    '    funktion = InputBox("Indtast funktion som udtrykket skal l*oe*ses for", "Variabel", "y")
    '    variabel = InputBox("Indtast variabel som funktionen afh*ae*nger af", "Variabel", "x")
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
            funktion = "f(" & variabel & ")"
        ElseIf InStr(omax.Kommando, "^' (") > 0 Then
            ea.text = omax.Kommando
            funktion = ea.GetPrevVar(InStr(omax.Kommando, "^' (")) & "(" & variabel & ")"
        ElseIf InStr(omax.Kommando, "g(") > 0 Then
            funktion = "g(" & variabel & ")"
        ElseIf InStr(omax.Kommando, "g^' (") > 0 Then
            funktion = "g(" & variabel & ")"
        ElseIf InStr(omax.Kommando, "N(") > 0 Then
            funktion = "N(" & variabel & ")"
        ElseIf InStr(omax.Kommando, "N") > 0 Then
            funktion = "N"
        ElseIf InStr(omax.Kommando, "y(" & variabel & ")") > 0 Then
            funktion = "y(" & variabel & ")"
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
        If MaximaVidNotation Then
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
            omax.MaximaOutput = Replace(omax.MaximaOutput, "}", " ") ' mellemrummet sikrer afstand s*aa* det bliver gange efter c_1
'            If Left(omax.MaximaOutput, 1) = "{" Then omax.MaximaOutput = Mid(omax.MaximaOutput, 2, Len(omax.MaximaOutput) - 2)
            omax.MaximaOutput = Replace(omax.MaximaOutput, ";", "    " & VBA.ChrW(8744) & "    ")    ' komma erstattes med eller
            Application.Activate
    End If
    
        
    If omax.StopNow Then GoTo slut
    If omax.CheckForError Then GoTo slut
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

    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal v*ae*re f*oe*rst ellers g*aa*r det galt
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

    comm = Replace(comm, "?merror(""Anumberwasfoundwhereavariablewasexpected-`solve'"")", "Fejl: Du har bedt om at l*oe*se en ligning for en variabel der allerede er defineret." & vbCrLf & "S*ae*t en 'slet def:' kommando ind f*oe*r ligningen.")

    TranslateReplaceOutput = comm
End Function
Sub InsertOutput(text As String)
' inds*ae*tter text efter den mathobj cursoren er i p*aa* sikker m*aa*de
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
    Selection.TypeText text    ' giver problemer med at =tegn kommer under br*oe*kstreg
    Selection.Move wdCharacter, -1
    If Selection.OMaths(1).Range.Font.Bold Then
        IsBold = True
        Selection.OMaths(1).Range.Font.Bold = False
    End If
    Selection.OMaths.BuildUp
    If IsBold Then Selection.OMaths(1).Range.Font.Bold = True
    omax.GoToEndOfSelectedMaths
End Sub
Sub unicodevals()
    Dim s As String
    Dim i As Integer
    Dim c As Range
    MsgBox Selection.text
    For Each c In Selection.Characters
        i = AscW(c)
        s = s & c & " - " & i & vbCrLf
    Next

    MsgBox s
End Sub

Sub unicodevals2()
    Dim text As String
    Dim i As Integer
    Dim j As Integer
    Dim s As String
    Selection.OMaths.Linearize
    Selection.OMaths(1).ConvertToNormalText
    text = Selection.text
    Selection.OMaths(1).ConvertToMathText
    Selection.OMaths(1).Range.Select
    Selection.OMaths.BuildUp

    For j = 1 To Len(text)
        i = AscW(Mid(text, j, 1))
        s = s & Mid(text, j, 1) & " - " & i & vbCrLf
    Next
    MsgBox s

End Sub
Sub UnicodeValsToString()
' laver alle Omaths i selection om til en streng der kan indsættes i VBA-kode. Bruges primært til testmodul
' Strengene indsættes efter selection i rækkefølge. Hver på ny linje
    Dim text As String
    Dim j As Integer
    Dim i As Integer
    Dim k As Integer, n As Integer
    Dim s As String
    Dim mo As OMath
    Dim Arr() As String
    Dim MoArr() As Variant

    n = Selection.OMaths.Count
    If n = 0 Then
        MsgBox "You must select an equation", vbOKOnly, "Error"
        Exit Sub
    End If
    ReDim Arr(n - 1)
    ReDim MoArr(n - 1)
'    Selection.OMaths.Linearize
    For k = 0 To n - 1
        Set MoArr(k) = Selection.OMaths(k + 1)
    Next
    For k = 0 To n - 1
        Set mo = MoArr(k)
        mo.Linearize
        mo.ConvertToNormalText
        Arr(k) = Trim(mo.Range.text)
        mo.ConvertToMathText
        mo.Range.Select
        mo.BuildUp
    Next
    Selection.Collapse wdCollapseEnd
    Selection.EndKey Unit:=wdLine

    For k = 0 To UBound(Arr)
        text = Arr(k)
        s = ""
        For j = 1 To Len(text)
            i = AscW(Mid(text, j, 1))
            If i > 200 Or i = 183 Then
                s = s & """ & VBA.ChrW(" & i & ") & """
            Else
                s = s & Mid(text, j, 1)
            End If
        Next
        If Left(s, 4) = """ & " Then
            s = right(s, Len(s) - 4)
        ElseIf Left(s, 1) <> """" Then
            s = """" & s
        End If
        If right(s, 4) = " & """ Then
            s = Left(s, Len(s) - 4)
        ElseIf right(s, 1) <> """" Then
            s = s & """"
        End If
        Selection.Collapse wdCollapseEnd
        Selection.TypeParagraph
        Selection.TypeText (s)
    Next

End Sub

#If Mac Then
#Else
Sub LukMaximaProcess()
'    MaxProc.CloseProcess
    MaxProc.ConsoleInterrupt
End Sub
#End If

Function ValidateInput(Expr) As Boolean
   Dim n As Integer
   
   ValidateInput = True
   ' validate brackets
   
   If GetCountOfChar(Expr, "(") <> GetCountOfChar(Expr, ")") Then
      MsgBox "The number of brackets do not match in" & vbCrLf & Expr, vbOKOnly, "Syntax error"
      ValidateInput = False
   End If
   
End Function

Private Function GetCountOfChar(ByVal ar_sText As String, ByVal a_sChar As String) As Integer
  Dim l_iIndex As Integer
  Dim l_iMax As Integer
  Dim l_iLen As Integer

  GetCountOfChar = 0
  l_iMax = Len(ar_sText)
  l_iLen = Len(a_sChar)
  For l_iIndex = 1 To l_iMax
    If (Mid(ar_sText, l_iIndex, l_iLen) = a_sChar) Then 'found occurrence
      GetCountOfChar = GetCountOfChar + 1
      If (l_iLen > 1) Then l_iIndex = l_iIndex + (l_iLen - 1) 'if matching more than 1 char, need to move more than one char ahead to continue searching
    End If
  Next l_iIndex
End Function

