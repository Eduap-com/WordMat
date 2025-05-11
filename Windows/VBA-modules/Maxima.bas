Attribute VB_Name = "Maxima"
Option Explicit
Public UFSelectVar As UserFormSelectVar
Public omax As CMaxima
Public tid As Double
Private DeVarList As String
Private TempCas As Integer

Public Function PrepareMaxima(Optional FindDefinitioner As Boolean = True) As Boolean
    On Error GoTo Fejl

    RunFirst
    SetMaxProc
    
    If DebugWM Then
        UserFormDebug.Label_time.Caption = ""
        tid = timer
    End If
    
    SaveBackup
    
    If omax Is Nothing Then
        SetMathAutoCorrect
        On Error Resume Next
        Application.Run macroname:="Popstart"
        Err.Clear
        On Error GoTo Fejl
        DoEvents
        Set omax = New CMaxima
    End If
    
#If Mac Then
#Else
    If Not MaxProc Is Nothing And DllConnType <= 1 Then ' on Windows Maxima must be started first, if using api. Skip if using wsh
        If Not MaxProc.IsMaximaStarted And CASengine = 0 Then
            MaxProc.Units = 0
            MaxProc.StartMaximaProcess
            WaitForMaximaUntil
            If MaxProc.ErrCode > 0 Then
                MsgBox Sprog.A(55) & vbCrLf & "ErrCode: " & MaxProc.ErrCode & vbCrLf & vbCrLf & MaxProc.LastMaximaOutput, vbOKOnly, Sprog.Error
                GoTo Fejl
            End If
        End If
        If MaximaUnits Then
            If MaxProcUnit Is Nothing Then
                DoEvents
                On Error Resume Next
                Set MaxProcUnit = GetMaxProc() 'CreateObject("MaximaProcessClass")
                If Err.Number <> 0 Then
                    MsgBox Sprog.A(54), vbOKOnly, Sprog.Error
                    GoTo slut
                End If
                On Error GoTo Fejl
                MaxProcUnit.Units = 1
                MaxProcUnit.OutUnits = omax.ConvertUnits(OutUnits)
                MaxProcUnit.StartMaximaProcess
                WaitForMaximaUnitUntil
            End If
        End If
    End If
#End If

finish:
    omax.ConvertLnLog = True ' other functions can change this setting. its reset.
    If FindDefinitioner Then
        If Not omax.PrepareNewCommand Then    ' resets and finds definitions
            GoTo Fejl
        End If
    End If
    On Error Resume Next
    PrepareMaxima = True
    GoTo slut
Fejl:
    On Error Resume Next
    PrepareMaxima = False
slut:
    On Error Resume Next
End Function

#If Mac Then
#Else
Function GetMaxProc() As Object
'    If DllConnType = 0 Then
        On Error Resume Next
        Set GetMaxProc = CreateObject("MaximaProcessClass")
'    ElseIf DllConnType = 1 Then
'        Set GetMaxProc = PGetMaxProc()
'    End If
    DoEvents
End Function

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
    Do While MaxProcUnit.Finished = 0 And MaxProcUnit.ErrCode = 0 And Not (omax.StopNow) And i < StopTime
        Wait (0.1)
        i = i + 1
    Loop
slut:
End Sub
#End If

Sub RestartMaxima()
    Set omax = New CMaxima
    
#If Mac Then
#Else
    If Not CASengine = 0 Then Exit Sub
    
    If Not MaxProc Is Nothing Then MaxProc.CloseProcess
    
    On Error Resume Next
    Wait 1 ' Old MaxProc must be closed, before new can be created
    Set MaxProc = GetMaxProc() 'CreateObject("MaximaProcessClass")

    If Err.Number <> 0 Then
        MsgBox Sprog.A(54), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    On Error GoTo Fejl
    If omax Is Nothing Then
        Set omax = New CMaxima
    End If
    MaxProc.StartMaximaProcess
    If Not MaxProcUnit Is Nothing Then
        MaxProcUnit.CloseProcess
        MaxProcUnit.StartMaximaProcess
    End If
    GoTo slut
Fejl:
     MsgBox Sprog.A(55) & vbCrLf & "ErrCode: " & MaxProc.ErrCode & vbCrLf & vbCrLf & MaxProc.LastMaximaOutput, vbOKOnly, Sprog.Error
slut:
#End If
End Sub
Sub MaximaCommand()
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long, DontGoBack As Boolean
    On Error GoTo Fejl
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled

    PrepareMaxima
    omax.prevspr = ""
    If Not omax.MaximaInstalled Then GoTo slut
    
    If Selection.OMaths.Count = 0 And Len(Selection.Range.Text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    ElseIf Selection.OMaths.Count > 0 Then
        omax.ReadSelection
    Else
        omax.Kommando = Selection.Range.Text
        DontGoBack = True
    End If
    
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If CASengine = 0 Then
        omax.ExecuteMaximaCommand
    ElseIf CASengine = 1 Then
        OpenGeoGebraWeb omax.Kommando, "CAS", True, True
        GoTo slut
    ElseIf CASengine = 2 Then
        omax.MaximaOutput = RunGeoGebraDirect(omax.Kommando)
    End If

    If CheckForError Then
        scrollpos = ActiveWindow.VerticalPercentScrolled
        sslut = Selection.End
        sstart = Selection.start
        GoTo slut
    End If

    omax.GoToEndOfSelectedMaths DontGoBack
    Selection.TypeParagraph
    omax.InsertMaximaOutput
    '   UFWait.Hide
    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' The end must come first or things will go wrong.
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos

End Sub
Sub MaximaSolveInequality(Optional variabel As String)
    On Error GoTo Fejl
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
    
    Set UFSelectVar = New UserFormSelectVar

    If Selection.OMaths.Count < 2 Then

        omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If
        eqs = False
        If InStr(omax.Kommando, VBA.ChrW(8804)) > 0 Or InStr(omax.Kommando, VBA.ChrW(8805)) > 0 Then
            omax.Kommando = Replace(omax.Kommando, VBA.ChrW(8804), "<")
            omax.Kommando = Replace(omax.Kommando, VBA.ChrW(8805), ">")
            eqs = True
        End If
        If Not omax.FindVariable Then GoTo slut
        If variabel = vbNullString Then
            UFSelectVar.Vars = omax.Vars
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
                    InsertForklaring Sprog.A(829) & " " & variabel & " " & Sprog.A(831), True
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
                ea.Text = omax.MaximaOutput
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
        
        If CheckForError Then
            scrollpos = ActiveWindow.VerticalPercentScrolled
            sslut = Selection.End
            sstart = Selection.start
            GoTo slut
        End If
        
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
            InsertForklaring vbTab & Sprog.A(56) & " " & variabel & " " & Sprog.A(57)
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

    Else    ' system of equations
        MsgBox2 Sprog.A(376), vbOKOnly, ""
    End If
    
    Oundo.EndCustomRecord

    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut    ' slut must be first
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos

End Sub
Sub MaximaSolve()
' This construction is necessary to be able to call MaximaSolve with parameters from Test
' You cannot attach shortcuts to subs that take parameters even if they are optional
    MaximaSolvePar
End Sub
Sub MaximaSolvePar(Optional variabel As String)
    Dim Arr As Variant, s As String, t As String, v As String
    Dim fejlm As String
    On Error GoTo Fejl
    Application.ScreenUpdating = False
    Dim IsSolved As Boolean
    Dim scrollpos As Double
    Dim ea As New ExpressionAnalyser, SaveKommando As String
    Dim sstart As Long, sslut As Long, p As Long, p2 As Long
    Dim SaveSettingsCifre As Integer
    Dim SaveSettingsExact As Integer
    Dim SaveSettingsOutunits As String
    Dim SaveSettingsLog As Integer
    Dim SaveSettingsDecOutType As Integer
    
    scrollpos = ActiveWindow.VerticalPercentScrolled
    
    TempCas = CASengine
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

    If omax.DefFejl Then GoTo slut

    If Selection.OMaths.Count = 0 And Len(Selection.Range.Text) < 2 Then
        MsgBox2 Sprog.A(803), vbOKOnly, Sprog.A(804)
        GoTo slut
    End If
    If sstart = sslut Then
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    If InStr(Selection.OMaths(1).Range.Text, "<") > 1 Or InStr(Selection.OMaths(1).Range.Text, ">") > 1 Or InStr(Selection.OMaths(1).Range.Text, VBA.ChrW(8804)) > 1 Or InStr(Selection.OMaths(1).Range.Text, VBA.ChrW(8805)) > 1 Then
        MaximaSolveInequality variabel
        GoTo slut
    End If
    p = InStr(Selection.OMaths(1).Range.Text, "=")
    If p < 1 Then
        MsgBox Sprog.A(141), vbOKOnly, Sprog.Error
        GoTo slut
    Else
        p2 = InStr(p + 1, Selection.OMaths(1).Range.Text, "=")
        If p2 > 0 Then
            p = InStr(Selection.OMaths(1).Range.Text, ChrW(8743))
            If p < 1 Then
                MsgBox Sprog.A(882), vbOKOnly, Sprog.Error
                GoTo slut
            End If
        End If
    End If

    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    If Selection.OMaths.Count < 2 And InStr(Selection.OMaths(1).Range.Text, VBA.ChrW(8743)) < 1 Then
        ' only 1 equation
        
        If Not omax.FindVariable Then GoTo slut
        If Not ValiderVariable Then GoTo slut
        SaveKommando = omax.Kommando
newcas:
        omax.StopNow = False
        omax.Kommando = SaveKommando
        If CASengine > 0 And Not AllTrig Then ' In GeoGebra, it must be recognized via VBA whether it is a trigonometric equation
            If Not InStr(omax.Vars, ";") > 0 Then ' the method only works with 1 variable
                ea.SetNormalBrackets
                ea.Text = omax.Kommando
                ea.Text = Replace(ea.Text, VBA.ChrW(8289), "")
                s = ""
                Do
                    v = ea.GetNextVar()
                    If v = "sin" Or v = "cos" Or v = "tan" Then
                        t = ea.GetNextBracketContent()
                        If InStr(t, omax.Vars) > 0 Then
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
                        s = ChrW(960) & "/2" ' pi
                    Else
                        s = ChrW(960) '"pi"
                    End If
                End If
                If s <> "" Then
                    UFSelectVar.TextBox_def.Text = "0<=" & omax.Vars & "<=" & s & VbCrLfMac
                    UFSelectVar.TempDefs = "0<=" & omax.Vars & "<=" & s
                End If
            End If
        End If
        If variabel = vbNullString Then
            UFSelectVar.Vars = omax.Vars
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
                    InsertForklaring Sprog.A(829) & " " & variabel & " " & Sprog.A(831), True
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
                ea.Text = omax.MaximaOutput
                omax.MaximaOutput = ""
                ea.Pos = 1
                Do
                    s = Trim(ea.GetNextListItem(ea.Pos, ";"))
                    If s <> "" Then omax.MaximaOutput = omax.MaximaOutput & s & "    " & VBA.ChrW(8744) & "    "
                Loop While s <> ""
                If omax.MaximaOutput <> "" Then omax.MaximaOutput = Left(omax.MaximaOutput, Len(omax.MaximaOutput) - 9)
                '                omax.MaximaOutput = Replace(omax.MaximaOutput, ";", "    " & VBA.ChrW(8744) & "    ")    ' Comma replaced by or
            End If
            Application.Activate
        End If
                
        
        If omax.StopNow Then
            If omax.AntalVars > 1 Then
                GoTo slut
            End If
        End If
        If CheckForError Then
            scrollpos = ActiveWindow.VerticalPercentScrolled
            sslut = Selection.End
            sstart = Selection.start
            GoTo slut
        End If

        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
        
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph
        If Len(omax.MaximaOutput) > 250 Then
            Dim resultat As VbMsgBoxResult
            resultat = MsgBox(Sprog.A(127) & vbCrLf & vbCrLf & omax.MaximaOutput, vbOKCancel, Sprog.A(846))
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

        ' insert explanation if desired
        If MaximaForklaring And (IsSolved Or InStr(omax.KommentarOutput, "solving system of equations")) Then
            InsertForklaring Sprog.A(829) & " " & variabel & " " & Sprog.A(831)
        End If

        If InStr(omax.MaximaOutput, VBA.ChrW(8709)) Then    ' no solution
            omax.InsertMaximaOutput
            Selection.TypeParagraph
            If MaximaComplex Then
                Selection.TypeText Sprog.A(129) & variabel
            Else
                Selection.TypeText Sprog.A(130) & variabel
            End If
        ElseIf InStr(omax.MaximaOutput, VBA.ChrW(8477)) > 0 Then  ' all are solution
            omax.InsertMaximaOutput
            Selection.TypeParagraph
            Selection.TypeText Sprog.A(132)
        ElseIf omax.MaximaOutput = "?" Or omax.MaximaOutput = "" Or InStr(omax.KommentarOutput, "Lisp error") > 0 Or (Not LmSet And Not IsSolved) Then
            If CASengine = 0 Then
                GoTo stophop
            End If
            UserFormChooseCAS.Show
            If UserFormChooseCAS.ChosenCAS = 2 Then ' maxima num
                CASengineTempOnly = 0
                GoTo stophop
            ElseIf UserFormChooseCAS.ChosenCAS = -1 Then
                GoTo slut
            ElseIf UserFormChooseCAS.ChosenCAS = 1 Then ' Maxima sym
                CASengineTempOnly = 0
                GoTo newcas
            ElseIf UserFormChooseCAS.ChosenCAS = 3 Then ' geogebra sym
                CASengineTempOnly = 2
                GoTo newcas
            ElseIf UserFormChooseCAS.ChosenCAS = 4 Then ' geogebra num
                CASengineTempOnly = 2
                Selection.MoveLeft wdCharacter, 1
                MaximaNsolve variabel
                GoTo slut
            ElseIf UserFormChooseCAS.ChosenCAS = 5 Then ' geogebra browser sym
                CASengineTempOnly = 1
                GoTo newcas
            Else ' grafisk geogebra
                CASengineTempOnly = 1
                Selection.MoveLeft wdCharacter, 1
                MaximaNsolve variabel
                GoTo slut
            End If
        ElseIf False Then
stophop:
            omax.Nsolve variabel, -15, 15, 15, 20, 30, 30, True
            InsertForklaring Sprog.A(830) & " " & variabel & " " & Sprog.A(831), False
            omax.InsertMaximaOutput
            Selection.TypeParagraph
        Else    ' if there is a solution
            omax.InsertMaximaOutput
        End If

        Oundo.EndCustomRecord

    Else    '--------------- system of equations ----------------------

        omax.FindVariable
        If Not ValiderVariable Then GoTo slut
        UFSelectVar.NoEq = omax.AntalKom
        UFSelectVar.Vars = omax.Vars
        UFSelectVar.DefS = omax.DefString
        UFSelectVar.Show
        variabel = UFSelectVar.SelectedVar

        If variabel = "" Then GoTo slut
        omax.TempDefs = UFSelectVar.TempDefs
        
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
                omax.ConvertOutputToSolvedGGB ' should perhaps be improved with ea.getnextlistitem like solve
            End If
            Application.Activate
        End If
        
        If omax.StopNow Then GoTo slut
        If CheckForError Then
            scrollpos = ActiveWindow.VerticalPercentScrolled
            sslut = Selection.End
            sstart = Selection.start
            GoTo slut
        End If
        If omax.StopNow Then GoTo slut
        Application.ScreenUpdating = False
        
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
        
        If omax.MaximaOutput = "?merror(""Anumberwasfoundwhereavariablewasexpected-`solve'"")" Then
            MsgBox Sprog.A(133), vbOKOnly, Sprog.Error
            GoTo slut
        End If

        If omax.DefFejl Then GoTo slut

        variabel = Replace(omax.ConvertToWordSymbols(variabel), ";", ",")
        
        If omax.MaximaOutput = "?" Or omax.MaximaOutput = "" Or InStr(omax.KommentarOutput, "Lisp error") > 0 Then
            UserFormChooseCAS.Show
            If UserFormChooseCAS.ChosenCAS = 2 Then ' maxima num
                GoTo stophop
            ElseIf UserFormChooseCAS.ChosenCAS = -1 Then
                GoTo slut
            ElseIf UserFormChooseCAS.ChosenCAS = 1 Then ' Maxima sym
                CASengineTempOnly = 0
                GoTo newcassys
            ElseIf UserFormChooseCAS.ChosenCAS = 3 Then ' geogebra sym
                CASengineTempOnly = 2
                GoTo newcassys
            ElseIf UserFormChooseCAS.ChosenCAS = 4 Then ' geogebra num
                TempCas = CASengine
                CASengineTempOnly = 2
                '                Selection.MoveLeft wdCharacter, 1
                MaximaNsolve variabel
                GoTo slut
                CASengineTempOnly = TempCas
            ElseIf UserFormChooseCAS.ChosenCAS = 5 Then ' geogebra browser sym
                CASengineTempOnly = 1
                GoTo newcassys
            Else ' grafisk geogebra
                TempCas = CASengine
                CASengineTempOnly = 1
                Selection.MoveLeft wdCharacter, 1
                MaximaNsolve variabel
                CASengineTempOnly = TempCas
                GoTo slut
            End If
        ElseIf Len(omax.MaximaOutput) > 1 Then
            omax.GoToEndOfSelectedMaths
            Selection.TypeParagraph
            If MaximaForklaring Then
                InsertForklaring Sprog.A(134) & variabel & Sprog.A(135)
            End If
            omax.InsertMaximaOutput
            If InStr(omax.MaximaOutput, VBA.ChrW(8709)) > 0 And CASengine > 0 Then
                Selection.TypeParagraph
                Selection.TypeText Sprog.A(873)
            End If
        Else    ' no solutions
            omax.GoToEndOfSelectedMaths
            Selection.TypeParagraph
            If Len(omax.KommentarOutput) <= 1 Then
                omax.MaximaOutput = "L=" & VBA.ChrW(8709)
                omax.InsertMaximaOutput
                Selection.TypeParagraph
                If MaximaComplex Then
                    Selection.TypeText Sprog.A(137) & variabel
                Else
                    Selection.TypeText Sprog.A(138) & variabel
                End If
            Else
                If omax.DefFejl Then GoTo slut
                fejlm = Sprog.A(131) & vbCrLf
                '                End If
                If InStr(omax.Kommando, "=") < 1 Then
                    fejlm = fejlm & Sprog.A(139) & vbCrLf
                End If
                Selection.End = sslut  ' end must be first
                Selection.start = sstart
                ActiveWindow.VerticalPercentScrolled = scrollpos
            End If
        End If
        Oundo.EndCustomRecord
    End If
    GoTo slut
Fejl:
    MsgBox2 Sprog.ErrorGeneral & vbCrLf & "Err. no: " & Err.Number & vbCrLf & Err.Description & vbCrLf & "Line number: " & Erl, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    MaximaCifre = SaveSettingsCifre
    MaximaExact = SaveSettingsExact
    OutUnits = SaveSettingsOutunits
    MaximaLogOutput = SaveSettingsLog
    MaximaDecOutType = SaveSettingsDecOutType
    CASengineTempOnly = TempCas
    Selection.End = sslut    ' slut must be first
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
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
        Selection.TypeText Text:=vbTab
        Selection.OMaths.Add Range:=Selection.Range
        Selection.InsertSymbol CharacterNumber:=8661, Unicode:=True, Bias:=0
        Selection.OMaths(1).Range.Font.Size = 16
        Selection.MoveRight unit:=wdCharacter, Count:=1
        Selection.TypeText Text:=vbTab
    End If
    Selection.Font.Size = 8
    Selection.Font.ColorIndex = wdGray50
    Selection.Font.Italic = True

    Selection.TypeText ForklarTekst
    If omax.TempDefs <> "" Then
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
    On Error GoTo Fejl
    Application.ScreenUpdating = False
    Dim scrollpos As Double
    scrollpos = ActiveWindow.VerticalPercentScrolled

    PrepareMaxima
    omax.prevspr = ""
    Set UFSelectVar = New UserFormSelectVar
    Dim variabel As String
    Dim sstart As Long, sslut As Long
    If Not omax.MaximaInstalled Then GoTo slut
    sstart = Selection.start
    sslut = Selection.End

    If omax.DefFejl Then GoTo slut

    If Selection.OMaths.Count = 0 And Len(Selection.Range.Text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    If sstart = sslut Then
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    If InStr(Selection.OMaths(1).Range.Text, "=") < 1 Then
        Dim result As VbMsgBoxResult
        result = MsgBox(Sprog.A(141), vbYesNo, Sprog.A(846))
        If result = vbNo Then GoTo slut
    End If

    If Selection.OMaths.Count < 2 And InStr(Selection.OMaths(1).Range.Text, VBA.ChrW(8743)) < 1 Then
        GoTo slut
    Else    ' system of equations

        omax.ReadSelection
        omax.FindVariable
        If Not ValiderVariable Then GoTo slut
        UFSelectVar.Eliminate = True
        UFSelectVar.NoEq = omax.AntalKom
        UFSelectVar.Vars = omax.Vars
        UFSelectVar.DefS = omax.DefString
        UFSelectVar.Caption = Sprog.A(377)
        UFSelectVar.Show
        variabel = UFSelectVar.SelectedVar
        If variabel = "" Then GoTo slut
        omax.TempDefs = UFSelectVar.TempDefs
        
        If CASengine = 0 Then
            omax.Eliminate (variabel)
        Else
            MsgBox Sprog.A(874), vbOKOnly, "No support"
            GoTo slut
        End If
        
        If omax.StopNow Then GoTo slut
        If CheckForError Then
            scrollpos = ActiveWindow.VerticalPercentScrolled
            sslut = Selection.End
            sstart = Selection.start
            GoTo slut
        End If
        
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
        
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph
        If omax.StopNow Then GoTo slut
        Application.ScreenUpdating = False

        If omax.MaximaOutput = "?merror(""Anumberwasfoundwhereavariablewasexpected-`solve'"")" Then
            MsgBox Sprog.A(133), vbOKOnly, Sprog.Error
            GoTo slut
        End If

        If omax.DefFejl Then GoTo slut

        variabel = Replace(omax.ConvertToWordSymbols(variabel), ";", ",")

        If MaximaForklaring Then
            ForklarTekst = Sprog.A(142) & " " & variabel & " " & Sprog.A(143)
            InsertForklaring ForklarTekst, False
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
        Else    ' no solutions
            omax.MaximaOutput = "L=" & VBA.ChrW(8709)
            omax.InsertMaximaOutput
            Selection.TypeParagraph
            If MaximaComplex Then
                Selection.TypeText Sprog.A(137) & variabel
            Else
                Selection.TypeText Sprog.A(138) & variabel
            End If
        End If
    End If
    
    Oundo.EndCustomRecord

    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub
Sub MaximaNsolve(Optional ByVal variabel As String)
    Dim Arr As Variant
    Dim fejlm As String
    On Error GoTo Fejl
    Application.ScreenUpdating = False
    Dim IsSolved As Boolean
    Dim scrollpos As Double
    Dim ea As New ExpressionAnalyser, s As String, v As String, t As String
    scrollpos = ActiveWindow.VerticalPercentScrolled

    PrepareMaxima
    omax.prevspr = ""
    Set UFSelectVar = New UserFormSelectVar
    UFSelectVar.NoEq = 1
    Dim sstart As Long, sslut As Long
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo slut
    sstart = Selection.start
    sslut = Selection.End

    If omax.DefFejl Then GoTo slut

    If Selection.OMaths.Count = 0 And Len(Selection.Range.Text) < 2 Then
        MsgBox Sprog.A(803), vbOKOnly, Sprog.A(804)
        GoTo slut
    End If
    If sstart = sslut Then
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    If InStr(Selection.OMaths(1).Range.Text, "<") > 1 Or InStr(Selection.OMaths(1).Range.Text, ">") > 1 Or InStr(Selection.OMaths(1).Range.Text, VBA.ChrW(8804)) > 1 Or InStr(Selection.OMaths(1).Range.Text, VBA.ChrW(8805)) > 1 Then
        MaximaSolveInequality
        GoTo slut
    End If
    If InStr(Selection.OMaths(1).Range.Text, "=") < 1 Then
        Dim result As VbMsgBoxResult
        result = MsgBox(Sprog.A(141), vbYesNo, Sprog.A(846))
        If result = vbNo Then GoTo slut
    End If

    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Selection.OMaths.Count < 2 And InStr(Selection.OMaths(1).Range.Text, VBA.ChrW(8743)) < 1 Then
        ' only 1 equation

        If Not omax.FindVariable Then GoTo slut
        
        If CASengine > 0 And Not AllTrig Then ' In GeoGebra, it must be recognized via VBA whether it is a trigonometric equation
            If Not InStr(omax.Vars, ";") > 0 Then ' the method only works with 1 variable
                ea.SetNormalBrackets
                ea.Text = omax.Kommando
                ea.Text = Replace(ea.Text, VBA.ChrW(8289), "")
                s = ""
                Do
                    v = ea.GetNextVar()
                    If v = "sin" Or v = "cos" Or v = "tan" Then
                        t = ea.GetNextBracketContent()
                        If InStr(t, omax.Vars) > 0 Then s = "90"
                    End If
                    ea.Pos = ea.Pos + 1
                Loop While v <> ""
                If s <> "" And Radians Then s = "pi/2"
                If s <> "" Then
                    UFSelectVar.TextBox_def.Text = "0<=" & omax.Vars & "<=" & s & VbCrLfMac
                    UFSelectVar.TempDefs = "0<=" & omax.Vars & "<=" & s
                End If
            End If
        End If
        If variabel = vbNullString Then
            UFSelectVar.Vars = omax.Vars
            UFSelectVar.DefS = omax.DefString
            UFSelectVar.Show
            variabel = UFSelectVar.SelectedVar
        End If
        
        If variabel = "" Then GoTo slut
        omax.TempDefs = UFSelectVar.TempDefs
        
        Dim LHS As String, RHS As String
        If CASengine = 1 Or CASengine = 2 Then
            s = Trim(omax.Kommando)
            s = Replace(s, vbCrLf, "")
            s = Replace(s, vbCr, "")
            s = Replace(s, vbLf, "")
            s = "nsolve(" & Replace(s, ",", ".") & "," & variabel & ")"
        End If
        
        If CASengine = 0 Then
            omax.Nsolve variabel, -15, 15, 15, 20, 30, 30, True
        ElseIf CASengine = 1 Then
            If MaximaForklaring Then
                omax.GoToEndOfSelectedMaths
                Selection.TypeParagraph
                InsertForklaring Sprog.A(875), False
                Selection.TypeParagraph
            End If
            s = Replace(omax.Kommando, ",", ".")
            Arr = Split(s, "=")
            LHS = Arr(0)
            RHS = Arr(1)
            If variabel <> "x" Then
                ea.Text = LHS
                ea.ReplaceVar variabel, "x"
                LHS = ea.Text
                ea.Text = RHS
                ea.ReplaceVar variabel, "x"
                RHS = ea.Text
            End If
            OpenGeoGebraWeb "y=" & LHS & ";y=" & RHS & ";intersect(" & LHS & "," & RHS & ");" & "Nsolve(" & s & "," & variabel & ")", "CAS", True, True
            GoTo slut
        ElseIf CASengine = 2 Then
            If MaximaDecOutType = 3 Then
                s = "ScientificText(" & s & " , " & MaximaCifre & ")"
            Else
                s = "Numeric(" & s & " , " & MaximaCifre & ")"
            End If
            Call RunGeoGebraDirect(s)
            If Left(omax.MaximaOutput, 1) = "{" Then omax.MaximaOutput = Mid(omax.MaximaOutput, 2, Len(omax.MaximaOutput) - 2)
            omax.MaximaOutput = Replace(omax.MaximaOutput, ";", "    " & VBA.ChrW(8744) & "    ")    ' comma replaced by or
            Application.Activate
        End If

        If omax.StopNow Then GoTo slut
        If CheckForError Then
            scrollpos = ActiveWindow.VerticalPercentScrolled
            sslut = Selection.End
            sstart = Selection.start
            GoTo slut
        End If
        
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord

        Selection.End = sslut
        Selection.start = sstart
        ActiveWindow.VerticalPercentScrolled = scrollpos
        If CASengine = 0 Then
            GoTo ghop
        Else
            s = Replace(omax.Kommando, ",", ".")
            Arr = Split(s, "=")
            LHS = Arr(0)
            RHS = Arr(1)
            If variabel <> "x" Then
                ea.Text = LHS
                ea.ReplaceVar variabel, "x"
                LHS = ea.Text
                ea.Text = RHS
                ea.ReplaceVar variabel, "x"
                RHS = ea.Text
            End If
            OpenGeoGebraWeb "y=" & LHS & ";y=" & RHS & ";intersect(" & LHS & "," & RHS & ");" & "Nsolve(" & s & "," & variabel & ")", "CAS", True, True
        End If
        GoTo slut
ghop:
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph

        InsertForklaring Sprog.A(830) & " " & variabel & " " & Sprog.A(57), False

        If Len(omax.MaximaOutput) > 150 Then
            Dim resultat As VbMsgBoxResult
            resultat = MsgBox(Sprog.A(127) & omax.MaximaOutput, vbOKCancel, Sprog.A(846))
            If resultat = vbCancel Then GoTo slut
        End If

        variabel = omax.ConvertToWordSymbols(variabel)

        If omax.IsAllSolved(omax.MaximaOutput, variabel, VBA.ChrW(8744)) = "false" And Not (InStr(variabel, "+") > 0) Then
            IsSolved = False
        Else
            IsSolved = True
        End If

        If MaximaForklaring And IsSolved Then
            InsertForklaring Sprog.A(830) & " " & variabel & " " & Sprog.A(57)
        End If

        omax.InsertMaximaOutput
    Else    ' system of equations
        Dim Variable As String, guess As String, inp As String, j As Integer
        Dim Arr2 As Variant

        omax.ReadSelection
        omax.FindVariable
        UFSelectVar.NoEq = Selection.OMaths.Count
        UFSelectVar.Vars = omax.Vars
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
                InsertForklaring Sprog.A(883), False
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
            omax.MaximaOutput = Replace(omax.MaximaOutput, ";", "    " & VBA.ChrW(8744) & "    ")    ' comma replaced by or
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
            If omax.DefFejl Then GoTo slut
            fejlm = Sprog.A(383) & "." & vbCrLf
            If InStr(omax.Kommando, "=") < 1 Then
                fejlm = fejlm & Sprog.A(139) & "." & vbCrLf
            End If
            MsgBox fejlm & Sprog.A(384) & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
        End If
    End If
    
    Oundo.EndCustomRecord

    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub

Sub beregn()
    Dim fejlm As String, RemoveEqual As Boolean
        On Error GoTo Fejl
   ' Application.ScreenUpdating = False
    Dim tid As Single
    tid = timer
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
    
    RunFirst
    
    If SettUseVBACAS And (MaximaExact = 2 And Not MaximaUnits And Not MaximaComplex) Then
        On Error Resume Next
        Err.Clear
        Application.Run macroname:="CASCALC"
        If Err.Number = 513 Then
            MsgBox2 Err.Description, vbOKOnly, Sprog.Error
            GoTo slut
        ElseIf Err.Number = 0 Then ' if no errors, then the calculation is done and inserted into Word
            IncreaseCalcCounter
            GoTo slut
        End If
        Err.Clear
        On Error GoTo Fejl
    End If
    
    If Not PrepareMaxima Then
        If omax.DefFejl Then
            Exit Sub
        Else
            GoTo slut
        End If
    End If
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
    
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    
    If right(omax.Kommando, 1) = "=" Then
        omax.Kommando = Trim(Left(omax.Kommando, Len(omax.Kommando) - 1))
        RemoveEqual = True
    End If
    
    If omax.Kommando = "" Then GoTo slut
    
    If Not ValidateInput(omax.Kommando) Then
        If omax.DefFejl Then
            Exit Sub
        Else
            GoTo slut
        End If
    End If
    
    Dim s As String, t As String, fo As String
    
    omax.Kommando = GetCmdAfterEqualSign(Trim(omax.Kommando))
    If CASengine > 0 Then
        s = omax.Kommando
        If MaximaDecOutType = 3 Then
            s = "ScientificText(" & s & " , " & MaximaCifre & ")"
        ElseIf MaximaExact = 2 Then
            s = "numeric(" & s & " , " & MaximaCifre & ")"
        End If
    End If
    
    If CASengine = 0 Then
        omax.beregn
        If RemoveEqual Then
            If Left(omax.MaximaOutput, 1) = "=" Then
                omax.MaximaOutput = right(omax.MaximaOutput, Len(omax.MaximaOutput) - 1)
            End If
        End If
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
                ' the first result cannot simply be fed back into GeoGebra. This causes problems in special cases. Example: '\cbrt(79/138)^(2)' Here there should be parentheses. The correct interpretation is cbrt((79/138)^2), which appears if it is typed and read in Word. If it is run directly in WordMat the translator does not insert the correct parentheses. Normally not a problem, as everything is normally read from Word
                '                fo = omax.ReadFromWord(fo) ' attempt to work around the problem of feeding results directly into geogebra again
                '                s = "numeric(" & fo & " , " & MaximaCifre & ")"
                s = "numeric(" & s & " , " & MaximaCifre & ")" ' There are examples where it is better to calculate numerically further on the exact result instead of directly on the original, but due to the above problems
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
        If Not RemoveEqual Then omax.MaximaOutput = "=" & omax.MaximaOutput
    End If
    
    If omax.StopNow Then GoTo slut
    If CheckForError Then Exit Sub ' do not scroll down if there is an error, deffejl is selected
    '    TimeText = TimeText & vbCrLf & "beregn: " & Timer - st

    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord
    
    If Len(omax.MaximaOutput) > 0 And Trim(omax.MaximaOutput) <> "=res1" Then
        InsertOutput omax.MaximaOutput, False
    ElseIf omax.KommentarOutput <> vbNullString Then
        fejlm = Sprog.A(128) & vbCrLf
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
Fejl:
    MsgBox Sprog.ErrorGeneral & vbCrLf & "Err. no: " & Err.Number & vbCrLf & Err.Description & vbCrLf & "Line number: " & Erl, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
#If Mac Then
    D.Activate
#End If
    On Error Resume Next
    If ActiveWindow.VerticalPercentScrolled <> scrollpos Then ActiveWindow.VerticalPercentScrolled = scrollpos
    Application.ScreenUpdating = True
    '    TimeText = TimeText & vbCrLf & "beregn ialt: " & Timer - st
    '    MsgBox TimeText
    '    MsgBox Timer - tid
End Sub

Function GetCmdAfterEqualSign(Kommando As String) As String
' This function returns the expression to the right of the rightmost equal sign in the expression. If the equal sign is in a bracket, it is ignored
    Dim posligmed As Integer, possumtegn As Integer, posca As Integer, i As Integer, c As String
    
    posligmed = InStrRev(Kommando, "=") ' go back to the nearest equal
    If posligmed > 0 Then ' only cut the string if it has at least one equal sign
        Dim BracketStack As New Collection
        For i = Len(Kommando) To 1 Step -1 ' loop through all characters in the string
            c = Mid$(Kommando, i, 1)
            If (c = "=" Or c = ChrW(8776)) And BracketStack.Count = 0 Then ' we found the equal sign, now cut and stop
                GetCmdAfterEqualSign = right(Kommando, Len(Kommando) - i)
                Exit Function
            ElseIf c = ")" Or c = "]" Or c = "}" Then
                BracketStack.Add c
            ElseIf c = "(" Then
                If BracketStack.Count = 0 Then ' Bracketproblem stop and return full string
                    GoTo TheEnd
                ElseIf BracketStack.Item(BracketStack.Count) <> ")" Then ' brackets dont match. Stop and return full string
                    GoTo TheEnd
                Else ' pop the bracket from the stack
                    BracketStack.Remove (BracketStack.Count)
                End If
            ElseIf c = "[" Then
                If BracketStack.Count = 0 Then ' Bracketproblem stop and return full string
                    GoTo TheEnd
                ElseIf BracketStack.Item(BracketStack.Count) <> "]" Then ' brackets dont match. Stop and return full string
                    GoTo TheEnd
                Else ' pop the bracket from the stack
                    BracketStack.Remove (BracketStack.Count)
                End If
            ElseIf c = "{" Then
                If BracketStack.Count = 0 Then ' Bracketproblem stop and return full string
                    GoTo TheEnd
                ElseIf BracketStack.Item(BracketStack.Count) <> "}" Then ' brackets dont match. Stop and return full string
                    GoTo TheEnd
                Else ' pop the bracket from the stack
                    BracketStack.Remove (BracketStack.Count)
                End If
            End If
        Next
    End If
TheEnd:
    GetCmdAfterEqualSign = Kommando
End Function

Sub Omskriv()
    On Error GoTo Fejl
    Dim s As String
    Dim UFomskriv As New UserFormReduce
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

    If Selection.OMaths.Count = 0 And Len(Selection.Range.Text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    If sstart = sslut Then    ' ensures that gotoendofselectedmaths works
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    If Not omax.FindVariable Then GoTo slut
    UFomskriv.Vars = omax.Vars
    UFomskriv.Show
    If UFomskriv.annuller Then GoTo slut

    omax.TempDefs = UFomskriv.TempDefs
    
    If CASengine > 0 Then
        s = omax.Kommando
        If UFomskriv.CheckBox_trigreduce.Value Then s = "trigsimplify(" & s & ")"
        If UFomskriv.CheckBox_factor.Value Then s = "factor(" & s & ")"
        If UFomskriv.CheckBox_expand.Value Then s = "expand(" & s & ")"
        If UFomskriv.CheckBox_auto.Value Then s = "simplify(" & s & ")"
        If UFomskriv.CheckBox_completesquare.Value Then s = "completesquare(" & s & ")"
        If Not UFomskriv.CheckBox_factor.Value Then ' numeric expands factorised expressions.
            If MaximaDecOutType = 3 Then
                s = "ScientificText(" & s & " , " & MaximaCifre & ")"
            ElseIf MaximaExact = 2 Then
                s = "numeric(" & s & " , " & MaximaCifre & ")"
            End If
        End If
    End If
    
    If CASengine = 0 Then
        omax.Omskriv False, UFomskriv.CheckBox_auto.Value, UFomskriv.CheckBox_factor.Value, UFomskriv.CheckBox_expand.Value, UFomskriv.CheckBox_rationaliser.Value, UFomskriv.CheckBox_trigreduce.Value, UFomskriv.CheckBox_completesquare.Value
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
        If CheckForError Then
            scrollpos = ActiveWindow.VerticalPercentScrolled
            sslut = Selection.End
            sstart = Selection.start
            GoTo slut
        End If

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
                    s = s & Sprog.A(807) & ", "
                End If
                If UFomskriv.CheckBox_factor.Value Then
                    s = s & Sprog.A(806) & ", "
                End If
                If UFomskriv.CheckBox_rationaliser.Value Then
                    s = s & Sprog.A(794) & ", "
                End If
                If UFomskriv.CheckBox_trigreduce.Value Then
                    s = s & Sprog.A(152)
                End If
                If UFomskriv.CheckBox_completesquare.Value Then
                    s = s & Sprog.A(697)
                End If
                InsertForklaring s, False
            End If
            omax.InsertMaximaOutput
        Else
            Dim fejlm As String
            fejlm = Sprog.A(153) & vbCrLf
            MsgBox fejlm & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
        End If
    Else    ' not newline
        InsertOutput "=" & omax.MaximaOutput
    End If
    
    Oundo.EndCustomRecord

    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    MaximaCifre = SaveSettingsCifre
    MaximaExact = SaveSettingsExact
    OutUnits = SaveSettingsOutunits
    MaximaLogOutput = SaveSettingsLog
    MaximaDecOutType = SaveSettingsDecOutType
    On Error Resume Next
    Selection.End = sslut
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub
Sub reducer()
    On Error GoTo Fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim scrollpos As Double, s As String
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo slut
    '    MsgBox WordWindowNavn

    If Selection.OMaths.Count = 0 And Len(Selection.Range.Text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    If sstart = sslut Then    ' ensures that gotoendofselectedmaths works
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then
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
        If CheckForError Then
            scrollpos = ActiveWindow.VerticalPercentScrolled
            sslut = Selection.End
            sstart = Selection.start
            GoTo slut
        End If
    
    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord

    omax.GoToEndOfSelectedMaths
    Selection.TypeParagraph

    If Len(omax.MaximaOutput) > 0 Then
        If MaximaForklaring Then
            InsertForklaring Sprog.A(154), False
        End If
        omax.InsertMaximaOutput
    Else
        Dim fejlm As String
        fejlm = Sprog.A(153) & vbCrLf

        MsgBox fejlm & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
    End If

    Oundo.EndCustomRecord

    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub
Sub CompareTest()
    On Error GoTo Fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    If Not omax.MaximaInstalled Then GoTo slut

    If Selection.OMaths.Count = 0 And Len(Selection.Range.Text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    If sstart = sslut Then    ' ensures that gotoendofselectedmaths works
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    omax.CompareTest
    If omax.StopNow Then GoTo slut
    If CheckForError Then
        scrollpos = ActiveWindow.VerticalPercentScrolled
        sslut = Selection.End
        sstart = Selection.start
        GoTo slut
    End If
    omax.GoToEndOfSelectedMaths
    Selection.TypeParagraph
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
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub
Sub faktoriser()
    On Error GoTo Fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim scrollpos As Double, s As String
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    If Not omax.MaximaInstalled Then GoTo slut

    If Selection.OMaths.Count = 0 And Len(Selection.Range.Text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    If sstart = sslut Then    ' ensures that gotoendofselectedmaths works
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then
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

    If omax.StopNow Then GoTo slut
    If CheckForError Then
        scrollpos = ActiveWindow.VerticalPercentScrolled
        sslut = Selection.End
        sstart = Selection.start
        GoTo slut
    End If
    
    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord

    omax.GoToEndOfSelectedMaths
    Selection.TypeParagraph
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
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub
Sub udvid()
    On Error GoTo Fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    Dim scrollpos As Double, s As String
    scrollpos = ActiveWindow.VerticalPercentScrolled
    If Not omax.MaximaInstalled Then GoTo slut
    If sstart = sslut Then
        Selection.OMaths(1).ParentOMath.Range.Select
    End If

    If Selection.OMaths.Count = 0 And Len(Selection.Range.Text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' only look at fields with := defligmed and := symbol
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
    If CheckForError Then
        scrollpos = ActiveWindow.VerticalPercentScrolled
        sslut = Selection.End
        sstart = Selection.start
        GoTo slut
    End If
    
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
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub
Sub Differentier()
    On Error GoTo Fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim variabel As String, s As String
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    If Not omax.MaximaInstalled Then GoTo slut
    If sstart = sslut Then    ' ensures that gotoendofselectedmaths works
        Selection.OMaths(1).ParentOMath.Range.Select
    End If

    If Selection.OMaths.Count = 0 And Len(Selection.Range.Text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    variabel = InputBox(Sprog.A(51), Sprog.A(845), "x")
    If variabel = "" Then GoTo slut
    
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
    If CheckForError Then
        scrollpos = ActiveWindow.VerticalPercentScrolled
        sslut = Selection.End
        sstart = Selection.start
        GoTo slut
    End If
    
    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord
    
    omax.GoToEndOfSelectedMaths
    Selection.TypeParagraph
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
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub
Sub Integrer()
    On Error GoTo Fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim variabel As String, s As String
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo slut

    If Selection.OMaths.Count = 0 And Len(Selection.Range.Text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    If sstart = sslut Then    ' ensures that gotoendofselectedmaths works
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' only look at fields with := defligmed and := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    variabel = InputBox(Sprog.A(51), Sprog.A(845), "x")
    If variabel = "" Then GoTo slut
    
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
    If CheckForError Then
        scrollpos = ActiveWindow.VerticalPercentScrolled
        sslut = Selection.End
        sstart = Selection.start
        GoTo slut
    End If
    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord

    omax.GoToEndOfSelectedMaths
    Selection.TypeParagraph
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
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub
Sub SolveDENumeric()
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    Dim variabel As String
    Dim ea As New ExpressionAnalyser
    Dim UFdiffeq As New UserFormDeSolveNumeric
    On Error GoTo Fejl

    ea.SetNormalBrackets

    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled

    PrepareMaxima

    If Not omax.MaximaInstalled Then GoTo slut

    If Selection.OMaths.Count > 0 And sstart = sslut Then    ' ensures that gotoendofselectedmaths works
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' only look at fields with := defligmed and := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    omax.FindVariable
    If InStr(omax.Vars, "t") > 0 Then
        variabel = "t"
    ElseIf InStr(omax.Vars, "x") > 0 Then
        variabel = "x"
    Else
        variabel = "x"
    End If
    If omax.Kommando <> vbNullString Or omax.KommandoArrayLength > 0 Then
        UFdiffeq.TextBox_eq1.Text = GetRHS(omax.KommandoArray(0))
        UFdiffeq.TextBox_eq2.Text = GetRHS(omax.KommandoArray(1))
        UFdiffeq.TextBox_eq3.Text = GetRHS(omax.KommandoArray(2))
        UFdiffeq.TextBox_eq4.Text = GetRHS(omax.KommandoArray(3))
        UFdiffeq.TextBox_eq5.Text = GetRHS(omax.KommandoArray(4))
        UFdiffeq.TextBox_eq6.Text = GetRHS(omax.KommandoArray(5))
        UFdiffeq.TextBox_eq7.Text = GetRHS(omax.KommandoArray(6))
        UFdiffeq.TextBox_eq8.Text = GetRHS(omax.KommandoArray(7))
        UFdiffeq.TextBox_eq9.Text = GetRHS(omax.KommandoArray(8))
        DeVarList = ""
        UFdiffeq.TextBox_var1.Text = GetLHSvar(omax.KommandoArray(0))
        UFdiffeq.TextBox_var2.Text = GetLHSvar(omax.KommandoArray(1))
        UFdiffeq.TextBox_var3.Text = GetLHSvar(omax.KommandoArray(2))
        UFdiffeq.TextBox_var4.Text = GetLHSvar(omax.KommandoArray(3))
        UFdiffeq.TextBox_var5.Text = GetLHSvar(omax.KommandoArray(4))
        UFdiffeq.TextBox_var6.Text = GetLHSvar(omax.KommandoArray(5))
        UFdiffeq.TextBox_var7.Text = GetLHSvar(omax.KommandoArray(6))
        UFdiffeq.TextBox_var8.Text = GetLHSvar(omax.KommandoArray(7))
        UFdiffeq.TextBox_var9.Text = GetLHSvar(omax.KommandoArray(8))
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
    If UFdiffeq.TextBox_eq1.Text <> vbNullString And UFdiffeq.TextBox_init1.Text = vbNullString Then
        UFdiffeq.TextBox_init1.Text = "1"
    End If
    If UFdiffeq.TextBox_eq2.Text <> vbNullString And UFdiffeq.TextBox_init2.Text = vbNullString Then
        UFdiffeq.TextBox_init2.Text = "1"
    End If
    If UFdiffeq.TextBox_eq3.Text <> vbNullString And UFdiffeq.TextBox_init3.Text = vbNullString Then
        UFdiffeq.TextBox_init3.Text = "1"
    End If
    If UFdiffeq.TextBox_eq4.Text <> vbNullString And UFdiffeq.TextBox_init4.Text = vbNullString Then
        UFdiffeq.TextBox_init4.Text = "1"
    End If
    If UFdiffeq.TextBox_eq5.Text <> vbNullString And UFdiffeq.TextBox_init5.Text = vbNullString Then
        UFdiffeq.TextBox_init5.Text = "1"
    End If
    If UFdiffeq.TextBox_eq6.Text <> vbNullString And UFdiffeq.TextBox_init6.Text = vbNullString Then
        UFdiffeq.TextBox_init6.Text = "1"
    End If
    If UFdiffeq.TextBox_eq7.Text <> vbNullString And UFdiffeq.TextBox_init7.Text = vbNullString Then
        UFdiffeq.TextBox_init7.Text = "1"
    End If
    If UFdiffeq.TextBox_eq8.Text <> vbNullString And UFdiffeq.TextBox_init8.Text = vbNullString Then
        UFdiffeq.TextBox_init8.Text = "1"
    End If
    If UFdiffeq.TextBox_eq9.Text <> vbNullString And UFdiffeq.TextBox_init9.Text = vbNullString Then
        UFdiffeq.TextBox_init9.Text = "1"
    End If

    UFdiffeq.TextBox_varx.Text = variabel

    UFdiffeq.Show
    If UFdiffeq.luk Then GoTo slut

    If UFdiffeq.InsertType = 3 Then
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph
        omax.MaximaOutput = UFdiffeq.ListOutput
        omax.InsertMaximaOutput
    End If

    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    Selection.End = sslut
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
    ea.Text = s
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
    On Error GoTo Fejl
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
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo slut
    dAsDiffChrTemp = True
    
    If Selection.OMaths.Count = 0 And Len(Selection.Range.Text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    If sstart = sslut Then
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' only look at fields with := defligmed and := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo slut

    If funktion = vbNullString And variabel = vbNullString Then
        If Not omax.FindVariable Then GoTo slut
        If InStr(omax.Vars, "t") > 0 Then
            variabel = "t"
        ElseIf InStr(omax.Vars, "x") > 0 Then
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
        ElseIf InStr(omax.Vars, "y") > 0 Then
            funktion = "y"
        Else
            ea.Text = omax.Kommando
            funktion = ea.GetNextVar(1)
        End If
        UFdiffeq.Vars = omax.Vars
        UFdiffeq.DefS = omax.DefString
        UFdiffeq.TextBox_funktion.Text = funktion
        UFdiffeq.TextBox_variabel.Text = variabel
        UFdiffeq.Label_ligning.Caption = Replace(omax.ConvertToAscii(omax.Kommando), " ", "")
        UFdiffeq.Show
        If UFdiffeq.luk Then GoTo slut
        variabel = UFdiffeq.TextBox_variabel.Text
        funktion = UFdiffeq.TextBox_funktion.Text
    Else
        Dim Arr() As String
        Arr = Split(variabel, "=")
        If UBound(Arr) > 0 Then
            UFdiffeq.TextBox_startx.Text = Arr(1)
            variabel = Arr(0)
        End If
        Arr = Split(funktion, "=")
        If UBound(Arr) > 0 Then
            UFdiffeq.TextBox_starty.Text = Arr(1)
            funktion = Arr(0)
        End If
    End If
    If variabel = "" Then GoTo slut
    omax.TempDefs = UFdiffeq.TempDefs

    If CASengine > 0 Then
        Dim s As String
        s = Replace(omax.Kommando, "^'", "'")
        s = Replace(s, vbCrLf, "")
        s = Replace(s, vbCr, "")
        s = Replace(s, vbLf, "")
        s = Replace(s, ",", ".")
        t = "," & funktion & "," & variabel
        If UFdiffeq.TextBox_startx.Text <> "" And UFdiffeq.TextBox_starty.Text <> "" Then
            t = t & ",{(" & Replace(UFdiffeq.TextBox_startx.Text, ",", ".") & " , " & Replace(UFdiffeq.TextBox_starty.Text, ",", ".") & ")"
        End If
        If UFdiffeq.TextBox_bcx.Text <> "" And UFdiffeq.TextBox_bcy.Text <> "" Then
            t = t & ",(" & Replace(UFdiffeq.TextBox_bcx.Text, ",", ".") & " , " & Replace(UFdiffeq.TextBox_bcy.Text, ",", ".") & ")}"
        ElseIf UFdiffeq.TextBox_startx.Text <> "" And UFdiffeq.TextBox_starty.Text <> "" Then
            t = t & "}"
        End If
        If UFdiffeq.TextBox_starty2.Text <> "" Then
        t = t & ",(0," & Replace(UFdiffeq.TextBox_starty2.Text, ",", ".") & ")"
        End If
        s = "solveODE(" & s & t & ")"
        If MaximaDecOutType = 3 Then
            s = "ScientificText(" & s & " , " & MaximaCifre & ")"
        ElseIf MaximaExact = 2 Then
            s = "numeric(" & s & " , " & MaximaCifre & ")"
        End If
    End If

    If CASengine = 0 Then
        omax.SolveDE funktion, variabel, UFdiffeq.TextBox_startx.Text, UFdiffeq.TextBox_starty.Text, UFdiffeq.TextBox_starty2.Text, UFdiffeq.TextBox_bcx.Text, UFdiffeq.TextBox_bcy.Text
    ElseIf CASengine = 1 Then
        If MaximaForklaring Then
            omax.GoToEndOfSelectedMaths
            If MaximaForklaring Then
                Selection.TypeParagraph
                InsertForklaring Sprog.A(829) & " " & variabel & " " & Sprog.A(831), True
                Selection.TypeParagraph
            End If
        End If
        OpenGeoGebraWeb s, "CAS", True, True
        GoTo slut
    ElseIf CASengine = 2 Then
            Call RunGeoGebraDirect(s)
            omax.MaximaOutput = Replace(omax.MaximaOutput, "{", "")
            omax.MaximaOutput = Replace(omax.MaximaOutput, "}", " ") ' the space ensures distance so that it becomes times after c_1
'            If Left(omax.MaximaOutput, 1) = "{" Then omax.MaximaOutput = Mid(omax.MaximaOutput, 2, Len(omax.MaximaOutput) - 2)
            omax.MaximaOutput = Replace(omax.MaximaOutput, ";", "    " & VBA.ChrW(8744) & "    ")    ' comma replaced by or
            Application.Activate
    End If
    
    If omax.StopNow Then GoTo slut
    If CheckForError Then
        scrollpos = ActiveWindow.VerticalPercentScrolled
        sslut = Selection.End
        sstart = Selection.start
        GoTo slut
    End If
    
    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord
    
    omax.GoToEndOfSelectedMaths
    
    Selection.TypeParagraph
    If Len(omax.MaximaOutput) > 1 Then
        If MaximaForklaring Then
            funktion = Split(funktion, "(")(0)
            t = Sprog.A(390)
            If Len(UFdiffeq.TextBox_startx.Text) > 0 And Len(UFdiffeq.TextBox_starty.Text) > 0 And Len(UFdiffeq.TextBox_bcx.Text) > 0 And Len(UFdiffeq.TextBox_bcy.Text) > 0 Then
                t = t & " " & Sprog.A(391) & " " & funktion & "(" & UFdiffeq.TextBox_startx.Text & ")=" & UFdiffeq.TextBox_starty.Text & " og " & funktion & "(" & UFdiffeq.TextBox_starty.Text & ")=" & UFdiffeq.TextBox_bcy.Text
            ElseIf Len(UFdiffeq.TextBox_startx.Text) > 0 And Len(UFdiffeq.TextBox_starty.Text) > 0 And Len(UFdiffeq.TextBox_starty2.Text) > 0 Then
                t = t & " " & Sprog.A(392) & " " & funktion & "(" & UFdiffeq.TextBox_startx.Text & ")=" & UFdiffeq.TextBox_starty.Text & " og " & funktion & "'(" & UFdiffeq.TextBox_startx.Text & ")=" & UFdiffeq.TextBox_starty2.Text
            ElseIf Len(UFdiffeq.TextBox_startx.Text) > 0 And Len(UFdiffeq.TextBox_starty.Text) > 0 Then
                t = t & " " & Sprog.A(393) & " " & funktion & "(" & UFdiffeq.TextBox_startx.Text & ")=" & UFdiffeq.TextBox_starty.Text
            End If
            InsertForklaring t, False
        End If
        omax.InsertMaximaOutput
    Else
        Dim fejlm As String
        fejlm = Sprog.A(160) & vbCrLf

        MsgBox fejlm & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
    End If

    Oundo.EndCustomRecord

    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
slut:
    On Error Resume Next
    dAsDiffChrTemp = False
    Selection.End = sslut
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub

Sub InsertOutput(Text As String, Optional ResultAfterTable As Boolean = True)
' inserts text after the mathobj the cursor is in, safely
    Dim IsBold As Boolean
    If Selection.start = Selection.End Then
        omax.GoToEndOfSelectedMaths False, ResultAfterTable
    Else
        Selection.Collapse wdCollapseEnd
    End If
    If Selection.OMaths.Count > 0 Then
        Selection.OMaths(1).Range.Select
        Selection.Collapse wdCollapseEnd
    Else
    End If
    Selection.TypeText Text    ' causes problems with = sign coming under fraction line
    Selection.Move wdCharacter, -1
    If Selection.OMaths(1).Range.Font.Bold Then
        IsBold = True
        Selection.OMaths(1).Range.Font.Bold = False
    End If
    Selection.OMaths.BuildUp
    If IsBold Then Selection.OMaths(1).Range.Font.Bold = True
    omax.GoToEndOfSelectedMaths
End Sub

Function ValidateInput(Expr, Optional MathObj As OMath) As Boolean
    Dim ED As ErrorDefinition
   
    ValidateInput = True
    
    'validate brackets
    If GetCountOfChar(Expr, "(") <> GetCountOfChar(Expr, ")") Then
        ED.Title = Sprog.SyntaxError
        ED.MaximaOutput = Expr
        ED.Description = Sprog.A(876) & vbCrLf & Expr
    ElseIf InStr(Expr, "\left(") > 0 Or InStr(Expr, "\ast") > 0 Then
        ED.Title = Sprog.A(877)
        ED.Description = Sprog.A(878)
    End If
   
    If ED.Title <> vbNullString Then
        If Not MathObj Is Nothing Then
            ED.Description = ED.Title & vbCrLf & ED.Description
            ED.Title = Sprog.A(894)
            Application.ScreenUpdating = True
            MathObj.Range.Select
            omax.DefFejl = True
        End If
        
        UserFormError.SetErrorDefinition ED
        UserFormError.Show
        ValidateInput = False
    End If
    
End Function

Public Function GetVersion(n As String) As Single
' finds the version number based on folder name on maxima
Dim v As String, p As Integer
    p = InStr(n, "-")
    If p > 0 Then
        v = right(n, Len(n) - p)
    End If
    GetVersion = val(v)
End Function

#If Mac Then
#Else
Function GetMaximaPath() As String
' Finds the Maxima path. searches in Appdata and programfiles. If Maxima is present in both places, the latest version is used
' If the same version is present in both places, the path to appdata is returned.

    Dim FN As String, DN As String, v As Single ' For AppData
    Dim FN1 As String, DN1 As String, V1 As Single ' For Program files
    Dim s As String
    On Error Resume Next

    DN = Environ("AppData") & "\WordMat\"
    FN = Dir(DN & "Maxima*", vbDirectory)
    v = GetVersion(FN)
    Do
        s = Dir()
        If GetVersion(s) > v Then
            FN = s
            v = GetVersion(FN)
        End If
    Loop While s <> ""
    
    DN1 = GetProgramFilesDir() & "\WordMat\"
    FN1 = Dir(DN1 & "Maxima*", vbDirectory)
    If FN = vbNullString Then
        DN = DN1
        FN = FN1
        v = GetVersion(FN)
    Else
        V1 = GetVersion(FN1)
        If V1 > v Then
            DN = DN1
            FN = FN1
            v = V1
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
    Dim Arr() As String, i As Integer, ED As ErrorDefinition
    '    On Error Resume Next
    ValiderVariable = True
    If omax.AntalVars > 0 Then
        Arr = Split(omax.Vars, ";")
        For i = 0 To UBound(Arr)
            If Left(Arr(i), 1) = "_" Then
                ValiderVariable = False
                ED.Title = Sprog.A(879)
                ED.DefFejl = False
                ED.Description = Sprog.A(880) & Arr(i) & vbCrLf & vbCrLf & Sprog.A(881)
                UserFormError.SetErrorDefinition ED
                UserFormError.Show
                Exit Function
            End If
        Next
    End If
End Function
Public Sub IncreaseCalcCounter()
#If Mac Then
    Antalberegninger = Antalberegninger + 1
    AntalB = Antalberegninger
#Else
    AntalB = AntalB + 1
#End If
End Sub
