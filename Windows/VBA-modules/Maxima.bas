Attribute VB_Name = "Maxima"
Option Explicit
'Public UFWait As UserFormWaitForMaxima
Public UFSelectVar As UserFormSelectVar
Public UFSelectVars As UserFormSelectVars
Public omax As CMaxima
Public tid As Double
Private DeVarList As String

Public Sub PrepareMaxima(Optional Unit As Boolean = False)
'    Dim UFwait2 As UserFormWaitForMaxima

    On Error GoTo Fejl
    Dim op As Boolean
    If DebugWM Then
        UserFormDebug.Label_time.Caption = ""
        tid = Timer
    End If
#If Mac Then
    Dim d As Document
#End If
    
    SaveBackup
    
    Dim ufwait2 As UserFormWaitStartup
    op = False
    If Not SettingsRead Then ReadAllSettingsFromRegistry

    If omax Is Nothing Then
        LavRCMenu    ' højreklikmenu på ligninger
        SetMathAutoCorrect
        If ufwait2 Is Nothing Then Set ufwait2 = New UserFormWaitStartup
        ufwait2.Show vbModeless
        op = True
#If Mac Then
        Set d = ActiveDocument
#End If
        DoEvents
        '        Wait (0.1)
        Set omax = New CMaxima
    End If

    If MaxProc Is Nothing Then
        '        Set MaxProc = New MathMenu.MaximaProcessClass
        If Not op Then
            '                Set ufwait2 = New UserFormWaitStartup
            If ufwait2 Is Nothing Then Set ufwait2 = New UserFormWaitStartup
            ufwait2.Show vbModeless
            op = True
#If Mac Then
            Set d = ActiveDocument
#End If
            DoEvents
        End If
        On Error Resume Next
        
        Set MaxProc = GetMaxProc() 'CreateObject("MaximaProcessClass")
        If Err.Number <> 0 Then
            MsgBox Sprog.A(54), vbOKOnly, Sprog.Error
            GoTo Slut
        End If
        On Error GoTo Fejl
        MaxProc.Units = 0
        MaxProc.StartMaximaProcess
        If SettCheckForUpdate Then CheckForUpdateSilent
#If Mac Then
        If MaximaUnits Then
            TurnUnitsOn
        End If
#Else
        WaitForMaximaUntil
#End If
        If MaxProc.ErrCode > 0 Then
            MsgBox Sprog.A(55) & vbCrLf & "ErrCode: " & MaxProc.ErrCode & vbCrLf & vbCrLf & MaxProc.LastMaximaOutput, vbOKOnly, Sprog.Error
            GoTo Fejl
        End If
    End If
#If Mac Then
#Else
    If MaximaUnits Then
        If MaxProcUnit Is Nothing Then
            If Not op Then
                If ufwait2 Is Nothing Then Set ufwait2 = New UserFormWaitStartup
                ufwait2.Show vbModeless
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
                GoTo Slut
            End If
            On Error GoTo Fejl
            MaxProcUnit.Units = 1
            MaxProcUnit.OutUnits = omax.ConvertUnits(OutUnits)
            MaxProcUnit.StartMaximaProcess
            WaitForMaximaUnitUntil
        End If
    End If
#End If
    omax.PrepareNewCommand    ' nulstiller og finder definitioner

    If op Then
        Unload ufwait2
#If Mac Then
        d.Activate
#End If
    End If
    GoTo Slut
Fejl:
    On Error Resume Next
    Unload ufwait2
Slut:
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
Slut:
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
Slut:
End Sub

Public Sub PrepareMaximaNoSplash()
    On Error GoTo Slut
    If Not SettingsRead Then ReadAllSettingsFromRegistry
    If omax Is Nothing Then
        LavRCMenu    ' højreklikmenu på ligninger
        Set omax = New CMaxima
        If MaxProc Is Nothing Then
            '        Set MaxProc = New MathMenu.MaximaProcessClass
            Set MaxProc = GetMaxProc() 'CreateObject("MaximaProcessClass")
                On Error Resume Next
                If Err.Number <> 0 Then
                    MsgBox Sprog.A(54), vbOKOnly, Sprog.Error
                    GoTo Slut
                End If
                On Error GoTo Fejl
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
                    GoTo Slut
                End If
                On Error GoTo Fejl
                MaxProcUnit.Units = 1
                MaxProcUnit.OutUnits = omax.ConvertUnits(OutUnits)
                MaxProcUnit.StartMaximaProcess
            End If
        End If
#End If
        
    End If
    omax.PrepareNewCommand    ' nulstiller og finder definitioner
    '    WaitForMaximaUntil
Fejl:
Slut:
End Sub
Sub RestartMaxima()

    Set omax = New CMaxima
    If Not MaxProc Is Nothing Then
        MaxProc.CloseProcess
        MaxProc.StartMaximaProcess
    Else
        '        Set MaxProc = New MathMenu.MaximaProcessClass
        On Error Resume Next
        Set MaxProc = GetMaxProc() 'CreateObject("MaximaProcessClass")
        If Err.Number <> 0 Then
            MsgBox Sprog.A(54), vbOKOnly, Sprog.Error
            GoTo Slut
        End If
        On Error GoTo Fejl
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

    GoTo Slut
Fejl:
     MsgBox Sprog.A(55) & vbCrLf & "ErrCode: " & MaxProc.ErrCode & vbCrLf & vbCrLf & MaxProc.LastMaximaOutput, vbOKOnly, Sprog.Error
Slut:
End Sub
Sub MaximaCommand()
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    On Error GoTo Fejl
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled

    '    Set UFWait = New UserFormWaitForMaxima
    PrepareMaxima
    omax.prevspr = ""
    If Not omax.MaximaInstalled Then GoTo Slut
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo Slut
    End If

    '    UFWait.Show vbModeless
    omax.ExecuteMaximaCommand

    If omax.CheckForError Then GoTo Slut

    omax.GoToEndOfSelectedMaths
    Selection.TypeParagraph
    omax.InsertMaximaOutput
    '   UFWait.Hide
    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
Slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos

End Sub
Sub MaximaSolveInequality()
' løser een ulighed
    On Error GoTo Fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim eqs As Boolean
    Dim scrollpos As Double
    Dim sstart As Long
    Dim sslut As Long
    Dim variabel As String
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    If Not omax.MaximaInstalled Then GoTo Slut
    '    Set UFWait = New UserFormWaitForMaxima
    Set UFSelectVar = New UserFormSelectVar

    If Selection.OMaths.Count < 2 Then

        omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo Slut
    End If
        eqs = False
        If InStr(omax.Kommando, VBA.ChrW(8804)) > 0 Or InStr(omax.Kommando, VBA.ChrW(8805)) > 0 Then
            omax.Kommando = Replace(omax.Kommando, VBA.ChrW(8804), "<")
            omax.Kommando = Replace(omax.Kommando, VBA.ChrW(8805), ">")
            eqs = True
        End If
        omax.FindVariable
        UFSelectVar.vars = omax.vars
        UFSelectVar.Show
        variabel = UFSelectVar.SelectedVar
        If variabel = "" Then
            GoTo Slut
        End If
        If variabel = "" Then GoTo Slut
        omax.MaximaSolveInequality (variabel)

        If omax.CheckForError Then GoTo Slut

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
    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
Slut:
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
    Dim arr As Variant
    Dim fejlm As String
    On Error GoTo Fejl
    Application.ScreenUpdating = False
    '    LockWindow
    Dim IsSolved As Boolean
    Dim scrollpos As Double
    Dim UFSolvenumeric As New UserFormNumericQuestion
    scrollpos = ActiveWindow.VerticalPercentScrolled

    '    PrepareMaximaNoSplash  ' ved ikke hvorfor det var nosplash, erstattet
    PrepareMaxima
    omax.prevspr = ""

    Set UFSelectVar = New UserFormSelectVar
    Set UFSelectVars = New UserFormSelectVars
    Dim sstart As Long, sslut As Long
    If Not omax.MaximaInstalled Then GoTo Slut
    '    If UFWait Is Nothing Then Set UFWait = New UserFormWaitForMaxima
    '    Set UFWait = New UserFormWaitForMaxima
    sstart = Selection.start
    sslut = Selection.End

    If omax.deffejl Then
        MsgBox Sprog.DefError & vbCrLf & VisDef & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
        GoTo Slut
    End If

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.EquationMissingError, vbOKOnly, Sprog.EquationMissingError2
        GoTo Slut
    End If
    If sstart = sslut Then
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    If InStr(Selection.OMaths(1).Range.text, "<") > 1 Or InStr(Selection.OMaths(1).Range.text, ">") > 1 Or InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8804)) > 1 Or InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8805)) > 1 Then
        MaximaSolveInequality
        GoTo Slut
    End If
    If InStr(Selection.OMaths(1).Range.text, "=") < 1 Then
        Dim Result As VbMsgBoxResult
        Result = MsgBox(Sprog.A(141), vbYesNo, Sprog.Warning)
        If Result = vbNo Then GoTo Slut
    End If

    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo Slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo Slut

    If Selection.OMaths.Count < 2 And InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8743)) < 1 Then
        ' kun 1 ligning

        UFSolvenumeric.Ligning = omax.Kommando

        omax.FindVariable
        If variabel = vbNullString Then
            UFSelectVar.vars = omax.vars
            UFSelectVar.defs = omax.defstring
            UFSelectVar.Show
        variabel = UFSelectVar.SelectedVar
        End If
        If variabel = "" Then
            GoTo Slut
        End If
        If variabel = "" Then GoTo Slut
        omax.tempDefs = UFSelectVar.tempDefs
        omax.MaximaSolve (variabel)
        If omax.StopNow Then
            If omax.AntalVars > 1 Then
                GoTo Slut
            End If
        End If
        If omax.CheckForError Then GoTo Slut

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
            If resultat = vbCancel Then GoTo Slut
        End If

        variabel = omax.ConvertToWordSymbols(variabel)

        If omax.StopNow Or (omax.IsAllSolved(omax.MaximaOutput, variabel, VBA.ChrW(8744)) = "false" And Not (InStr(variabel, "+") > 0)) Then
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
            '                MsgBox "Ligningen havde ingen løsninger for variabel " & variabel & ". Det kan muligvis skyldes at du har indtastet en forkert variabel.", vbOKOnly, "Ingen løsninger"
        ElseIf omax.MaximaOutput = "" Or InStr(omax.KommentarOutput, "Lisp error") > 0 Then    ' kunne ikke løses
            GoTo stophop
            If omax.AntalVars > 1 Then
                MsgBox Sprog.A(131), vbOKOnly, "WordMat"
                GoTo Slut
            End If
            omax.PrepareNewCommand
            omax.tempDefs = UFSelectVar.tempDefs
            UFSolvenumeric.Label_omskrevet.Caption = omax.MaximaOutput
            UFSolvenumeric.FejlMeld = omax.KommentarOutput
            UFSolvenumeric.variabel = variabel
            UFSolvenumeric.Show
            If UFSolvenumeric.Result = "num" Then
                Selection.End = sslut    ' slut skal være først eller går det galt
                Selection.start = sstart
                ActiveWindow.VerticalPercentScrolled = scrollpos
                MaximaSolveNumeric UFSelectVar.ListBox_vars.text
            Else
                If UFSolvenumeric.Result = "nsolve" Then
                    InsertForklaring Sprog.EquationSolvedNumFor & variabel & Sprog.ByCAS, False
                ElseIf UFSolvenumeric.Result = "omskriv" Then
                    InsertForklaring "", True
                End If
                omax.InsertMaximaOutput
                Selection.TypeParagraph
                Selection.TypeText UFSolvenumeric.Kommentar
            End If
        ElseIf InStr(omax.MaximaOutput, VBA.ChrW(8477)) > 0 Then  ' alle er løsning
            omax.InsertMaximaOutput
            Selection.TypeParagraph
            Selection.TypeText Sprog.A(132)
        ElseIf Not LmSet And Not IsSolved Then
stophop:             If omax.AntalVars > 1 Then
                UFSolvenumeric.Result = ""
            Else
                UFSolvenumeric.Label_omskrevet.Caption = omax.MaximaOutput
                omax.PrepareNewCommand
                omax.tempDefs = UFSelectVar.tempDefs
                UFSolvenumeric.FejlMeld = omax.KommentarOutput
                UFSolvenumeric.variabel = variabel
                UFSolvenumeric.Show
            End If
            If UFSolvenumeric.Result = "num" Then
                Selection.End = sslut    ' slut skal være først eller går det galt
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

        Else    ' hvis der er løsning
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
        UFSelectVars.defs = omax.defstring
        UFSelectVars.Show
        variabel = UFSelectVars.SelectedVar

        '    variabel = InputBox("Indtast variable som ligningssystemet skal løses for adskilt af komma", "Variable", "x,y")

        If variabel = "" Then GoTo Slut
        omax.tempDefs = UFSelectVars.tempDefs
        '    UFWait.ActionToPerform = "solvesystem"
        '    UFWait.VarParam = variabel

        '    Set UFWait.omax = omax
        '    UFWait.Show
        omax.SolveSystem (variabel)
        If omax.StopNow Then GoTo Slut
        If omax.CheckForError Then GoTo Slut
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph
        If omax.StopNow Then GoTo Slut
        Application.ScreenUpdating = False
        '    omax.KommentarOutput = TranslateReplaceComment(omax.KommentarOutput)

        If omax.MaximaOutput = "?merror(""Anumberwasfoundwhereavariablewasexpected-`solve'"")" Then
            MsgBox Sprog.A(133), vbOKOnly, Sprog.Error
            GoTo Slut
        End If

        If (omax.deffejl = True) Then
            MsgBox Sprog.DefError & vbCrLf & VisDef & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
            GoTo Slut
        End If

        variabel = Replace(omax.ConvertToWordSymbols(variabel), ";", ",")
        If MaximaForklaring Then
            InsertForklaring Sprog.A(134) & variabel & Sprog.A(135)
        End If

        If Len(omax.MaximaOutput) > 1 Then
            omax.InsertMaximaOutput
            arr = Split(omax.MaximaOutput, "=")
            If UBound(arr) = 1 Then
                If InStr(arr(0), variabel) > 0 And InStr(arr(1), variabel) > 0 Then
                    '                    Result = MsgBox("Maxima kunne ikke løse ligningssystemet. Den var for kompleks." & vbCrLf & vbCrLf & omax.KommentarOutput & vbCrLf & vbCrLf & "Tryk OK hvis du vil forsøge at løse ligningen numerisk.", vbOKCancel, "Fejl")
                    UFSolvenumeric.FejlMeld = omax.KommentarOutput
                    UFSolvenumeric.Show
                    If UFSolvenumeric.Result = "num" Then
                        Selection.End = sslut    ' slut skal være først eller går det galt
                        Selection.start = sstart
                        ActiveWindow.VerticalPercentScrolled = scrollpos
                        MaximaSolveNumeric
                    Else
                        Selection.TypeParagraph
                        Selection.TypeText Sprog.A(136)
                    End If
                End If
            End If
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
    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
Slut:
    '    omax.Luk
    On Error Resume Next
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
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.TypeText text:=vbTab
    End If
    Selection.Font.Size = 8
    Selection.Font.ColorIndex = wdGray50
    Selection.Font.Italic = True

    Selection.TypeText ForklarTekst
    If omax.tempDefs <> "" Then
        '                If Not MaximaSeparator Then
        '                    tdefs = Replace(omax.TempDefs, ",", ";")
        '                    tdefs = Replace(tdefs, ".", ",")
        '                Else
        tdefs = omax.tempDefs
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
    '    LockWindow
    Dim scrollpos As Double
    scrollpos = ActiveWindow.VerticalPercentScrolled

    PrepareMaxima
    omax.prevspr = ""
    '    Set UFSelectVar = New UserFormSelectVar
    Set UFSelectVars = New UserFormSelectVars
    Dim variabel As String
    Dim sstart As Long, sslut As Long
    If Not omax.MaximaInstalled Then GoTo Slut
    '    If UFWait Is Nothing Then Set UFWait = New UserFormWaitForMaxima
    '    Set UFWait = New UserFormWaitForMaxima
    sstart = Selection.start
    sslut = Selection.End

    If omax.deffejl Then
        MsgBox Sprog.DefError & vbCrLf & VisDef & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
        GoTo Slut
    End If

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo Slut
    End If
    If sstart = sslut Then
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    If InStr(Selection.OMaths(1).Range.text, "=") < 1 Then
        Dim Result As VbMsgBoxResult
        Result = MsgBox(Sprog.A(141), vbYesNo, Sprog.Warning)
        If Result = vbNo Then GoTo Slut
    End If

    If Selection.OMaths.Count < 2 And InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8743)) < 1 Then
        GoTo Slut
    Else    ' ligningssystem

        omax.ReadSelection
        omax.FindVariable
        UFSelectVars.Eliminate = True
        UFSelectVars.NoEq = omax.AntalKom
        UFSelectVars.vars = omax.vars
        UFSelectVars.defs = omax.defstring
        UFSelectVars.Caption = Sprog.A(377)
        UFSelectVars.Show
        variabel = UFSelectVars.SelectedVar

        '    variabel = InputBox("Indtast variable som ligningssystemet skal løses for adskilt af komma", "Variable", "x,y")

        If variabel = "" Then GoTo Slut
        omax.tempDefs = UFSelectVars.tempDefs
        '    UFWait.ActionToPerform = "solvesystem"
        '    UFWait.VarParam = variabel
        '    Set UFWait.omax = omax
        '    UFWait.Show
        omax.Eliminate (variabel)
        If omax.StopNow Then GoTo Slut
        If omax.CheckForError Then GoTo Slut
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph
        If omax.StopNow Then GoTo Slut
        Application.ScreenUpdating = False
        '    omax.KommentarOutput = TranslateReplaceComment(omax.KommentarOutput)

        If omax.MaximaOutput = "?merror(""Anumberwasfoundwhereavariablewasexpected-`solve'"")" Then
            MsgBox Sprog.A(133), vbOKOnly, Sprog.Error
            GoTo Slut
        End If

        If (omax.deffejl = True) Then
            MsgBox Sprog.DefError & vbCrLf & VisDef & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
            GoTo Slut
        End If

        variabel = Replace(omax.ConvertToWordSymbols(variabel), ";", ",")

        If MaximaForklaring Then
            '            Selection.TypeParagraph
            ForklarTekst = Sprog.A(142) & variabel & Sprog.A(143)
            InsertForklaring ForklarTekst, False
            '            Selection.TypeText ForklarTekst
            If omax.tempDefs <> "" Then
                If Not MaximaSeparator Then
                    tdefs = Replace(omax.tempDefs, ",", ";")
                    tdefs = Replace(tdefs, ".", ",")
                Else
                    tdefs = omax.tempDefs
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
    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
Slut:
    '    omax.Luk
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
    '   UnLockWindow
End Sub
Sub MaximaNsolve()
    Dim arr As Variant
    Dim fejlm As String
    Dim solutions As String
    Dim UFnsolve As New UserFormNumericQuestion
    '    On Error GoTo fejl
    Application.ScreenUpdating = False
    '    LockWindow
    Dim IsSolved As Boolean
    Dim scrollpos As Double
    scrollpos = ActiveWindow.VerticalPercentScrolled

    '    PrepareMaximaNoSplash
    PrepareMaxima
    omax.prevspr = ""
    Set UFSelectVar = New UserFormSelectVar
    Set UFSelectVars = New UserFormSelectVars
    Dim variabel As String
    Dim sstart As Long, sslut As Long
    If Not omax.MaximaInstalled Then GoTo Slut
    '    If UFWait Is Nothing Then Set UFWait = New UserFormWaitForMaxima
    '    Set UFWait = New UserFormWaitForMaxima
    sstart = Selection.start
    sslut = Selection.End

    If omax.deffejl Then
        MsgBox Sprog.DefError & vbCrLf & VisDef & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
        GoTo Slut
    End If

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.EquationMissingError, vbOKOnly, Sprog.EquationMissingError2
        GoTo Slut
    End If
    If sstart = sslut Then
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    If InStr(Selection.OMaths(1).Range.text, "<") > 1 Or InStr(Selection.OMaths(1).Range.text, ">") > 1 Or InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8804)) > 1 Or InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8805)) > 1 Then
        MaximaSolveInequality
        GoTo Slut
    End If
    If InStr(Selection.OMaths(1).Range.text, "=") < 1 Then
        Dim Result As VbMsgBoxResult
        Result = MsgBox(Sprog.A(141), vbYesNo, Sprog.Warning)
        If Result = vbNo Then GoTo Slut
    End If

    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo Slut
    End If

    If Selection.OMaths.Count < 2 And InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8743)) < 1 Then
        ' kun 1 ligning

        UFnsolve.Ligning = omax.Kommando

        omax.FindVariable
        UFSelectVar.vars = omax.vars
        UFSelectVar.defs = omax.defstring
        UFSelectVar.Show
        
        variabel = UFSelectVar.SelectedVar
        If variabel = "" Then
            GoTo Slut
        End If
        If variabel = "" Then GoTo Slut
        omax.tempDefs = UFSelectVar.tempDefs

        UFnsolve.Label_overskrift.Caption = ""
        UFnsolve.Label_omskrevet.Caption = ""
        UFnsolve.FejlMeld = ""
        UFnsolve.variabel = variabel
        UFnsolve.Show

        If omax.StopNow Then GoTo Slut
        If omax.CheckForError Then GoTo Slut
        If UFnsolve.Result = "afbryd" Then GoTo Slut

        If UFnsolve.Result = "num" Then
            Selection.End = sslut    ' slut skal være først ellers går det galt
            Selection.start = sstart
            ActiveWindow.VerticalPercentScrolled = scrollpos
            Unload UFnsolve
            MaximaSolveNumeric UFSelectVar.ListBox_vars.text
            GoTo Slut
        End If

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
            If resultat = vbCancel Then GoTo Slut
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
        If inp = "" Then GoTo Slut    ' trykket cancel
        arr = Split(inp, ListSeparator)

        Variable = ""
        For j = 0 To UBound(arr)
            arr2 = Split(arr(j), "=")
            Variable = Variable & Trim(arr2(0)) & ","
            If UBound(arr2) = 0 Or Trim(arr2(1)) = "" Then
                guess = guess & ",1"
            Else
                guess = guess & Replace(Trim(arr2(1)), ",", ".") & ","
            End If
        Next
        Variable = Left(Variable, Len(Variable) - 1)
        guess = Left(guess, Len(guess) - 1)


        If Variable = "" Then GoTo Slut

        '    UFWait.ActionToPerform = "solvesystemnumeric"
        '    Set UFWait.omax = omax
        '    UFWait.VarParam = Variable
        '    UFWait.Param2 = guess
        '    UFWait.Show
        omax.SolveSystemNumeric Variable, guess
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
    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
Slut:
    '    omax.Luk
    On Error Resume Next
    Unload UFnsolve
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
    '   UnLockWindow
End Sub
Sub MaximaSolveNumeric(Optional var As String)
' grafisk løsning + newton + intervalmetode
' var indsættes som den variabel der løses for
    On Error GoTo Fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim variabel As String
    Dim t As String
    Dim scrollpos As Double
    Dim VarGuess As String
    Dim arr As Variant
    Dim arr2 As Variant
    Dim j As Integer
    scrollpos = ActiveWindow.VerticalPercentScrolled
    Dim sstart As Long
    Dim sslut As Long
    sstart = Selection.start
    sslut = Selection.End

    If Not omax.MaximaInstalled Then GoTo Slut
    '    Set UFWait = New UserFormWaitForMaxima
    Dim UFSolvenumeric As New UserFormSolveNumeric

    Set UFSelectVars = New UserFormSelectVars

    '    MsgBox WordWindowNavn
    Dim sep As String
    If MaximaSeparator Then
        sep = ","
    Else
        sep = ";"
    End If

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.EquationMissingError, vbOKOnly, Sprog.EquationMissingError2
        GoTo Slut
    End If
    If sstart = sslut Then
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    If InStr(Selection.OMaths(1).Range.text, "=") < 1 Then
        Dim Result As VbMsgBoxResult
        Result = MsgBox(Sprog.A(141), vbYesNo, Sprog.Warning)
        If Result = vbNo Then GoTo Slut
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
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo Slut
    End If
        omax.FindVariable
        omax.tempDefs = UFSelectVar.tempDefs

        If UFSolvenumeric.SelectedVar = "" Then GoTo Slut
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
        If omax.StopNow Then GoTo Slut
        If omax.CheckForError Then GoTo Slut
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
            If resultat = vbCancel Then GoTo Slut
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

    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
Slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub
Sub beregn()
'    MsgBox WordWindowNavn
'    Dim omax As New CMaxima
    Dim fejlm As String
    On Error GoTo Fejl
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

    If Not omax.MaximaInstalled Then GoTo Slut
    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo Slut
    End If
    If Selection.OMaths.Count > 1 Then
        MsgBox Sprog.A(149), vbOKOnly, Sprog.Error
        GoTo Slut
    End If

    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo Slut
    End If
    If omax.Kommando = "" Then GoTo Slut
    If Not ValidateInput(omax.Kommando) Then GoTo Slut
    '    Set UFWait.omax = omax
    '    UFWait.ActionToPerform = "beregn"
    '    UFWait.Show
    '    If omax.StopNow Then GoTo slut
    omax.beregn
    If omax.StopNow Then GoTo Slut
    If omax.CheckForError Then GoTo Slut
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
        MsgBox fejlm & vbCrLf & vbCrLf & omax.KommentarOutput & vbCrLf & MaxProc.LastMaximaOutput, vbOKOnly, Sprog.Error
        RestartMaxima
    End If
#If Mac Then
#Else
        Oundo.EndCustomRecord
#End If

    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
Slut:
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
Sub Omskriv()
    On Error GoTo Fejl
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
    If Not omax.MaximaInstalled Then GoTo Slut
    '    MsgBox WordWindowNavn

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo Slut
    End If
    If sstart = sslut Then    ' sikrer at gotoendofselectedmaths virker
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo Slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo Slut

    '    omax.OpenCmd
    omax.FindVariable
    UFomskriv.vars = omax.vars
    UFomskriv.Show
    If UFomskriv.annuller Then GoTo Slut

    omax.tempDefs = UFomskriv.tempDefs
    omax.Omskriv UFomskriv.CheckBox_vidnotation.Value, UFomskriv.CheckBox_auto.Value, UFomskriv.CheckBox_factor.Value, UFomskriv.CheckBox_expand.Value, UFomskriv.CheckBox_rationaliser.Value, UFomskriv.CheckBox_trigreduce.Value
    If omax.StopNow Then GoTo Slut
    If omax.CheckForError Then GoTo Slut

    If Not (UFomskriv.SammeLinje) Then
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph

        '    omax.CloseCmd
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

    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
Slut:
    MaximaCifre = SaveSettingsCifre
    MaximaExact = SaveSettingsExact
    OutUnits = SaveSettingsOutunits
    MaximaLogOutput = SaveSettingsLog
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos

End Sub
Sub reducer()
    On Error GoTo Fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    If Not omax.MaximaInstalled Then GoTo Slut
    '    MsgBox WordWindowNavn

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo Slut
    End If
    If sstart = sslut Then    ' sikrer at gotoendofselectedmaths virker
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo Slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo Slut

    '    omax.OpenCmd
    omax.Reduce
    If omax.StopNow Then GoTo Slut
    If omax.CheckForError Then GoTo Slut
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

    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
Slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
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
    If Not omax.MaximaInstalled Then GoTo Slut
    '    MsgBox WordWindowNavn

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo Slut
    End If
    If sstart = sslut Then    ' sikrer at gotoendofselectedmaths virker
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo Slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo Slut

    '    omax.OpenCmd
    omax.CompareTest
    If omax.StopNow Then GoTo Slut
    If omax.CheckForError Then GoTo Slut
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

    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
Slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos

End Sub
Sub faktoriser()
    On Error GoTo Fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    If Not omax.MaximaInstalled Then GoTo Slut
    '    MsgBox WordWindowNavn

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo Slut
    End If
    If sstart = sslut Then    ' sikrer at gotoendofselectedmaths virker
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo Slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo Slut

    '    omax.OpenCmd
    omax.Factor
    If omax.StopNow Then GoTo Slut
    If omax.CheckForError Then GoTo Slut
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

    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
Slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
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
    Dim scrollpos As Double
    scrollpos = ActiveWindow.VerticalPercentScrolled
    If Not omax.MaximaInstalled Then GoTo Slut
    '    MsgBox WordWindowNavn
    '    Set UFWait = New UserFormWaitForMaxima

    If sstart = sslut Then    ' sikrer at gotoendofselectedmaths virker
        Selection.OMaths(1).ParentOMath.Range.Select
    End If

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo Slut
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo Slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo Slut

    '    omax.OpenCmd
    omax.Expand
    If omax.StopNow Then GoTo Slut
    If omax.CheckForError Then GoTo Slut
    omax.GoToEndOfSelectedMaths
    Selection.TypeParagraph
    '    omax.CloseCmd
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
    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
Slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos

End Sub
Sub Differentier()
    On Error GoTo Fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim variabel As String
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    If Not omax.MaximaInstalled Then GoTo Slut
    '    MsgBox WordWindowNavn
    If sstart = sslut Then    ' sikrer at gotoendofselectedmaths virker
        Selection.OMaths(1).ParentOMath.Range.Select
    End If

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo Slut
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo Slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo Slut

    variabel = InputBox(Sprog.A(51), Sprog.Variable, "x")
    If variabel = "" Then GoTo Slut
    '    omax.OpenCmd
    omax.Differentier variabel
    If omax.StopNow Then GoTo Slut
    If omax.CheckForError Then GoTo Slut
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

    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
Slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos

End Sub
Sub Integrer()
    On Error GoTo Fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim variabel As String
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    '    MsgBox WordWindowNavn
    If Not omax.MaximaInstalled Then GoTo Slut

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo Slut
    End If
    If sstart = sslut Then    ' sikrer at gotoendofselectedmaths virker
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo Slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo Slut

    variabel = InputBox(Sprog.A(51), Sprog.Variable, "x")
    If variabel = "" Then GoTo Slut
    '    omax.OpenCmd
    omax.Integrer variabel
    If omax.StopNow Then GoTo Slut
    If omax.CheckForError Then GoTo Slut
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

    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
Slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
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
    On Error GoTo Fejl

    ea.SetNormalBrackets

    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled

    PrepareMaxima

    If Not omax.MaximaInstalled Then GoTo Slut

    If Selection.OMaths.Count > 0 And sstart = sslut Then    ' sikrer at gotoendofselectedmaths virker
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    omax.ReadSelection
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo Slut
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
    If UFdiffeq.luk Then GoTo Slut


    '    MsgBox omax.MaximaOutput

    If UFdiffeq.InsertType = 3 Then
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph
        omax.MaximaOutput = UFdiffeq.ListOutput
        omax.InsertMaximaOutput
    End If


    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
Slut:
    On Error Resume Next
    Selection.End = sslut    ' slut skal være først eller går det galt
    Selection.start = sstart
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub
Function GetRHS(s As String) As String
    Dim arr As Variant
    s = omax.ConvertToAscii(s)
    arr = Split(s, "=")
    If UBound(arr) > 0 Then
        GetRHS = arr(UBound(arr))
    Else
        GetRHS = s
    End If
End Function
Function GetLHSvar(s As String) As String
    Dim arr As Variant
    Dim ea As New ExpressionAnalyser
    Dim var As String, i As Integer
    If s = vbNullString Then
        GetLHSvar = vbNullString
        Exit Function
    End If
    s = omax.ConvertToAscii(s)
    arr = Split(s, "=")
    s = arr(0)
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
    On Error GoTo Fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim vars As String
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    Dim t As String
    Dim arr As Variant
    Dim UFdiffeq As New UserFormDiffEq
    Dim ea As New ExpressionAnalyser
    ea.SetNormalBrackets
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    '    MsgBox WordWindowNavn
    If Not omax.MaximaInstalled Then GoTo Slut

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo Slut
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
        GoTo Slut
    End If

    If Not ValidateInput(omax.Kommando) Then GoTo Slut

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
        UFdiffeq.defs = omax.defstring
        UFdiffeq.TextBox_funktion.text = funktion
        UFdiffeq.TextBox_variabel.text = variabel
        UFdiffeq.Label_ligning.Caption = Replace(omax.ConvertToAscii(omax.Kommando), " ", "")
        UFdiffeq.Show
        If UFdiffeq.luk Then GoTo Slut
        variabel = UFdiffeq.TextBox_variabel.text
        funktion = UFdiffeq.TextBox_funktion.text
    End If
    If variabel = "" Then GoTo Slut
    '    omax.OpenCmd
    omax.tempDefs = UFdiffeq.tempDefs
    omax.SolveDE funktion, variabel, UFdiffeq.TextBox_startx, UFdiffeq.TextBox_starty, UFdiffeq.TextBox_starty2, UFdiffeq.TextBox_bcx, UFdiffeq.TextBox_bcy
    If omax.StopNow Then GoTo Slut
    If omax.CheckForError Then GoTo Slut
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

    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
Slut:
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
    Selection.OMaths.BuildUp
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
Sub unicodevals3()
    Dim text As String
    Dim j As Integer
    Dim i As Integer
    Dim s As String

    Selection.OMaths.Linearize
    Selection.OMaths(1).ConvertToNormalText
    text = Selection.text
    Selection.OMaths(1).ConvertToMathText
    Selection.OMaths(1).Range.Select
    Selection.OMaths.BuildUp

    For j = 1 To Len(text)
        i = AscW(Mid(text, j, 1))
        If i > 200 Then
            s = s & """ & VBA.ChrW(" & i & ") & """
        Else
            s = s & Mid(text, j, 1)
        End If
    Next
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    Selection.TypeText (s)

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

