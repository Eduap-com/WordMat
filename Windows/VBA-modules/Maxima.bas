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

    On Error GoTo Fejl
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
        On Error GoTo Fejl
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
            GoTo Slut
        End If
        If DllConnType = 2 Then
            If SettCheckForUpdate Then CheckForUpdateSilent
            GoTo finish
        End If
        On Error GoTo Fejl
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
            GoTo Fejl
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

finish:
    omax.ConvertLnLog = True ' andre funktioner kan ændre denne. den nulstilles
    If Not omax.PrepareNewCommand Then    ' nulstiller og finder definitioner
        GoTo Fejl
    End If
    On Error Resume Next
    If op Then
'        Unload UfWait2
#If Mac Then
        D.Activate
#End If
    End If
    PrepareMaxima = True
    GoTo Slut
Fejl:
    On Error Resume Next
'    If Not UfWait2 Is Nothing Then Unload UfWait2
    PrepareMaxima = False
Slut:
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
'        LavRCMenu    ' højreklikmenu på ligninger
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
            GoTo Slut
        End If
        On Error GoTo Fejl
        If omax Is Nothing Then
            Set omax = New CMaxima
        End If
    End If
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
    Dim sstart As Long, sslut As Long, DontGoBack As Boolean
    On Error GoTo Fejl
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled

    '    Set UFWait = New UserFormWaitForMaxima
    PrepareMaxima
    omax.prevspr = ""
    If Not omax.MaximaInstalled Then GoTo Slut
    
    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo Slut
    ElseIf Selection.OMaths.Count > 0 Then
        omax.ReadSelection
    Else
        omax.Kommando = Selection.Range.text
        DontGoBack = True
    End If
    
    
    
    If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
        MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
        GoTo Slut
    End If

    '    UFWait.Show vbModeless
    If CASengine = 0 Then
        omax.ExecuteMaximaCommand
    ElseIf CASengine = 1 Then
        OpenGeoGebraWeb omax.Kommando, "CAS", True, True
        GoTo Slut
    ElseIf CASengine = 2 Then
        omax.MaximaOutput = RunGeoGebraDirect(omax.Kommando)
    End If


    If CheckForError Then
        scrollpos = ActiveWindow.VerticalPercentScrolled
        sslut = Selection.End
        sstart = Selection.start
        GoTo Slut
    End If

    omax.GoToEndOfSelectedMaths DontGoBack
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
Sub MaximaSolveInequality(Optional variabel As String)
' løser een ulighed
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
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo Slut
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
        If variabel = vbNullString Then
            UFSelectVar.Vars = omax.Vars
            UFSelectVar.DefS = omax.DefString
            UFSelectVar.Show
            variabel = UFSelectVar.SelectedVar
        End If
        
        If variabel = "" Then
            GoTo Slut
        End If
        If variabel = "" Then GoTo Slut
        
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
            GoTo Slut
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
        
        
        If CheckForError Then
            scrollpos = ActiveWindow.VerticalPercentScrolled
            sslut = Selection.End
            sstart = Selection.start
            GoTo Slut
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
          Dim Arr As Variant, s As String, t As String, v As String
          Dim fejlm As String
10        On Error GoTo Fejl
20        Application.ScreenUpdating = False
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
          
30        scrollpos = ActiveWindow.VerticalPercentScrolled
          
40        TempCas = CASengine
      '    On Error Resume Next: oData.GetFromClipboard:   ClipText = oData.GetText: On Error GoTo fejl ' Skulle sikre at clipboard ikke ændres, men virker ikke
50        PrepareMaxima
60        omax.prevspr = ""

70        Set UFSelectVar = New UserFormSelectVar
80        UFSelectVar.NoEq = 1
          
90        SaveSettingsCifre = MaximaCifre
100       SaveSettingsExact = MaximaExact
110       SaveSettingsOutunits = OutUnits
120       SaveSettingsLog = MaximaLogOutput
130       SaveSettingsDecOutType = MaximaDecOutType
          
          
140       If CASengine = 0 And Not omax.MaximaInstalled Then GoTo Slut
150       sstart = Selection.start
160       sslut = Selection.End

170       If omax.DefFejl Then GoTo Slut

180       If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
190           MsgBox2 Sprog.EquationMissingError, vbOKOnly, Sprog.EquationMissingError2
200           GoTo Slut
210       End If
220       If sstart = sslut Then
230           Selection.OMaths(1).ParentOMath.Range.Select
240       End If
250       If InStr(Selection.OMaths(1).Range.text, "<") > 1 Or InStr(Selection.OMaths(1).Range.text, ">") > 1 Or InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8804)) > 1 Or InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8805)) > 1 Then
260           MaximaSolveInequality variabel
270           GoTo Slut
280       End If
290       p = InStr(Selection.OMaths(1).Range.text, "=")
300       If p < 1 Then
310           MsgBox Sprog.A(141), vbOKOnly, Sprog.Error
320           GoTo Slut
330       Else
340           p2 = InStr(p + 1, Selection.OMaths(1).Range.text, "=")
350           If p2 > 0 Then
360               If Sprog.SprogNr = 1 Then
370                   MsgBox "Der kan ikke være to ligmedtegn i en ligning", vbOKOnly, "Fejl"
380               Else
390                   MsgBox "You cant place two equal sign in an equation", vbOKOnly, "Fejl"
400               End If
410               GoTo Slut
420           End If
430       End If

440       omax.ReadSelection
450       If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
460           MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
470           GoTo Slut
480       End If

490       If Not ValidateInput(omax.Kommando) Then GoTo Slut

500       If Selection.OMaths.Count < 2 And InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8743)) < 1 Then
              ' kun 1 ligning

510           UFSolvenumeric.Ligning = omax.Kommando
              
520           omax.FindVariable
530           If Not ValiderVariable Then GoTo Slut
540           SaveKommando = omax.Kommando
newcas:
550           omax.StopNow = False
560           omax.Kommando = SaveKommando
570           If CASengine > 0 And Not AllTrig Then ' På geogebra skal der via vba genkendes om det er trigonometrisk ligning
580               If Not InStr(omax.Vars, ";") > 0 Then ' metoden virker kun med 1 variabel
590                   ea.SetNormalBrackets
600                   ea.text = omax.Kommando
610                   ea.text = Replace(ea.text, VBA.ChrW(8289), "")
620                   s = ""
630                   Do
640                       v = ea.GetNextVar()
650                       If v = "sin" Or v = "cos" Or v = "tan" Then
660                           t = ea.GetNextBracketContent()
670                           If InStr(t, omax.Vars) > 0 Then
680                               If v = "cos" Then
690                                   s = "180"
700                               Else
710                                   s = "90"
720                               End If
730                           End If
740                       End If
750                       ea.Pos = ea.Pos + 1
760                   Loop While v <> ""
770                   If s <> "" And Radians Then
780                       If s = "90" Then
790                           s = ChrW(960) & "/2" ' pi
800                       Else
810                           s = ChrW(960) '"pi"
820                       End If
830                   End If
840                   If s <> "" Then
850                       UFSelectVar.TextBox_def.text = "0<=" & omax.Vars & "<=" & s & VbCrLfMac
860                       UFSelectVar.TempDefs = "0<=" & omax.Vars & "<=" & s
870                   End If
880               End If
890           End If
900           If variabel = vbNullString Then
910               UFSelectVar.Vars = omax.Vars
920               UFSelectVar.DefS = omax.DefString
930               UFSelectVar.Show
940               variabel = UFSelectVar.SelectedVar
950           End If
960           If variabel = "" Then GoTo Slut
970           omax.TempDefs = Replace(UFSelectVar.TempDefs, "%pi", "pi")
980           If CASengine = 1 Or CASengine = 2 Then
990               s = Trim(omax.Kommando)
1000              s = Replace(s, vbCrLf, "")
1010              s = Replace(s, vbCr, "")
1020              s = Replace(s, vbLf, "")
1030              If MaximaComplex Then
1040                  s = "Csolve(" & Replace(s, ",", ".") & "," & variabel & ")"
1050              Else
1060                  s = "solve(" & Replace(s, ",", ".") & "," & variabel & ")"
1070              End If
1080              If MaximaDecOutType = 3 Then
1090                  s = "ScientificText(" & s & " , " & MaximaCifre & ")"
1100              ElseIf MaximaExact = 2 Then
1110                  s = "numeric(" & s & " , " & MaximaCifre & ")"
1120              End If

1130          End If
              
1140          If CASengine = 0 Then
1150              omax.MaximaSolve (variabel)
1160          ElseIf CASengine = 1 Then
1170              If MaximaForklaring Then
1180                  omax.GoToEndOfSelectedMaths
1190                  If MaximaForklaring Then
1200                      Selection.TypeParagraph
1210                      InsertForklaring Sprog.EquationSolvedFor & variabel & Sprog.ByCAS, True
1220                      Selection.TypeParagraph
1230                  End If
1240              End If
1250              OpenGeoGebraWeb s, "CAS", True, True
1260              GoTo Slut
1270          ElseIf CASengine = 2 Then
1280              Call RunGeoGebraDirect(s)
1290              If omax.MaximaOutput = "{}" Then
1300                  omax.MaximaOutput = variabel & VBA.ChrW(8712) & VBA.ChrW(8709)
1310              ElseIf omax.MaximaOutput = "{" & variabel & "=" & variabel & "}" Or omax.MaximaOutput = "{x=x}" Then
1320                  omax.MaximaOutput = variabel & VBA.ChrW(8712) & VBA.ChrW(8477)
1330              Else
1340                  If Left(omax.MaximaOutput, 1) = "{" Then omax.MaximaOutput = Mid(omax.MaximaOutput, 2, Len(omax.MaximaOutput) - 2)
1350                  ea.text = omax.MaximaOutput
1360                  omax.MaximaOutput = ""
1370                  ea.Pos = 1
1380                  Do
1390                      s = Trim(ea.GetNextListItem(ea.Pos, ";"))
1400                      If s <> "" Then omax.MaximaOutput = omax.MaximaOutput & s & "    " & VBA.ChrW(8744) & "    "
1410                  Loop While s <> ""
1420                  If omax.MaximaOutput <> "" Then omax.MaximaOutput = Left(omax.MaximaOutput, Len(omax.MaximaOutput) - 9)
      '                omax.MaximaOutput = Replace(omax.MaximaOutput, ";", "    " & VBA.ChrW(8744) & "    ")    ' komma erstattes med eller
1430              End If
1440              Application.Activate
1450          End If
                      
              
1460          If omax.StopNow Then
1470              If omax.AntalVars > 1 Then
1480                  GoTo Slut
1490              End If
1500          End If
1510          If CheckForError Then
1520              scrollpos = ActiveWindow.VerticalPercentScrolled
1530              sslut = Selection.End
1540              sstart = Selection.start
1550              GoTo Slut
1560          End If

              Dim Oundo As UndoRecord
1570          Set Oundo = Application.UndoRecord
1580          Oundo.StartCustomRecord
              
1590          omax.GoToEndOfSelectedMaths
1600          Selection.TypeParagraph
1610          If Len(omax.MaximaOutput) > 250 Then
                  Dim resultat As VbMsgBoxResult
1620              resultat = MsgBox(Sprog.A(127) & vbCrLf & vbCrLf & omax.MaximaOutput, vbOKCancel, Sprog.Warning)
1630              If resultat = vbCancel Then GoTo Slut
1640          End If

1650          variabel = omax.ConvertToWordSymbols(variabel)
              
              Dim Sep As String
      '        If CASengine = 0 Then
1660              Sep = VBA.ChrW(8744)
      '        Else
      '            sep = ";"
      '        End If
1670          If omax.StopNow Or (omax.IsAllSolved(omax.MaximaOutput, variabel, Sep) = "false" And Not (InStr(variabel, "+") > 0)) Then
1680              IsSolved = False
1690          Else
1700              IsSolved = True
1710          End If

              ' indsæt forklaring hvis ønsket
1720          If MaximaForklaring And (IsSolved Or InStr(omax.KommentarOutput, "solving system of equations")) Then
                  '            InsertForklaring "Ligningen løses for " & variabel & " vha. CAS-værktøjet " & AppNavn & ". "
1730              InsertForklaring Sprog.EquationSolvedFor & variabel & Sprog.ByCAS
1740          End If

1750          fejlm = TranslateComment(omax.KommentarOutput)

1760          If Len(fejlm) > 0 Then
1770              fejlm = Sprog.A(128) & vbCrLf & fejlm
1780              MsgBox fejlm & vbCrLf & vbCrLf & omax.KommentarOutput & vbCrLf & MaxProc.LastMaximaOutput, vbOKOnly, Sprog.Error
1790              RestartMaxima
1800          ElseIf InStr(omax.MaximaOutput, VBA.ChrW(8709)) Then    ' ingen løsning
1810              omax.InsertMaximaOutput
1820              Selection.TypeParagraph
1830              If MaximaComplex Then
1840                  Selection.TypeText Sprog.A(129) & variabel
1850              Else
1860                  Selection.TypeText Sprog.A(130) & variabel
1870              End If
1880          ElseIf InStr(omax.MaximaOutput, VBA.ChrW(8477)) > 0 Then  ' alle er løsning
1890              omax.InsertMaximaOutput
1900              Selection.TypeParagraph
1910              Selection.TypeText Sprog.A(132)
1920          ElseIf omax.MaximaOutput = "?" Or omax.MaximaOutput = "" Or InStr(omax.KommentarOutput, "Lisp error") > 0 Or (Not LmSet And Not IsSolved) Then
1930              UserFormChooseCAS.Show
1940              If UserFormChooseCAS.ChosenCAS = 2 Then ' maxima num
1950                  CASengine = 0
1960                  GoTo stophop
1970              ElseIf UserFormChooseCAS.ChosenCAS = -1 Then
1980                  GoTo Slut
1990              ElseIf UserFormChooseCAS.ChosenCAS = 1 Then ' Maxima sym
2000                  CASengine = 0
2010                  GoTo newcas
2020              ElseIf UserFormChooseCAS.ChosenCAS = 3 Then ' geogebra sym
2030                  CASengine = 2
2040                  GoTo newcas
2050              ElseIf UserFormChooseCAS.ChosenCAS = 4 Then ' geogebra num
2060                  CASengine = 2
2070                  Selection.MoveLeft wdCharacter, 1
2080                  MaximaNsolve variabel
2090                  GoTo Slut
2100              ElseIf UserFormChooseCAS.ChosenCAS = 5 Then ' geogebra browser sym
2110                  CASengine = 1
2120                  GoTo newcas
2130              Else ' grafisk geogebra
2140                  CASengine = 1
2150                  Selection.MoveLeft wdCharacter, 1
2160                  MaximaNsolve variabel
2170                  GoTo Slut
2180              End If
2190          ElseIf False Then
2200 stophop:      If omax.AntalVars > 1 Then
2210                  UFSolvenumeric.result = ""
2220              Else
2230                  UFSolvenumeric.Label_omskrevet.Caption = omax.MaximaOutput
2240                  omax.PrepareNewCommand
2250                  omax.TempDefs = UFSelectVar.TempDefs
2260                  UFSolvenumeric.FejlMeld = omax.KommentarOutput
2270                  UFSolvenumeric.variabel = variabel
2280                  UFSolvenumeric.Show
2290              End If
2300              If UFSolvenumeric.result = "num" Then ' grafisk løsning valgt
2310                  Selection.End = sslut    ' slut skal være først eller går det galt
2320                  Selection.start = sstart
2330                  ActiveWindow.VerticalPercentScrolled = scrollpos
2340                  If GraphApp = 0 And CASengine = 0 Then
2350                      MaximaSolveNumeric UFSelectVar.ListBox_vars.text
2360                  Else
2370                      CASengine = 1
2380                      Selection.MoveLeft wdCharacter, 1
2390                      MaximaNsolve variabel
2400                      GoTo Slut
2410                  End If
2420              Else
2430                  If UFSolvenumeric.result = "nsolve" Then
2440                      InsertForklaring Sprog.EquationSolvedNumFor & variabel & Sprog.ByCAS, False
2450                      omax.MaximaOutput = UFSolvenumeric.Label_nsolve.Caption
2460                  ElseIf UFSolvenumeric.result = "omskriv" Then
2470                      InsertForklaring "", True
2480                      omax.MaximaOutput = UFSolvenumeric.Label_omskrevet.Caption
2490                  End If
2500                  omax.InsertMaximaOutput
2510                  Selection.TypeParagraph
2520                  Selection.TypeText UFSolvenumeric.Kommentar
2530              End If

2540          Else    ' hvis der er løsning
      '            If MaximaUnits Then
      '                omax.Kommando = omax.MaximaOutput
      '                omax.beregn
      '                omax.MaximaOutput = omax.MaximaOutput
      '            End If
2550              omax.InsertMaximaOutput
2560          End If

2570          Oundo.EndCustomRecord

2580      Else    '--------------- ligningssystem ----------------------

2590          omax.FindVariable
2600          If Not ValiderVariable Then GoTo Slut
2610          UFSelectVar.NoEq = omax.AntalKom
2620          UFSelectVar.Vars = omax.Vars
2630          UFSelectVar.DefS = omax.DefString
2640          UFSelectVar.Show
2650          variabel = UFSelectVar.SelectedVar

              '    variabel = InputBox("Indtast variable som ligningssystemet skal løses for adskilt af komma", "Variable", "x,y")

2660          If variabel = "" Then GoTo Slut
2670          omax.TempDefs = UFSelectVar.TempDefs
              '    UFWait.ActionToPerform = "solvesystem"
              '    UFWait.VarParam = variabel

              '    Set UFWait.omax = omax
              '    UFWait.Show
              
newcassys:
2680          If CASengine = 1 Or CASengine = 2 Then
2690              If MaximaComplex Then
2700                  s = "Csolve({" & Replace(Replace(omax.KommandoerStreng, ",", "."), ";", " , ") & "},{" & Replace(variabel, ";", " , ") & "})"
2710              Else
2720                  s = "solve({" & Replace(Replace(omax.KommandoerStreng, ",", "."), ";", " , ") & "},{" & Replace(variabel, ";", " , ") & "})"
2730              End If
2740              If MaximaDecOutType = 3 Then
2750                  s = "ScientificText(" & s & " , " & MaximaCifre & ")"
2760              ElseIf MaximaExact = 2 Then
2770                  s = "numeric(" & s & " , " & MaximaCifre & ")"
2780              End If
2790          End If
              
2800          If CASengine = 0 Then
2810              omax.SolveSystem (variabel)
2820          ElseIf CASengine = 1 Then
2830              If MaximaForklaring Then
2840                  omax.GoToEndOfSelectedMaths
2850                  If MaximaForklaring Then
2860                      Selection.TypeParagraph
2870                      InsertForklaring Sprog.A(134) & variabel & Sprog.A(135), False
2880                      Selection.TypeParagraph
2890                  End If
2900              End If
2910              OpenGeoGebraWeb s, "CAS", True, True
2920              GoTo Slut
2930          ElseIf CASengine = 2 Then
2940              Call RunGeoGebraDirect(s)
2950              If omax.MaximaOutput = "{}" Then
2960                  omax.MaximaOutput = variabel & VBA.ChrW(8712) & VBA.ChrW(8709)
2970              ElseIf omax.MaximaOutput = "{" & variabel & "=" & variabel & "}" Or omax.MaximaOutput = "{x=x}" Then
2980                  omax.MaximaOutput = variabel & VBA.ChrW(8712) & VBA.ChrW(8477)
2990              ElseIf omax.MaximaOutput <> "?" And omax.MaximaOutput <> "" Then
3000                  omax.ConvertOutputToSolvedGGB ' burde måske forbedres med ea.getnextlistitem ligesom solve
3010              End If
3020              Application.Activate
3030          End If
              
3040          If omax.StopNow Then GoTo Slut
3050          If CheckForError Then
3060              scrollpos = ActiveWindow.VerticalPercentScrolled
3070              sslut = Selection.End
3080              sstart = Selection.start
3090              GoTo Slut
3100          End If
3110          If omax.StopNow Then GoTo Slut
3120          Application.ScreenUpdating = False
              
3130          Set Oundo = Application.UndoRecord
3140          Oundo.StartCustomRecord
              
              '    omax.KommentarOutput = TranslateReplaceComment(omax.KommentarOutput)

3150          If omax.MaximaOutput = "?merror(""Anumberwasfoundwhereavariablewasexpected-`solve'"")" Then
3160              MsgBox Sprog.A(133), vbOKOnly, Sprog.Error
3170              GoTo Slut
3180          End If

3190          If omax.DefFejl Then GoTo Slut

3200          variabel = Replace(omax.ConvertToWordSymbols(variabel), ";", ",")
              
3210          If omax.MaximaOutput = "?" Or omax.MaximaOutput = "" Or InStr(omax.KommentarOutput, "Lisp error") > 0 Then
3220              UserFormChooseCAS.Show
3230              If UserFormChooseCAS.ChosenCAS = 2 Then ' maxima num
3240                  GoTo stophop
3250              ElseIf UserFormChooseCAS.ChosenCAS = -1 Then
3260                  GoTo Slut
3270              ElseIf UserFormChooseCAS.ChosenCAS = 1 Then ' Maxima sym
3280                  CASengine = 0
3290                  GoTo newcassys
3300              ElseIf UserFormChooseCAS.ChosenCAS = 3 Then ' geogebra sym
3310                  CASengine = 2
3320                  GoTo newcassys
3330              ElseIf UserFormChooseCAS.ChosenCAS = 4 Then ' geogebra num
3340                  TempCas = CASengine
3350                  CASengine = 2
      '                Selection.MoveLeft wdCharacter, 1
3360                  MaximaNsolve variabel
3370                  GoTo Slut
3380                  CASengine = TempCas
3390              ElseIf UserFormChooseCAS.ChosenCAS = 5 Then ' geogebra browser sym
3400                  CASengine = 1
3410                  GoTo newcassys
3420              Else ' grafisk geogebra
3430                  TempCas = CASengine
3440                  CASengine = 1
3450                  Selection.MoveLeft wdCharacter, 1
3460                  MaximaNsolve variabel
3470                  CASengine = TempCas
3480                  GoTo Slut
3490              End If
3500          ElseIf Len(omax.MaximaOutput) > 1 Then
3510              omax.GoToEndOfSelectedMaths
3520              Selection.TypeParagraph
3530              If MaximaForklaring Then
3540                  InsertForklaring Sprog.A(134) & variabel & Sprog.A(135)
3550              End If
3560              omax.InsertMaximaOutput
3570              Arr = Split(omax.MaximaOutput, "=")
3580              If UBound(Arr) = 1 Then
3590                  If InStr(Arr(0), variabel) > 0 And InStr(Arr(1), variabel) > 0 Then
                          '                    Result = MsgBox("Maxima kunne ikke løse ligningssystemet. Den var for kompleks." & vbCrLf & vbCrLf & omax.KommentarOutput & vbCrLf & vbCrLf & "Tryk OK hvis du vil forsøge at løse ligningen numerisk.", vbOKCancel, "Fejl")
3600                      UFSolvenumeric.FejlMeld = omax.KommentarOutput
3610                      UFSolvenumeric.Show
3620                      If UFSolvenumeric.result = "num" Then
3630                          Selection.End = sslut    ' slut skal være først eller går det galt
3640                          Selection.start = sstart
3650                          ActiveWindow.VerticalPercentScrolled = scrollpos
3660                          MaximaSolveNumeric
3670                      Else
3680                          Selection.TypeParagraph
3690                          Selection.TypeText Sprog.A(136)
3700                      End If
3710                  End If
3720              ElseIf InStr(omax.MaximaOutput, VBA.ChrW(8709)) > 0 And CASengine > 0 Then
3730                  Selection.TypeParagraph
3740                  Selection.TypeText "GeoGebra har været anvendt til at løse ligningssystemet. Det er usikkert om der kan være løsninger. Det anbefales at forsøge med anden metode. Fx Maxima, eller numerisk/grafisk"
3750              End If
3760          Else    ' ingen løsninger
3770              omax.GoToEndOfSelectedMaths
3780              Selection.TypeParagraph
3790              If Len(omax.KommentarOutput) <= 1 Then
3800                  omax.MaximaOutput = "L=" & VBA.ChrW(8709)
                      '               omax.GoToEndOfSelectedMaths
3810                  omax.InsertMaximaOutput
3820                  Selection.TypeParagraph
3830                  If MaximaComplex Then
3840                      Selection.TypeText Sprog.A(137) & variabel
3850                  Else
3860                      Selection.TypeText Sprog.A(138) & variabel
3870                  End If
3880              Else
3890                  If omax.DefFejl Then GoTo Slut
      '                    MsgBox Sprog.DefError & vbCrLf & VisDef & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
      '                Else
3900                      fejlm = Sprog.A(131) & vbCrLf
      '                End If
3910                  If InStr(omax.Kommando, "=") < 1 Then
3920                      fejlm = fejlm & Sprog.A(139) & vbCrLf
3930                  End If
                      '                fejlm = fejlm & TranslateComment(omax.KommentarOutput)
3940                  UFSolvenumeric.FejlMeld = omax.KommentarOutput
3950                  UFSolvenumeric.Show
3960                  If UFSolvenumeric.result = "num" Then
3970                      Selection.End = sslut    ' slut skal være først eller går det galt
3980                      Selection.start = sstart
3990                      ActiveWindow.VerticalPercentScrolled = scrollpos
4000                      MaximaSolveNumeric
4010                  Else
4020                      Selection.TypeParagraph
4030                      Selection.TypeText Sprog.A(140)
4040                  End If
                      '                MsgBox fejlm & "Angav du de rigtige variable?" & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, "Fejl"
4050              End If
4060          End If
4070          Oundo.EndCustomRecord
4080      End If
4090      GoTo Slut
Fejl:
4100      MsgBox Sprog.ErrorGeneral & vbCrLf & "Err. no: " & Err.Number & vbCrLf & Err.Description & vbCrLf & "Line number: " & Erl, vbOKOnly, Sprog.Error
4110      RestartMaxima
Slut:
          '    omax.Luk
4120      On Error Resume Next
4130      MaximaCifre = SaveSettingsCifre
4140      MaximaExact = SaveSettingsExact
4150      OutUnits = SaveSettingsOutunits
4160      MaximaLogOutput = SaveSettingsLog
4170      MaximaDecOutType = SaveSettingsDecOutType
4180      CASengine = TempCas
4190      Selection.End = sslut    ' slut skal være først eller går det galt
4200      Selection.start = sstart
4210      ActiveWindow.VerticalPercentScrolled = scrollpos
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
    On Error GoTo Fejl
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
    If Not omax.MaximaInstalled Then GoTo Slut
    '    If UFWait Is Nothing Then Set UFWait = New UserFormWaitForMaxima
    '    Set UFWait = New UserFormWaitForMaxima
    sstart = Selection.start
    sslut = Selection.End

    If omax.DefFejl Then GoTo Slut

    If Selection.OMaths.Count = 0 And Len(Selection.Range.text) < 2 Then
        MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
        GoTo Slut
    End If
    If sstart = sslut Then
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    If InStr(Selection.OMaths(1).Range.text, "=") < 1 Then
        Dim result As VbMsgBoxResult
        result = MsgBox(Sprog.A(141), vbYesNo, Sprog.Warning)
        If result = vbNo Then GoTo Slut
    End If

    If Selection.OMaths.Count < 2 And InStr(Selection.OMaths(1).Range.text, VBA.ChrW(8743)) < 1 Then
        GoTo Slut
    Else    ' ligningssystem

        omax.ReadSelection
        omax.FindVariable
        If Not ValiderVariable Then GoTo Slut
        UFSelectVar.Eliminate = True
        UFSelectVar.NoEq = omax.AntalKom
        UFSelectVar.Vars = omax.Vars
        UFSelectVar.DefS = omax.DefString
        UFSelectVar.Caption = Sprog.A(377)
        UFSelectVar.Show
        variabel = UFSelectVar.SelectedVar

        '    variabel = InputBox("Indtast variable som ligningssystemet skal løses for adskilt af komma", "Variable", "x,y")

        If variabel = "" Then GoTo Slut
        omax.TempDefs = UFSelectVar.TempDefs
        '    UFWait.ActionToPerform = "solvesystem"
        '    UFWait.VarParam = variabel
        '    Set UFWait.omax = omax
        '    UFWait.Show
        
        
        
    If CASengine = 0 Then
        omax.Eliminate (variabel)
    Else
        MsgBox "This function is not supported by GeoGebra. Try equation solver", vbOKOnly, "No support"
        GoTo Slut
    End If
        
        
        If omax.StopNow Then GoTo Slut
        If CheckForError Then
            scrollpos = ActiveWindow.VerticalPercentScrolled
            sslut = Selection.End
            sstart = Selection.start
            GoTo Slut
        End If
        
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
        
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph
        If omax.StopNow Then GoTo Slut
        Application.ScreenUpdating = False
        '    omax.KommentarOutput = TranslateReplaceComment(omax.KommentarOutput)

        If omax.MaximaOutput = "?merror(""Anumberwasfoundwhereavariablewasexpected-`solve'"")" Then
            MsgBox Sprog.A(133), vbOKOnly, Sprog.Error
            GoTo Slut
        End If

        If omax.DefFejl Then GoTo Slut
'            MsgBox Sprog.DefError & vbCrLf & VisDef & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
'            GoTo slut
'        End If

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
                If omax.DefFejl Then GoTo Slut
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
Sub MaximaNsolve(Optional ByVal variabel As String)
    Dim Arr As Variant
    Dim fejlm As String
    Dim UFnsolve As New UserFormNumericQuestion
        On Error GoTo Fejl
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
    UFSelectVar.NoEq = 1
'    Set UFSelectVars = New UserFormSelectVars
'    Dim variabel As String
    Dim sstart As Long, sslut As Long
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo Slut
    '    If UFWait Is Nothing Then Set UFWait = New UserFormWaitForMaxima
    '    Set UFWait = New UserFormWaitForMaxima
    sstart = Selection.start
    sslut = Selection.End

    If omax.DefFejl Then GoTo Slut

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
        Dim result As VbMsgBoxResult
        result = MsgBox(Sprog.A(141), vbYesNo, Sprog.Warning)
        If result = vbNo Then GoTo Slut
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
        
        If CASengine > 0 And Not AllTrig Then ' På geogebra skal der via vba genkendes om det er trigonometrisk ligning
            If Not InStr(omax.Vars, ";") > 0 Then ' metoden virker kun med 1 variabel
                ea.SetNormalBrackets
                ea.text = omax.Kommando
                ea.text = Replace(ea.text, VBA.ChrW(8289), "")
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
                    UFSelectVar.TextBox_def.text = "0<=" & omax.Vars & "<=" & s & VbCrLfMac
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
        
        If variabel = "" Then GoTo Slut
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
            GoTo Slut
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

        If omax.StopNow Then GoTo Slut
        If CheckForError Then
            scrollpos = ActiveWindow.VerticalPercentScrolled
            sslut = Selection.End
            sstart = Selection.start
            GoTo Slut
        End If
        
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord

        If UFnsolve.result = "afbryd" Then GoTo Slut

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
            GoTo Slut
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
        If inp = "" Then GoTo Slut    ' trykket cancel
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

        If Variable = "" Then GoTo Slut

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
            GoTo Slut
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
            If omax.DefFejl Then GoTo Slut
            fejlm = Sprog.A(383) & "." & vbCrLf
            If InStr(omax.Kommando, "=") < 1 Then
                fejlm = fejlm & Sprog.A(139) & "." & vbCrLf
            End If
            fejlm = fejlm & TranslateComment(omax.KommentarOutput)
            MsgBox fejlm & Sprog.A(384) & vbCrLf & vbCrLf & omax.KommentarOutput, vbOKOnly, Sprog.Error
        End If
    End If
    
    Oundo.EndCustomRecord

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
Sub MaximaSolveNumeric(Optional Var As String)
' grafisk løsning + newton + intervalmetode
' var indsættes som den variabel der løses for
    On Error GoTo Fejl
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

    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo Slut
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
        GoTo Slut
    End If
    If sstart = sslut Then
        Selection.OMaths(1).ParentOMath.Range.Select
    End If
    If InStr(Selection.OMaths(1).Range.text, "=") < 1 Then
        Dim result As VbMsgBoxResult
        result = MsgBox(Sprog.A(141), vbYesNo, Sprog.Warning)
        If result = vbNo Then GoTo Slut
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
        GoTo Slut
    End If
        omax.FindVariable
        omax.TempDefs = UFSelectVar.TempDefs

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
        If CheckForError Then
            scrollpos = ActiveWindow.VerticalPercentScrolled
            sslut = Selection.End
            sstart = Selection.start
            GoTo Slut
        End If
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
            If omax.DefFejl Then GoTo Slut
            fejlm = Sprog.A(388) & "." & vbCrLf
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
10            On Error GoTo Fejl
         ' Application.ScreenUpdating = False
          '   LockWindow
          Dim tid As Single
20        tid = Timer
#If Mac Then
          Dim D As Document
30        Set D = ActiveDocument
#Else
40        system.Cursor = wdCursorWait
#End If
          Dim scrollpos As Double
          Dim sstart As Long, sslut As Long
50        sstart = Selection.start
60        sslut = Selection.End
          '    TimeText = ""
          '    Dim st As Double
          '    st = Timer
70        scrollpos = ActiveWindow.VerticalPercentScrolled
          '    Set UFWait = New UserFormWaitForMaxima
          
80        If SettUseVBACAS And (MaximaExact = 2 And Not MaximaUnits And Not MaximaComplex) Then
90            On Error Resume Next
100           Err.Clear
110           Application.Run macroname:="CASCALC"
120           If Err.Number = 513 Then
130               MsgBox2 Err.Description, vbOKOnly, Sprog.Error
140               GoTo Slut
150           ElseIf Err.Number = 0 Then ' hvis ingen fejl, så er beregningen done og sat ind i Word
160               GoTo Slut
170           End If
180           Err.Clear
190           On Error GoTo Fejl
200       End If
          
          
210       If Not PrepareMaxima Then
220           If omax.DefFejl Then
230               Exit Sub
240           Else
250               GoTo Slut
260           End If
270       End If
280       omax.prevspr = ""

290       If CASengine = 0 And Not omax.MaximaInstalled Then GoTo Slut
300       If Selection.OMaths.Count = 0 Then  'And Len(Selection.Range.text) < 2
310           MsgBox Sprog.A(47), vbOKOnly, Sprog.Error
320           GoTo Slut
330       End If
340       If Selection.OMaths.Count > 1 Then
350           MsgBox Sprog.A(149), vbOKOnly, Sprog.Error
360           GoTo Slut
370       End If

380       omax.ReadSelection
          
390       If InStr(omax.Kommando, VBA.ChrW(8788)) > 0 Or InStr(VBA.LCase(omax.Kommando), "definer:") > 0 Or InStr(VBA.LCase(omax.Kommando), "define:") > 0 Or InStr(VBA.LCase(omax.Kommando), "definer ligning:") > 0 Or InStr(omax.Kommando, VBA.ChrW(8801)) > 0 Then  ' kun se på felter med := defligmed og := symbol
400           MsgBox Sprog.A(48), vbOKOnly, Sprog.Error
410           GoTo Slut
420       End If
430       If omax.Kommando = "" Then GoTo Slut
440       If Not ValidateInput(omax.Kommando) Then
450           If omax.DefFejl Then
460               Exit Sub
470           Else
480               GoTo Slut
490           End If
500       End If
          '    Set UFWait.omax = omax
          '    UFWait.ActionToPerform = "beregn"
          '    UFWait.Show
          '    If omax.StopNow Then GoTo slut
          
          Dim s As String, t As String, fo As String
          
510       If CASengine > 0 Then
520           s = Trim(omax.Kommando)
              '        If Left(s, 1) = "=" Then s = Left(s, Len(s) - 1)
530           s = GetCmdAfterEqualSign(s)
540           If MaximaDecOutType = 3 Then
550               s = "ScientificText(" & s & " , " & MaximaCifre & ")"
560           ElseIf MaximaExact = 2 Then
570               s = "numeric(" & s & " , " & MaximaCifre & ")"
580           End If
590       End If
          
600       If CASengine = 0 Then
610           omax.beregn
620       ElseIf CASengine = 1 Then
630           If MaximaForklaring Then
640               omax.GoToEndOfSelectedMaths
650               If MaximaForklaring Then
660                   Selection.TypeParagraph
670                   InsertForklaring Sprog.A(682), False
680                   Selection.TypeParagraph
690               End If
700           End If
710           OpenGeoGebraWeb s, "CAS", True, True
720           GoTo Slut
730       ElseIf CASengine = 2 Then
740           fo = RunGeoGebraDirect(s)
750           If MaximaExact = 0 And MaximaDecOutType < 3 Then
760               If fo = "?" Or fo = "null" Or fo = "" Then
770                   s = "numeric(" & s & " , " & MaximaCifre & ")"
780               Else
                      ' det første resultat kan ikke bare fødes ind i GeoGebra igen. Det giver problemer i særlige tilfælde. Eksempel: '\cbrt(79/138)^(2)' Her burde være parentes. Den rigtige fortolkning er cbrt((79/138)^2), som kommer frem hvis den tastes og læses i Word. Hvis den køres direkte i WordMat oversætteren indsættes ikke korrekt parentes. Normalt ikke et problem, da alt normalt læses fra Word
                      '                fo = omax.ReadFromWord(fo) ' forsøg på at omgå problem med at føde resulkat direkte ind i geogebra igen
                      '                s = "numeric(" & fo & " , " & MaximaCifre & ")"
790                   s = "numeric(" & s & " , " & MaximaCifre & ")" ' Der er eksempler, hvor det er bedre at beregne numerisk videre på det eksakte resultat istedet for direkte på det originale, men pga ovenstående problematik
800               End If
810               MaximaExact = 2
820               t = RunGeoGebraDirect(s)
830               MaximaExact = 0
840               If (fo = "?" Or fo = "null" Or fo = "") And (t = "?" Or t = "null" Or t = "") Then
850                   omax.MaximaOutput = fo
860               ElseIf (fo = "?" Or fo = "null" Or fo = "") Then
870                   omax.MaximaOutput = t
880               ElseIf (t = "?" Or t = "null" Or t = "") Or fo = t Or t = "" Then
890                   omax.MaximaOutput = fo
900               Else
910                   omax.MaximaOutput = fo & VBA.ChrW(&H2248) & t
920               End If
930           End If
940           omax.MaximaOutput = "=" & omax.MaximaOutput
950       End If
          
960       If omax.StopNow Then GoTo Slut
970       If CheckForError Then Exit Sub ' der skal ikke scrolles ved fejl, deffejl skal markeres
          '    TimeText = TimeText & vbCrLf & "beregn: " & Timer - st

          Dim Oundo As UndoRecord
980       Set Oundo = Application.UndoRecord
990       Oundo.StartCustomRecord
          
1000      If Len(omax.MaximaOutput) > 0 And Trim(omax.MaximaOutput) <> "=res1" Then
1010          InsertOutput omax.MaximaOutput, False
1020      Else
1030          fejlm = Sprog.A(128) & vbCrLf
1040          fejlm = fejlm & TranslateComment(omax.KommentarOutput)
              '        If InStr(omax.KommentarOutput, "infix") > 0 Then
              '            fejlm = fejlm & "Husk at alle gangetegn skal laves. 2*x ikke 2x" & vbCrLf
              '        End If
#If Mac Then
1050          fejlm = fejlm & vbCrLf & vbCrLf & omax.KommentarOutput & vbCrLf
1060          MsgBox2 fejlm, vbOKOnly, Sprog.Error
#Else
1070          fejlm = fejlm & vbCrLf & vbCrLf & omax.KommentarOutput & vbCrLf & MaxProc.LastMaximaOutput
1080          MsgBox2 fejlm, vbOKOnly, Sprog.Error
1090          RestartMaxima
#End If
1100      End If
1110      Oundo.EndCustomRecord

1120      GoTo Slut
Fejl:
1130      MsgBox Sprog.ErrorGeneral & vbCrLf & "Err. no: " & Err.Number & vbCrLf & Err.Description & vbCrLf & "Line number: " & Erl, vbOKOnly, Sprog.Error
1140      RestartMaxima
Slut:
#If Mac Then
1150      D.Activate
#End If
1160      On Error Resume Next
1170      If ActiveWindow.VerticalPercentScrolled <> scrollpos Then ActiveWindow.VerticalPercentScrolled = scrollpos
1180      Application.ScreenUpdating = True
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
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo Slut
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
    UFomskriv.Vars = omax.Vars
    UFomskriv.Show
    If UFomskriv.annuller Then GoTo Slut

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
        GoTo Slut
    ElseIf CASengine = 2 Then
        omax.MaximaOutput = RunGeoGebraDirect(s)
    End If
    
    
    If omax.StopNow Then GoTo Slut
        If CheckForError Then
            scrollpos = ActiveWindow.VerticalPercentScrolled
            sslut = Selection.End
            sstart = Selection.start
            GoTo Slut
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

    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    RestartMaxima
Slut:
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
    On Error GoTo Fejl
    PrepareMaxima
    omax.prevspr = ""
    Dim scrollpos As Double, s As String
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo Slut
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
        GoTo Slut
    ElseIf CASengine = 2 Then
        omax.MaximaOutput = RunGeoGebraDirect(s)
    End If
    
    
    If omax.StopNow Then GoTo Slut
        If CheckForError Then
            scrollpos = ActiveWindow.VerticalPercentScrolled
            sslut = Selection.End
            sstart = Selection.start
            GoTo Slut
        End If
    
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
    If CheckForError Then
        scrollpos = ActiveWindow.VerticalPercentScrolled
        sslut = Selection.End
        sstart = Selection.start
        GoTo Slut
    End If
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
    Dim scrollpos As Double, s As String
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
        GoTo Slut
    ElseIf CASengine = 2 Then
        omax.MaximaOutput = RunGeoGebraDirect(s)
    End If
    

    '    omax.OpenCmd
    If omax.StopNow Then GoTo Slut
    If CheckForError Then
        scrollpos = ActiveWindow.VerticalPercentScrolled
        sslut = Selection.End
        sstart = Selection.start
        GoTo Slut
    End If
    
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
    Dim scrollpos As Double, s As String
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
        GoTo Slut
    ElseIf CASengine = 2 Then
        omax.MaximaOutput = RunGeoGebraDirect(s)
    End If

    If omax.StopNow Then GoTo Slut
    If CheckForError Then
        scrollpos = ActiveWindow.VerticalPercentScrolled
        sslut = Selection.End
        sstart = Selection.start
        GoTo Slut
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
    Dim variabel As String, s As String
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
        GoTo Slut
    ElseIf CASengine = 2 Then
        omax.MaximaOutput = RunGeoGebraDirect(s)
    End If
    
    
    If omax.StopNow Then GoTo Slut
    If CheckForError Then
        scrollpos = ActiveWindow.VerticalPercentScrolled
        sslut = Selection.End
        sstart = Selection.start
        GoTo Slut
    End If
    
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
    Dim variabel As String, s As String
    Dim scrollpos As Double
    Dim sstart As Long, sslut As Long
    sstart = Selection.start
    sslut = Selection.End
    scrollpos = ActiveWindow.VerticalPercentScrolled
    '    MsgBox WordWindowNavn
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo Slut

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
        GoTo Slut
    ElseIf CASengine = 2 Then
        omax.MaximaOutput = RunGeoGebraDirect(s)
    End If
    
    
    If omax.StopNow Then GoTo Slut
    If CheckForError Then
        scrollpos = ActiveWindow.VerticalPercentScrolled
        sslut = Selection.End
        sstart = Selection.start
        GoTo Slut
    End If
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
    Dim variabel As String
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
    If InStr(omax.Vars, "t") > 0 Then
        variabel = "t"
    ElseIf InStr(omax.Vars, "x") > 0 Then
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
    '    MsgBox WordWindowNavn
    If CASengine = 0 And Not omax.MaximaInstalled Then GoTo Slut

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
            ea.text = omax.Kommando
            funktion = ea.GetNextVar(1)
        End If
        UFdiffeq.Vars = omax.Vars
        UFdiffeq.DefS = omax.DefString
        UFdiffeq.TextBox_funktion.text = funktion
        UFdiffeq.TextBox_variabel.text = variabel
        UFdiffeq.Label_ligning.Caption = Replace(omax.ConvertToAscii(omax.Kommando), " ", "")
        UFdiffeq.Show
        If UFdiffeq.luk Then GoTo Slut
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
    If variabel = "" Then GoTo Slut
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
        GoTo Slut
    ElseIf CASengine = 2 Then
            Call RunGeoGebraDirect(s)
            omax.MaximaOutput = Replace(omax.MaximaOutput, "{", "")
            omax.MaximaOutput = Replace(omax.MaximaOutput, "}", " ") ' mellemrummet sikrer afstand så det bliver gange efter c_1
'            If Left(omax.MaximaOutput, 1) = "{" Then omax.MaximaOutput = Mid(omax.MaximaOutput, 2, Len(omax.MaximaOutput) - 2)
            omax.MaximaOutput = Replace(omax.MaximaOutput, ";", "    " & VBA.ChrW(8744) & "    ")    ' komma erstattes med eller
            Application.Activate
    End If
    
        
    If omax.StopNow Then GoTo Slut
    If CheckForError Then
        scrollpos = ActiveWindow.VerticalPercentScrolled
        sslut = Selection.End
        sstart = Selection.start
        GoTo Slut
    End If
    
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
Sub InsertOutput(text As String, Optional ResultAfterTable As Boolean = True)
' indsætter text efter den mathobj cursoren er i på sikker måde
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

Function ValidateInput(Expr, Optional MathObj As OMath) As Boolean
    Dim ED As ErrorDefinition
   
    ValidateInput = True
    '    ' validate brackets
   
    If GetCountOfChar(Expr, "(") <> GetCountOfChar(Expr, ")") Then
        ED.Title = Sprog.SyntaxError
            ED.MaximaOutput = Expr
        If Sprog.SprogNr = 1 Then
'            MsgBox "Antallet af parenteser passer ikke i udtrykket" & vbCrLf & vbCrLf & Expr, vbOKOnly, "Syntaks fejl"
            ED.Description = "Antallet af parenteser passer ikke i udtrykket" & vbCrLf & Expr
        Else
'            MsgBox "The number of brackets do not match in" & vbCrLf & vbCrLf & Expr, vbOKOnly, "Syntax error"
            ED.Description = "The number of brackets do not match" & vbCrLf & Expr
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
        If Not MathObj Is Nothing Then
            ED.Description = ED.Title & vbCrLf & ED.Description
            ED.Title = "Fejl i definition"
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
' finder versionsnr ud fra mappenavn på maxima
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
' Finder Maximastien. søger i Appdata og programfiles. Hvis der er Maxima begge steder så bruges nyeste version
' Hvis der er samme version begge steder, så returneres stien til appdata.

    Dim FN As String, DN As String, v As Single ' Til AppData
    Dim FN1 As String, DN1 As String, V1 As Single ' Til Program files
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
    Dim Arr() As String, i As Integer, s As String, ED As ErrorDefinition
'    On Error Resume Next
    ValiderVariable = True
    If omax.AntalVars > 0 Then
        Arr = Split(omax.Vars, ";")
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
