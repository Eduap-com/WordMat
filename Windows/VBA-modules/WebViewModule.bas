Attribute VB_Name = "WebViewModule"
Public WebV As Object ' WebViewWrap.Browser
Public GeoGebraWindowOpen As Boolean
Sub OpenWebV()
    Set UfWait2 = New UserFormWaitForMaxima
    UfWait2.Label_tip.Font.Size = 10
    UfWait2.Label_tip.Font.Italic = False
    UfWait2.Show vbModeless
    UfWait2.Label_tip.Caption = Sprog.A(684)
    UfWait2.Label_progress.Caption = "**"
    
    Set WebV = CreateObject("WebViewWrap.Browser")
    WebV.navigate "file://" & GetProgramFilesDir & "/WordMat/geogebra-math-apps/GeoGebraCASApplet.html"
    WebV.WaitWV
    
    Wait (2)
    
    Dim JS As String, i As Integer, j As Integer

    JS = "ggbApplet.initCAS();2+3;ggbApplet.evalCommandCAS('2+3');"
'    JS = "2+3;"
    UfWait2.Label_progress.Caption = UfWait2.Label_progress.Caption & "*"
hop:
'    Res = WebV.ExecuteScriptWaitTime(JS, 15000)
    WebV.ExecuteScriptNonBlock (JS)
    Do Until WebV.WaitUntilScriptFinished(1000) Or j >= 5
        UfWait2.Label_progress.Caption = UfWait2.Label_progress.Caption & "*"
        If omax.StopNow Then GoTo slut
        Res = WebV.GetJSReturnVal()
        If Replace(Res, """", "") = "5" Then Exit Do
        DoEvents
        j = j + 1
    Loop
    Res = WebV.GetJSReturnVal()
    If Replace(Res, """", "") <> "5" Then
        i = i + 1
        If i < 2 Then
            Wait (2)
            UfWait2.Label_progress.Caption = UfWait2.Label_progress.Caption & "*"
            JS = "2+3;"
            GoTo hop
        Else
            Set WebV = Nothing
        End If
    End If
   
slut:
    Unload UfWait2
End Sub

Sub PrepareGeoGebraCAS()
#If Mac Then
    Dim Res As String, i As Integer
    Dim UfWait2 As UserFormWaitForMaxima
    
    If Not GeoGebraWindowOpen Then
        Set UfWait2 = New UserFormWaitForMaxima
        UfWait2.Label_tip.Font.Size = 10
        UfWait2.Label_tip.Font.Italic = False
        UfWait2.Show vbModeless
        UfWait2.Label_tip.Caption = Sprog.A(684)
        UfWait2.Label_progress.Caption = "*"
        UfWait2.Show vbModeless
        UfWait2.Label_progress.Caption = UfWait2.Label_progress.Caption & "*"
        Do
            Res = RunScript("ExecuteGeoGebraCASCommand", "ggbApplet.reset();ggbApplet.evalCommandCAS('2+3');")
            If Res <> "5" Then
                UfWait2.Label_progress.Caption = UfWait2.Label_progress.Caption & "*"
                Wait (1)
                UfWait2.Label_progress.Caption = UfWait2.Label_progress.Caption & "*"
                Wait (1)
            End If
            i = i + 1
        Loop While Res <> "5" And i < 5
        Unload UfWait2
        If Res <> "5" Then GGBJSGuidance
'        MsgBox Res
    End If
#Else
    Dim JS As String, Res As String
    If WebV Is Nothing Then OpenWebV
    Err.Clear
    On Error Resume Next
'    Res = WebV.Title
'    If Err.Number <> 0 Then OpenWebV
#End If
     GeoGebraWindowOpen = True
     
'    JS = "ggbApplet.evalCommandCAS('2+3');"
'    Res = WebV.ExecuteScript(JS)
End Sub

Sub TestWV()
    OpenWebV
    WebV.Show
    WebV.Width = 1500

'    WebV.navigate "https://www.geogebra.org"
    WebV.navigate "file://" & GetProgramFilesDir & "/WordMat/geogebra-math-apps/GeoGebraCASApplet.html"
    WebV.WaitWV
    JS = "ggbApplet.evalCommandCAS('2+3')"
    Res = WebV.ExecuteScript(JS)
    MsgBox Res
End Sub

Sub TestJSCommand()
    JS = "2+4;"
    Res = WebV.ExecuteScript(JS)
    MsgBox Res
End Sub

Function ExecuteGeoGebraCasCommand(CmdString As String, Optional UseDefs As Boolean = True) As String
Dim Res As String

    Dim JS As String, Arr() As String, ArrDef() As String, ArrCas() As String, cmd As String, s As String, i As Integer, AssumeCol As New Collection, AssumeString As String, AE As Variant
    Dim FC As Integer
    If WebV Is Nothing Then PrepareGeoGebraCAS
    JS = "ggbApplet.reset();" 'ggbApplet.setRounding(""" & MaximaCifre & "s"");"
    If UseDefs Then
        If GeoGebraDefs <> "" Then
            ArrDef = Split(GeoGebraDefs, ";")
            For i = 0 To UBound(ArrDef)
                JS = JS & "ggbApplet.evalCommand(""" & ArrDef(i) & """);"
            Next
        End If
        
        If GeoGebraAssumes <> "" Then
            AssumeString = "Assume(" & GeoGebraAssumes
        End If
    End If
    
    ArrCas = Split(CmdString, ";")
    For i = 0 To UBound(ArrCas)
        If AssumeString <> "" Then
            JS = JS & "ggbApplet.evalCommandCAS(""" & AssumeString & "," & ArrCas(i) & ")"");"
        Else
            JS = JS & "ggbApplet.evalCommandCAS(""" & ArrCas(i) & """);"
        End If
    Next
    Res = ExecuteGeogebraCmdViaJS(JS)
    If Res = "xQw6rT" Then
        Wait 1
        Res = ExecuteGeogebraCmdViaJS(JS)
        If Res = "xQw6rT" Then
            Set WebV = Nothing
            PrepareGeoGebraCAS
            JS = "ggbApplet.reset();ggbApplet.evalCommandCAS(solve(x^2=9,x))"
            Res = ExecuteGeogebraCmdViaJS(JS)
            If Res = "xQw6rT" Then
                Res = "Fejl ved GeoGebra"
            End If
        End If
    ElseIf Res = "ScriptError" Then
        GGBJSGuidance
    End If
    ExecuteGeoGebraCasCommand = Res
'    MsgBox Res
End Function
Sub GGBJSGuidance()
    If Sprog.SprogNr = 1 Then
        If MsgBox("WordMat kan ikke kommunikere med GeoGebra. Det skyldes formentlig at du ikke har sat de rigtige sikkerhedsindstillinger i Safari. Klik ok for at fŒ instruktionerne", vbOKCancel, "Error") = vbOK Then
            OpenLink "https://eduap.com/geogebra-som-cas-motor-pa-mac/"
        End If
    Else
        If MsgBox("WordMath cannot communicate with GeoGebra. This is probably because you have not set the required security settings in Safari. Click ok to get the instructions.", vbOKCancel, "Error") = vbOK Then
            OpenLink "https://eduap.com/geogebra-as-cas-engine-on-mac/"
        End If
    End If
End Sub

Function ExecuteGeogebraCmdViaJS(JS As String) As String
Dim FejlC As Integer
#If Mac Then
    ExecuteGeogebraCmdViaJS = RunScript("ExecuteGeoGebraCASCommand", JS)
#Else
 '   ExecuteGeogebraCmdViaJS = WebV.ExecuteScriptWaitTime(JS, 10000)
    Dim j As Integer
    Dim UfWait2 As UserFormWaitForMaxima
start1:
    WebV.ExecuteScriptNonBlock JS
    Do Until WebV.WaitUntilScriptFinished(200) Or j >= 75
'    Do Until WebV.GetJSReturnVal() <> "xQw6rT" Or j >= 75
        If Not UfWait2 Is Nothing Then UfWait2.Label_progress.Caption = UfWait2.Label_progress.Caption & "*"
        If omax.StopNow Then GoTo slut
'        Wait (0.2)
        j = j + 1
        If j = 10 And UfWait2 Is Nothing Then
            Set UfWait2 = New UserFormWaitForMaxima
            UfWait2.Label_tip.Font.Size = 10
            UfWait2.Label_tip.Font.Italic = False
            UfWait2.Show vbModeless
            UfWait2.Label_tip.Caption = Sprog.A(684)
            UfWait2.Label_progress.Caption = "*"
        End If
    Loop
    If j >= 75 Then
        WebV.Reload
        ExecuteGeogebraCmdViaJS = ""
    ElseIf (WebV.GetJSReturnVal() = """?""" Or WebV.GetJSReturnVal() = "null") And FejlC = 0 Then
        If UfWait2 Is Nothing Then
            Set UfWait2 = New UserFormWaitForMaxima
            UfWait2.Label_tip.Font.Size = 10
            UfWait2.Label_tip.Font.Italic = False
            UfWait2.Show vbModeless
            UfWait2.Label_tip.Caption = Sprog.A(684)
            UfWait2.Label_progress.Caption = "*"
        End If
        WebV.Reload
        WebV.WaitWV
        If Not UfWait2 Is Nothing Then UfWait2.Label_progress.Caption = UfWait2.Label_progress.Caption & "*"
        FejlC = FejlC + 1
        Wait 1
        If Not UfWait2 Is Nothing Then UfWait2.Label_progress.Caption = UfWait2.Label_progress.Caption & "*"
        Wait 1
        If Not UfWait2 Is Nothing Then UfWait2.Label_progress.Caption = UfWait2.Label_progress.Caption & "*"
        GoTo start1
'    ElseIf (WebV.GetJSReturnVal() = """?""" Or WebV.GetJSReturnVal() = "null") And FejlC = 1 Then
'        Wait 2
'        FejlC = FejlC + 1
'        GoTo start1
    Else
        ExecuteGeogebraCmdViaJS = WebV.GetJSReturnVal()
    End If
slut:
    If Not UfWait2 Is Nothing Then Unload UfWait2
#End If
End Function


Sub WaitForGeoGebra()
#If Mac Then
#Else
    WebV.WaitWV
#End If
End Sub
