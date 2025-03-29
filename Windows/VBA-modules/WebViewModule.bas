Attribute VB_Name = "WebViewModule"
Public WebV As Object ' WebViewWrap.Browser
Public GeoGebraWindowOpen As Boolean

Sub OpenWebV()
#If Mac Then
#Else
    Dim DN As String, FN As String
    Set UFwait2 = New UserFormWaitForMaxima
    UFwait2.Label_tip.Font.Size = 10
    UFwait2.Label_tip.Font.Italic = False
    UFwait2.Show vbModeless
    UFwait2.Label_tip.Caption = Sprog.A(684)
    UFwait2.Label_progress.Caption = "**"
    
    Err.Clear
    If DllConnType = 0 Then
        On Error Resume Next
        Set WebV = CreateObject("WebViewWrap.Browser")
        On Error GoTo 0
        Err.Clear
    Else
        Set WebV = CreateObject("WebViewWrap.Browser")
    End If
'    If WebV = Nothing Then
'        MsgBox2 "Error"
'    End If
    
    DN = GetGeoGebraMathAppsFolder()
    If DN = vbNullString Then
        MsgBox "geogebra-math-apps could not be located"
        GoTo slut
    End If
    FN = DN & "GeoGebraCASApplet.html"
    If Dir(FN) = vbNullString Then
        MsgBox "geogebra-math-apps/GeoGebraCASApplet.html could not be located"
        GoTo slut
    End If
    
    WebV.navigate "file://" & FN
    WebV.WaitWV
    
    Wait (2)
    
    Dim JS As String, i As Integer, j As Integer

    JS = "ggbApplet.initCAS();2+3;ggbApplet.evalCommandCAS('2+3');"
    '    JS = "2+3;"
    UFwait2.Label_progress.Caption = UFwait2.Label_progress.Caption & "*"
hop:
    '    Res = WebV.ExecuteScriptWaitTime(JS, 15000)
    WebV.ExecuteScriptNonBlock (JS)
    Do Until WebV.WaitUntilScriptFinished(1000) Or j >= 5
        UFwait2.Label_progress.Caption = UFwait2.Label_progress.Caption & "*"
        If omax.StopNow Then GoTo slut
        res = WebV.GetJSReturnVal()
        If Replace(res, """", "") = "5" Then Exit Do
        DoEvents
        j = j + 1
    Loop
    res = WebV.GetJSReturnVal()
    If Replace(res, """", "") <> "5" Then
        i = i + 1
        If i < 2 Then
            Wait (2)
            UFwait2.Label_progress.Caption = UFwait2.Label_progress.Caption & "*"
            JS = "2+3;"
            GoTo hop
        Else
            Set WebV = Nothing
        End If
    End If
   
slut:
    Unload UFwait2
#End If
End Sub

Sub PrepareGeoGebraCAS()
#If Mac Then
    Dim res As String, i As Integer
    Dim UFwait2 As UserFormWaitForMaxima
    
    If Not GeoGebraWindowOpen Then
        Set UFwait2 = New UserFormWaitForMaxima
        UFwait2.Label_tip.Font.Size = 10
        UFwait2.Label_tip.Font.Italic = False
        UFwait2.Show vbModeless
        UFwait2.Label_tip.Caption = Sprog.A(684)
        UFwait2.Label_progress.Caption = "*"
        UFwait2.Show vbModeless
        UFwait2.Label_progress.Caption = UFwait2.Label_progress.Caption & "*"
        Do
            res = RunScript("ExecuteGeoGebraCASCommand", "ggbApplet.reset();ggbApplet.evalCommandCAS('2+3');")
            If res <> "5" Then
                UFwait2.Label_progress.Caption = UFwait2.Label_progress.Caption & "*"
                Wait (1)
                UFwait2.Label_progress.Caption = UFwait2.Label_progress.Caption & "*"
                Wait (1)
            End If
            i = i + 1
        Loop While res <> "5" And i < 5
        Unload UFwait2
        If res <> "5" Then GGBJSGuidance
'        MsgBox Res
    End If
#Else
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
    res = WebV.ExecuteScript(JS)
    MsgBox res
End Sub

Function ExecuteGeoGebraCasCommand(CmdString As String, Optional UseDefs As Boolean = True) As String
Dim res As String

    Dim JS As String, ArrDef() As String, ArrCas() As String, i As Integer, AssumeString As String
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
    res = ExecuteGeogebraCmdViaJS(JS)
    If res = "xQw6rT" Then
        Wait 1
        res = ExecuteGeogebraCmdViaJS(JS)
        If res = "xQw6rT" Then
            Set WebV = Nothing
            PrepareGeoGebraCAS
            JS = "ggbApplet.reset();ggbApplet.evalCommandCAS(solve(x^2=9,x))"
            res = ExecuteGeogebraCmdViaJS(JS)
            If res = "xQw6rT" Then
                res = "Fejl ved GeoGebra"
            End If
        End If
    ElseIf res = "ScriptError" Then
        GGBJSGuidance
    End If
    ExecuteGeoGebraCasCommand = res
'    MsgBox Res
End Function
Sub GGBJSGuidance()
    If MsgBox(Sprog.A(886), vbOKCancel, Sprog.Error) = vbOK Then
        OpenLink "https://eduap.com/geogebra-som-cas-motor-pa-mac/"
    End If
End Sub

Function ExecuteGeogebraCmdViaJS(JS As String) As String
Dim FejlC As Integer
#If Mac Then
    ExecuteGeogebraCmdViaJS = RunScript("ExecuteGeoGebraCASCommand", JS)
#Else
 '   ExecuteGeogebraCmdViaJS = WebV.ExecuteScriptWaitTime(JS, 10000)
    Dim j As Integer
    Dim UFwait2 As UserFormWaitForMaxima
start1:
    WebV.ExecuteScriptNonBlock JS
    Do Until WebV.WaitUntilScriptFinished(200) Or j >= 75
'    Do Until WebV.GetJSReturnVal() <> "xQw6rT" Or j >= 75
        If Not UFwait2 Is Nothing Then UFwait2.Label_progress.Caption = UFwait2.Label_progress.Caption & "*"
        If omax.StopNow Then GoTo slut
'        Wait (0.2)
        j = j + 1
        If j = 10 And UFwait2 Is Nothing Then
            Set UFwait2 = New UserFormWaitForMaxima
            UFwait2.Label_tip.Font.Size = 10
            UFwait2.Label_tip.Font.Italic = False
            UFwait2.Show vbModeless
            UFwait2.Label_tip.Caption = Sprog.A(684)
            UFwait2.Label_progress.Caption = "*"
        End If
    Loop
    If j >= 75 Then
        WebV.Reload
        ExecuteGeogebraCmdViaJS = ""
    ElseIf (WebV.GetJSReturnVal() = """?""" Or WebV.GetJSReturnVal() = "null") And FejlC = 0 Then
        If UFwait2 Is Nothing Then
            Set UFwait2 = New UserFormWaitForMaxima
            UFwait2.Label_tip.Font.Size = 10
            UFwait2.Label_tip.Font.Italic = False
            UFwait2.Show vbModeless
            UFwait2.Label_tip.Caption = Sprog.A(684)
            UFwait2.Label_progress.Caption = "*"
        End If
        WebV.Reload
        WebV.WaitWV
        If Not UFwait2 Is Nothing Then UFwait2.Label_progress.Caption = UFwait2.Label_progress.Caption & "*"
        FejlC = FejlC + 1
        Wait 1
        If Not UFwait2 Is Nothing Then UFwait2.Label_progress.Caption = UFwait2.Label_progress.Caption & "*"
        Wait 1
        If Not UFwait2 Is Nothing Then UFwait2.Label_progress.Caption = UFwait2.Label_progress.Caption & "*"
        GoTo start1
'    ElseIf (WebV.GetJSReturnVal() = """?""" Or WebV.GetJSReturnVal() = "null") And FejlC = 1 Then
'        Wait 2
'        FejlC = FejlC + 1
'        GoTo start1
    Else
        ExecuteGeogebraCmdViaJS = WebV.GetJSReturnVal()
    End If
slut:
    If Not UFwait2 Is Nothing Then Unload UFwait2
#End If
End Function


Sub WaitForGeoGebra()
#If Mac Then
#Else
    WebV.WaitWV
#End If
End Sub
