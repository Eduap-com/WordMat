Attribute VB_Name = "WebViewModule"
Public WebV As Object ' WebViewWrap.Browser
Public GeoGebraWindowOpen As Boolean

Sub OpenWebV()
    Set ufwait2 = New UserFormWaitForMaxima
    ufwait2.Label_tip.Font.Size = 10
    ufwait2.Label_tip.Font.Italic = False
    ufwait2.Show vbModeless
    ufwait2.Label_tip.Caption = Sprog.A(684)
    ufwait2.Label_progress.Caption = "**"
    
    Set WebV = CreateObject("WebViewWrap.Browser")
    WebV.navigate "file://" & GetProgramFilesDir & "/WordMat/geogebra-math-apps/GeoGebraCASApplet.html"
    WebV.WaitWV
    
    Dim JS As String, i As Integer

    JS = "ggbApplet.initCAS();2+3;"
'    JS = "2+3;"
hop:
    Res = WebV.ExecuteScript(JS)
    ufwait2.Label_progress.Caption = ufwait2.Label_progress.Caption & "*"
    If Res <> "5" Then
        i = i + 1
        If i < 10 Then
            Wait (1)
            ufwait2.Label_progress.Caption = ufwait2.Label_progress.Caption & "*"
            JS = "2+3;"
            GoTo hop
        End If
    End If
   
    Unload ufwait2
End Sub

Sub PrepareGeoGebraCAS()
#If Mac Then
    If Not GeoGebraWindowOpen Then RunScript "ExecuteGeoGebraCASCommand", "ggbApplet.reset();ggbApplet.evalCommandCAS('2+3');"
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

    Dim JS As String, Arr() As String, ArrDef() As String, ArrCas() As String, Cmd As String, s As String, i As Integer, AssumeCol As New Collection, AssumeString As String, AE As Variant
    Dim FC As Integer
    If WebV Is Nothing Then PrepareGeoGebraCAS
    JS = "ggbApplet.reset();"
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
    End If
    ExecuteGeoGebraCasCommand = Res
'    MsgBox Res
End Function

Function ExecuteGeogebraCmdViaJS(JS As String) As String
#If Mac Then
    ExecuteGeogebraCmdViaJS = RunScript("ExecuteGeoGebraCASCommand", JS)
#Else
    ExecuteGeogebraCmdViaJS = WebV.ExecuteScript(JS)
#End If
End Function


Sub WaitForGeoGebra()
#If Mac Then
#Else
    WebV.WaitWV
#End If
End Sub
