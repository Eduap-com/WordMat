Attribute VB_Name = "MenuFunctions"
Option Explicit

Function VisDef() As String
Dim deftext As String
    On Error GoTo fejl
    PrepareMaxima
    deftext = omax.DefString
    
    If Len(deftext) > 3 Then
        deftext = FormatDefinitions(deftext)
        deftext = TT.A(113) & vbCrLf & vbCrLf & deftext
    Else
        deftext = TT.A(114)
    End If
    VisDef = deftext
GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
End Function

Sub DefinerFunktion()
    Dim Var As String
    On Error GoTo fejl
    Var = InputBox(TT.A(122), TT.A(123), "f(x)=x+1")
    Var = Replace(Var, ":=", "=")
    '    var = Replace(var, "=", VBA.chrw$(&H2261))
    
    If Var <> "" Then
        Var = TT.A(126) & ": " & Var
        Selection.InsertAfter (Var)
        Selection.OMaths.Add Range:=Selection.Range
        Selection.OMaths(1).BuildUp
        Selection.MoveRight unit:=wdCharacter, Count:=2
    End If
    GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
End Sub


Sub MaximaSettings()
    On Error GoTo fejl
    If UFMSettings Is Nothing Then Set UFMSettings = New UserFormSettings
    UFMSettings.Show
    GoTo slut
fejl:
    Set UFMSettings = New UserFormSettings
    UFMSettings.Show
slut:
End Sub
