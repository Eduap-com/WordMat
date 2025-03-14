Attribute VB_Name = "MenuFunktioner"
Option Explicit


Function VisDef() As String
'Dim omax As New CMaxima
Dim deftext As String
    On Error GoTo Fejl
    PrepareMaxima
    deftext = omax.DefString
    
    If Len(deftext) > 3 Then
        deftext = FormatDefinitions(deftext)
        deftext = Sprog.A(113) & vbCrLf & vbCrLf & deftext
    Else
        deftext = Sprog.A(114)
    End If
    VisDef = deftext
GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Function

Sub DefinerFunktion()
    Dim Var As String
On Error GoTo Fejl
'    var = InputBox("Indtast definitionen på den nye funktion" & vbCrLf & vbCrLf & "Definitionen kan benyttes i resten af dokumentet, men ikke før. Hvis der indsættes en clearvars: kommando længere nede i dokumentet kan den ikke benyttes derefter." & vbCrLf & vbCrLf & "Definitionen kan indtastes på 3 forskellige måder" & vbCrLf & vbCrLf & "f(x):forskrift" & vbCrLf & "f(x):=forskrift" & vbCrLf & "f(x)" & VBA.ChrW(&H2261) & "forskrift  (Definitions ligmed)" & vbCrLf & "Der kan defineres flere funktioner i en ligningsboks ved at adskille definitionerne med semikolon. f.eks. f(x)=x ; g(x)=2x+1", "Ny funktion", "f(x)=x+1")
    Var = InputBox(Sprog.A(122), Sprog.A(123), "f(x)=x+1")
    Var = Replace(Var, ":=", "=")
'    var = Replace(var, "=", VBA.ChrW(&H2261))
    
    If Var <> "" Then
        Var = Sprog.A(126) & ": " & Var
        Selection.InsertAfter (Var)
        Selection.OMaths.Add Range:=Selection.Range
        Selection.OMaths(1).BuildUp
        Selection.MoveRight unit:=wdCharacter, Count:=2
    End If
GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Sub DefinerLigning()
    Dim Var As String
On Error GoTo Fejl
    Var = InputBox(Sprog.A(115), Sprog.A(124), Sprog.A(125) & ":     Area:A=1/2*h*b")
'    var = Replace(var, "=", VBA.ChrW(&H2261))
    
    If Var <> "" Then
        Selection.InsertAfter (Var)
        Selection.OMaths.Add Range:=Selection.Range
        Selection.OMaths(1).BuildUp
        Selection.MoveRight unit:=wdCharacter, Count:=2
    End If
GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub MaximaSettings()
On Error GoTo Fejl
    If UFMSettings Is Nothing Then Set UFMSettings = New UserFormMaximaSettings
    UFMSettings.Show
    GoTo slut
Fejl:
    Set UFMSettings = New UserFormMaximaSettings
    UFMSettings.Show
slut:
End Sub
