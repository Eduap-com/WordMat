Attribute VB_Name = "MenuFunktioner"
Option Explicit

Sub OmMathMenu()
    Dim V As String
    V = AppVersion
    If PatchVersion <> "" Then
        V = V & PatchVersion
    End If
    MsgBox Sprog.A(20), vbOKOnly, AppNavn & " version " & V
End Sub

Sub indsaetformel()
    On Error GoTo Fejl
'    MsgBox CommandBars.ActionControl.Caption
#If Mac Then
#Else
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
#End If

    Application.ScreenUpdating = False
    If CommandBars.ActionControl.DescriptionText <> "" Then
    Selection.InsertAfter (CommandBars.ActionControl.DescriptionText)
    Selection.Collapse (wdCollapseEnd)
    Selection.TypeParagraph
    End If
    Selection.InsertAfter (CommandBars.ActionControl.Tag)
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).BuildUp
    Selection.MoveRight Unit:=wdCharacter, Count:=2
#If Mac Then
#Else
        Oundo.EndCustomRecord
#End If

GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:
End Sub
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
GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:
End Function
Sub DefinerVar()
    Dim Var As String
    On Error GoTo Fejl
'    var = InputBox("Indtast definitionen på den nye variabel" & vbCrLf & vbCrLf & "Definitionen kan benyttes i resten af dokumentet, men ikke før. Hvis der indsættes en clearvars: kommando længere nede i dokumentet kan den ikke benyttes derefter." & vbCrLf & vbCrLf & "Definitionen kan indtastes på 4 forskellige måder" & vbCrLf & vbCrLf & "definer: variabel=værdi" & vbCrLf & "variabel:værdi" & vbCrLf & "variabel:=værdi" & vbCrLf & "variabel" & VBA.ChrW(&H2261) & "værdi  (Definitions ligmed)" & vbCrLf & "Der kan defineres flere variable i en ligningsboks ved at adskille definitionerne med semikolon. f.eks. a:1 ; b:2", "Ny variabel", "a=1")
    Var = InputBox(Sprog.A(120), Sprog.A(121), "a=1")
    Var = Replace(Var, ":=", "=")
'    var = Replace(var, "=", VBA.ChrW(&H2261))
    If Var <> "" Then
        Var = Sprog.A(126) & ": " & Var
        Selection.InsertAfter (Var)
        Selection.OMaths.Add Range:=Selection.Range
        Selection.OMaths(1).BuildUp
        Selection.MoveRight Unit:=wdCharacter, Count:=2
    End If
    
GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:
End Sub
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
        Selection.MoveRight Unit:=wdCharacter, Count:=2
    End If
GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:
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
        Selection.MoveRight Unit:=wdCharacter, Count:=2
    End If
GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:
End Sub
Sub ErstatPunktum()
'
' ErstatPunktum Makro
'
'
On Error GoTo Fejl
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Font.Name = "Cambria Math"
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ","
        .Replacement.Text = ";"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Font.Name = "Cambria Math"
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "."
        .Replacement.Text = ","
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:
End Sub
Sub ErstatKomma()
'
' ErstatPunktum Makro
'
'
On Error GoTo Fejl
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Font.Name = "Cambria Math"
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ","
        .Replacement.Text = "."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Font.Name = "Cambria Math"
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ";"
        .Replacement.Text = ","
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:
End Sub

Sub Gange()
    On Error Resume Next
    Selection.InsertSymbol CharacterNumber:=AscW(MaximaGangeTegn), Unicode:=True
End Sub
Sub SimpelUdregning()
' laver simpel udregning med 4 regningsarter og ^
    
    On Error GoTo Slut
    Dim r As Range
    Dim sindex As Integer
    Dim resultat As String
    
    Application.ScreenUpdating = False
    If Selection.OMaths.Count > 0 Then
        Selection.OMaths(1).Range.Select
        If Selection.Range.Font.Bold = True Then
            Selection.Range.Font.Bold = False
        End If
        Selection.OMaths(1).Range.Select
        Selection.OMaths(1).Linearize
    End If
    If Len(Selection.Text) < 2 Then
'        MsgBox "Marker det udtryk der skal beregnes. Udtrykket må kun indeholde tal, de fire regningsarter og ^ ."
        Set r = Selection.Range

        sindex = 0
        Call r.MoveStart(wdCharacter, -1)
        Do
        sindex = sindex + 1
        Call r.MoveStart(wdCharacter, -1)
        Loop While sindex < 20 And (AscW(r.Characters(1)) > 39 And AscW(r.Characters(1)) < 58 Or AscW(r.Characters(1)) = 94 Or AscW(r.Characters(1)) = 183)
        If sindex < 20 Then Call r.MoveStart(wdCharacter, 1)
        Selection.start = r.start
'        Selection.End = r.End
    End If
    
    Call ActiveDocument.Range.Find.Execute(VBA.ChrW(8727), , , , , , , , , "*", wdReplaceAll) ' nødvendig til mathboxes
    Call Selection.Range.Find.Execute(VBA.ChrW(183), , , , , , , , , "*", wdReplaceAll)
    Call Selection.Range.Find.Execute(".", , , , , , , , , ",", wdReplaceAll)
    resultat = Selection.Range.Calculate()
    Selection.Range.InsertAfter ("=" & resultat)
    Call Selection.Range.Find.Execute(VBA.ChrW(42), , , , , , , , , VBA.ChrW(183), wdReplaceAll)
    Selection.MoveEnd Unit:=wdCharacter, Count:=Len(resultat) + 1
    Call Selection.Range.Find.Execute(",", , , , , , , , , ".", wdReplaceAll)
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths.BuildUp
    Selection.Collapse (wdCollapseEnd)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
'    Selection.TypeText (" ")

Slut:
End Sub

Sub ReplaceStarMult()
' fjerner stjerner og indsætter alm. gangetegn
Application.ScreenUpdating = False
On Error GoTo Fejl

'    Call ActiveDocument.Range.Find.Execute(chr(42), , , , , , , , , VBA.ChrW(183), wdReplaceAll)
    Call ActiveDocument.Range.Find.Execute(VBA.ChrW(8727), , , , , , , , , VBA.ChrW(183), wdReplaceAll) ' nødvendig til mathboxes
    Call ActiveDocument.Range.Find.Execute("*", , , , , , , , , VBA.ChrW(183), wdReplaceAll)

'    MsgBox "Alle * er nu lavet om til " & VBA.ChrW(183)
GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:
End Sub
Sub ReplaceStarMultBack()
' fjerner alm gangetegn og indsætter *
Application.ScreenUpdating = False
On Error GoTo Fejl
    Call Selection.Range.Find.Execute(VBA.ChrW(183), , , , , , , , , "*", wdReplaceAll)
    Call Selection.Range.Find.Execute(VBA.ChrW(8901), , , , , , , , , "*", wdReplaceAll) '\cdot
    Call Selection.Range.Find.Execute(VBA.ChrW(8729), , , , , , , , , "*", wdReplaceAll) ' \cdot
    Call Selection.Range.Find.Execute(VBA.ChrW(8226), , , , , , , , , "*", wdReplaceAll) ' tyk prik
    
'    MsgBox "Alle " & VBA.ChrW(183) & " er nu lavet om til *"
GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:
End Sub
Sub MaximaSettings()
On Error GoTo Fejl
    If UFMSettings Is Nothing Then Set UFMSettings = New UserFormMaximaSettings
    UFMSettings.Show
    GoTo Slut
Fejl:
    Set UFMSettings = New UserFormMaximaSettings
    UFMSettings.Show
Slut:
End Sub
