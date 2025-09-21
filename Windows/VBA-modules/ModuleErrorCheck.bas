Attribute VB_Name = "ModuleErrorCheck"
Option Explicit

Type ErrorDefinition
    Title As String ' if this is empty, then there is no error
    Description As String
    MaximaOutput As String
    DefFejl As Boolean
    LocationError As String ' text from Maxima that should be written with consolas so that the error can be pointed out
    Stop As Boolean ' means that you must stop after the error has been shown
End Type

Function CheckForError() As Boolean
' Checks if latest output has error, and if thats the case shows the error in a Userform
'    Dim UFerror As UserFormError
    Dim ED As ErrorDefinition

    On Error Resume Next
    CheckForError = False

    ED = GetErrorDefinition(omax.MaximaOutput, omax.KommentarOutput)
    If omax.DefFejl = True Then
        ED.DefFejl = True
        Application.ScreenUpdating = True
        ActiveWindow.ScrollIntoView Selection.Range
    End If
    If ED.Stop Then CheckForError = True
    
    If ED.Title <> vbNullString Then ' Show the error in userform
        UserFormError.Label_maximaoutput.Caption = ED.MaximaOutput    ' eller maxproc.lastmaximaoutput
        If ED.DefFejl Then
            ED.Description = TT.A(758) & vbCrLf & vbCrLf & ED.Description
'            UserFormError.TextBox_definitioner.text = FormatDefinitions(omax.DefString) 'DefinitionsNice
        End If
        UserFormError.SetErrorDefinition ED
        UserFormError.Show
        RestartWordMat
    End If
End Function

Function GetErrorDefinition(MaximaOutput As String, KommentarOutput As String) As ErrorDefinition
' Classifies and interprets the error in an errordefinition.
' Checktext must be output from Maxima
    Dim pos As Integer, CheckText As String, CheckText2 As String, s As String, s2 As String
    GetErrorDefinition.Stop = True
    CheckText = MaximaOutput & KommentarOutput
    CheckText2 = Replace(CheckText, " ", vbNullString) ' on mac there are spaces, but not on windows
    If InStr(CheckText2, "syntaxerror") > 0 Then
        GetErrorDefinition.Title = TT.SyntaxError
        GetErrorDefinition.Description = TT.A(752)
    ElseIf InStr(CheckText2, "incorrectsyntax:Missing") > 0 Then
        GetErrorDefinition.Title = TT.SyntaxError
        pos = InStr(CheckText, "incorrectsyntax:Missing")
        GetErrorDefinition.Description = TT.A(753) & " " & Mid$(CheckText, pos + 26, 1)
    ElseIf InStr(CheckText2, "incorrectsyntax:Toomany") > 0 Then
        GetErrorDefinition.Title = TT.SyntaxError
        pos = InStr(CheckText, "incorrect syntax: Too many")
        GetErrorDefinition.Description = TT.SyntaxError & ". " & vbCrLf & TT.A(754) & " " & Mid$(CheckText, pos + 29, 1)
    ElseIf InStr(CheckText2, "incorrectsyntax:Found") > 0 Then
        GetErrorDefinition.Title = TT.SyntaxError
        GetErrorDefinition.Description = GetErrorText("incorrect syntax: ", CheckText)
    ElseIf InStr(CheckText2, "factorial:factorialofnegativeinteger") > 0 Then
        GetErrorDefinition.Title = TT.A(750)
        s = ExtractText(CheckText, "factorial: ", "#0", "-- an error")
        s2 = ExtractText(CheckText, "integer ", " not")
        If TT.LangNo = 1 Then
            s = TT.A(699)
            If InStr(CheckText, "K(") > 0 Then
                s = s & VbCrLfMac & TT.A(902)
            End If
        End If
        GetErrorDefinition.Description = s
    ElseIf InStr(CheckText2, "isnotaprefixoperator") > 0 Then
        GetErrorDefinition.Title = TT.SyntaxError
        GetErrorDefinition.Description = GetErrorText("is not a prefix operator", CheckText)
    ElseIf InStr(CheckText2, "isnotaninfixoperator") > 0 Then
        GetErrorDefinition.Title = TT.SyntaxError
        GetErrorDefinition.Description = GetErrorText("is not an infix operator", CheckText)
    ElseIf InStr(CheckText2, "Prematureterminationofinputat") > 0 Then
        GetErrorDefinition.Title = TT.SyntaxError
        GetErrorDefinition.LocationError = GetErrorText("Premature termination of input at", CheckText, 4)
    ElseIf InStr(CheckText2, "Toofewargumentssuppliedto") > 0 Then
        GetErrorDefinition.Title = TT.SyntaxError
        GetErrorDefinition.LocationError = TT.A(898) & ExtractText(CheckText2, "Toofewargumentssuppliedto", ";") & vbCrLf & TT.A(899)
    ElseIf InStr(CheckText2, "incorrectsyntax:") > 0 Then
        GetErrorDefinition.Title = TT.SyntaxError
        GetErrorDefinition.Description = TT.SyntaxError & "."
    ElseIf InStr(CheckText2, "toomanycolonsin") > 0 Then
        GetErrorDefinition.Title = TT.A(755)
        GetErrorDefinition.Description = "Der er for mange kolonner"
    ElseIf InStr(CheckText2, "lisperror") And InStr(CheckText, "[") > 0 Then
        GetErrorDefinition.Title = "Lisp error"
        GetErrorDefinition.Description = TT.A(755)
    ElseIf InStr(CheckText2, "encounteredaLisperror") > 0 Then
        GetErrorDefinition.Title = "Lisp error"
        GetErrorDefinition.Description = TT.A(755)
    ElseIf InStr(CheckText2, "Todebugthistry:debugmode(true)") > 0 Then ' The leading anerror is not included, as on Mac there is also a period between
        If InStr(CheckText2, "expt:undefined:0toanegativeexponent") > 0 Then
            GetErrorDefinition.Title = "Division by zero"
            GetErrorDefinition.Description = TT.A(756)
        Else
            GetErrorDefinition.Title = "Lisp error"
            GetErrorDefinition.Description = TT.A(755)
        End If
    ElseIf CheckText2 = "?merror(""Anumberwasfoundwhereavariablewasexpected-`solve'"")" Then
        GetErrorDefinition.Title = "Variable error"
        GetErrorDefinition.Description = TT.A(133)
    ElseIf (omax.DefFejl = True) Then
        GetErrorDefinition.Title = "Definition error"
        GetErrorDefinition.Description = TT.A(757) & vbCrLf & VisDef
'#If Mac Then ' ved ikke hvorfor dette er på mac. Det giver specifikt et problem ved at løse 'f(x)=2x+5' uden definition af f(x)
'    ElseIf (MaximaOutput = vbNullString) Then
'        GetErrorDefinition.Title = "Timeout"
'        GetErrorDefinition.Description = TT.A(884)
'        GetErrorDefinition.Stop = False
'#End If
        Else
        GetErrorDefinition.Stop = False
    End If
    
    GetErrorDefinition.MaximaOutput = KommentarOutput & vbCrLf & vbCrLf & MaximaOutput
    
End Function

Function GetErrorText(text As String, MaximaOutput As String, Optional RemoveChrS As Integer = 0) As String
' used by GetErrorDefinition()
    Dim pos As Integer, pos2 As Integer, pos4 As Integer
    Dim t As String
    Dim l As Integer
    On Error Resume Next
    l = Len(text) + RemoveChrS
    pos = InStr(MaximaOutput, "incorrect syntax")
    pos2 = InStr(pos, MaximaOutput, text)
    pos4 = InStrRev(pos2 + l, MaximaOutput, "^")
    If pos4 < 1 Then
        pos4 = Len(MaximaOutput)
    End If
    t = Mid$(MaximaOutput, pos2 + l, pos4 - pos2 - l + 1)
'    t = Replace(t, "^", vbCrLf & "    ^", 1, 1)
    t = TrimL(t, vbLf)
    t = TrimL(t, vbCr)
    t = TrimR(t, vbLf)
    t = TrimR(t, vbCr)
    t = TrimR(t, vbCrLf)
    t = Trim$(t)
    pos = InStrRev(t, ";")
    If pos > 0 Then
        t = Left$(t, pos - 1) & Right$(t, Len(t) - pos)
    End If
    t = Replace(t, ":=", "= ")
    
    GetErrorText = TT.A(759) & ":" & vbCrLf & t
'    GetErrorText = TT.SyntaxError & vbCrLf & TT.IllegalSymbol & ":" & vbCrLf & t

End Function
Function ExtractText(ByVal text As String, startText As String, endText As String, Optional endText2 As String) As String
    ' returns substring from Text between startText and EndText
    Dim p As Long

    If startText <> vbNullString Then
        p = InStr(text, startText)
        If p > 0 Then text = Right$(text, Len(text) - p - Len(startText) + 1)
    End If
    If endText <> vbNullString Then
        p = InStr(text, endText)
        If p > 0 Then text = Left$(text, p - 1)
    End If
    If endText2 <> vbNullString Then
        p = InStr(text, endText2)
        If p > 0 Then text = Left$(text, p - 1)
    End If
    ExtractText = text
End Function
Function DefinitionsNice() As String
' Used by CheckForError
    Dim DefS As String
    DefS = omax.DefString
    If Len(DefS) > 3 Then
        '    defs = mid$(defs, 2, Len(defs) - 3)
        DefS = omax.ConvertToAscii(DefS)
        DefS = Replace(DefS, "$", vbCrLf)
        DefS = Replace(DefS, ":=", " = ")
        DefS = Replace(DefS, ":", " = ")
        If DecSeparator = "," Then
            DefS = Replace(DefS, ",", ";")
            DefS = Replace(DefS, ".", ",")
        End If
    End If

    DefinitionsNice = DefS
End Function

