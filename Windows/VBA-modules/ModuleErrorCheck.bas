Attribute VB_Name = "ModuleErrorCheck"
Option Explicit

Type ErrorDefinition
    Title As String ' hvis denne er tom, så er der ingen fejl
    Description As String
    MaximaOutput As String
    DefFejl As Boolean
    Stop As Boolean ' means that you must stop after the error has been shown
End Type

Function CheckForError() As Boolean
' Checks if latest output has error, and if thats the case shows the error in a Userform
'    Dim UFerror As UserFormError
    Dim ndeferror As Integer
    Dim ED As ErrorDefinition

    If omax.DefString <> "" Then
        ndeferror = 5
    Else
        ndeferror = 3
    End If
    On Error Resume Next
    CheckForError = False

    ED = GetErrorDefinition(omax.MaximaOutput, omax.KommentarOutput)
    
    If ED.Stop Then CheckForError = True
    
    If ED.Title <> vbNullString Then ' Show the error in userform
        
'        Set UFerror = New UserFormError
        UserFormError.Label_maximaoutput.Caption = ED.MaximaOutput    ' eller maxproc.lastmaximaoutput
        If ED.DefFejl Then
'            If UBound(Minput) < ndeferror Then
                ED.Description = ED.Description & vbCrLf & vbCrLf & Sprog.DefError2
'            Else
'                ED.Description = ED.Description & vbCrLf & vbCrLf & Sprog.DefError3
'            End If
            UserFormError.TextBox_definitioner.text = FormatDefinitions(omax.DefString) 'DefinitionsNice
        Else
            UserFormError.TextBox_definitioner.visible = False
            UserFormError.Label_definitioner.visible = False
        End If
        UserFormError.SetErrorDefinition ED
        UserFormError.Show
        RestartWordMat
    End If
End Function

Function GetErrorDefinition(MaximaOutput As String, KommentarOutput As String) As ErrorDefinition
' Klassificerer og fortolker fejlen i en errordefinition.
' Checktext skal være output fra Maxima
    Dim Pos As Integer, CheckText As String, CheckText2 As String
    GetErrorDefinition.Stop = True
    CheckText = MaximaOutput & KommentarOutput
    CheckText2 = Replace(CheckText, " ", vbNullString) ' på mac er der mellemrum, men ikke på windows
    If InStr(CheckText2, "syntaxerror") > 0 Then
        GetErrorDefinition.Title = "Syntax error"
        GetErrorDefinition.Description = Sprog.SyntaxErrorLong
    ElseIf InStr(CheckText2, "incorrectsyntax:Missing") > 0 Then
        GetErrorDefinition.Title = "Syntax error"
        Pos = InStr(CheckText, "incorrectsyntax:Missing")
        GetErrorDefinition.Description = Sprog.Missing & " " & Mid(CheckText, Pos + 26, 1)
        GetErrorDefinition.DefFejl = True
    ElseIf InStr(CheckText2, "incorrectsyntax:Toomany") > 0 Then
        GetErrorDefinition.Title = "Syntax error"
        Pos = InStr(CheckText, "incorrect syntax: Too many")
        GetErrorDefinition.Description = Sprog.SyntaxError & ". " & vbCrLf & Sprog.TooMany & " " & Mid(CheckText, Pos + 29, 1)
        GetErrorDefinition.DefFejl = True
    ElseIf InStr(CheckText2, "isnotaprefixoperator") > 0 Then
        GetErrorDefinition.Title = "Syntax error"
        GetErrorDefinition.Description = GetErrorText("is not a prefix operator", CheckText)
        GetErrorDefinition.DefFejl = True
    ElseIf InStr(CheckText2, "isnotaninfixoperator") > 0 Then
        GetErrorDefinition.Title = "Syntax error"
        GetErrorDefinition.Description = GetErrorText("is not an infix operator", CheckText)
        GetErrorDefinition.DefFejl = True
    ElseIf InStr(CheckText2, "Prematureterminationofinputat") > 0 Then
        GetErrorDefinition.Title = "Syntax error"
        GetErrorDefinition.Description = GetErrorText("Premature termination of input at", CheckText)
        GetErrorDefinition.DefFejl = True
    ElseIf InStr(CheckText2, "incorrectsyntax:") > 0 Then
        GetErrorDefinition.Title = "Syntax error"
        GetErrorDefinition.Description = Sprog.SyntaxError & "."
        GetErrorDefinition.DefFejl = True
    ElseIf InStr(CheckText2, "lisperror") And InStr(CheckText, "[") > 0 Then
        GetErrorDefinition.Title = "Lisp error"
        GetErrorDefinition.Description = Sprog.LispError
    ElseIf InStr(CheckText2, "encounteredaLisperror") > 0 Then
        GetErrorDefinition.Title = "Lisp error"
        GetErrorDefinition.Description = Sprog.LispError
'    ElseIf InStr(KommentarOutput, "Division by 0") > 0 Then ' maybe not relevant in SBCL
'        fejltekst = Sprog.DivisionByZero
'        CheckForError = True
'    ElseIf InStr(CheckText, "expt:undefined:0toanegativeexponent") > 0 Then ' kommer ved mange alm beregninger uden fejl, flyttet ned med en anden
'        GetErrorDefinition.Title = "Division by zero"
'        GetErrorDefinition.Description = Sprog.DivisionByZero
    ElseIf InStr(CheckText2, "Todebugthistry:debugmode(true)") > 0 Then ' anerror foran er ikke taget med, da på Mac er der ogspå punktum mellem
        If InStr(CheckText2, "expt:undefined:0toanegativeexponent") > 0 Then
            GetErrorDefinition.Title = "Division by zero"
            GetErrorDefinition.Description = Sprog.DivisionByZero
        Else
            GetErrorDefinition.Title = "Lisp error"
            GetErrorDefinition.Description = Sprog.LispError
        End If
    ElseIf CheckText2 = "?merror(""Anumberwasfoundwhereavariablewasexpected-`solve'"")" Then
        GetErrorDefinition.Title = "Variable error"
        GetErrorDefinition.Description = Sprog.A(133) '"Du har bedt om at løse ligningen for en variabel der allerede er defineret. Indsæt en 'slet def:' kommando før ligningen"
    ElseIf (omax.DefFejl = True) Then
        GetErrorDefinition.Title = "Definition error"
        GetErrorDefinition.Description = Sprog.DefError & vbCrLf & VisDef
#If Mac Then
    ElseIf (MaximaOutput = vbNullString) Then
        GetErrorDefinition.Title = "Timeout"
        If Sprog.SprogNr = 1 Then
            GetErrorDefinition.Description = "Beregningen blev afbrudt, da den tog meget lang tid. Det kan nogle gange hjælpe at prøve beregningen med indstillingen 'Numerisk' istedet for Eksakt eller Auto"
        Else
            GetErrorDefinition.Description = "Calculation timed out. Try again with numerical setting."
        End If
        GetErrorDefinition.Stop = False
#End If
        Else
        GetErrorDefinition.Stop = False
    End If
    
    GetErrorDefinition.MaximaOutput = KommentarOutput & vbCrLf & vbCrLf & MaximaOutput
    
End Function

Function GetErrorText(text As String, MaximaOutput As String) As String
' used by GetErrorDefinition()
    Dim Pos As Integer, pos2 As Integer, pos4 As Integer
    Dim t As String
    Dim L As Integer
    On Error Resume Next
    L = Len(text)
    Pos = InStr(MaximaOutput, "incorrectsyntax")
    pos2 = InStr(Pos, MaximaOutput, text)
    pos4 = InStr(pos2 + L, MaximaOutput, "^")
    If pos4 < 1 Then
        pos4 = Len(MaximaOutput)
    End If
    t = Mid(MaximaOutput, pos2 + L, pos4 - pos2 - L + 1)
    t = Replace(t, "^", vbCrLf & "    ^", 1, 1)
    GetErrorText = Sprog.SyntaxError & vbCrLf & Sprog.IllegalSymbol & ":" & vbCrLf & t

End Function

Function DefinitionsNice() As String
' Used by CheckForError
    Dim DefS As String
    DefS = omax.DefString
    If Len(DefS) > 3 Then
        '    defs = Mid(defs, 2, Len(defs) - 3)
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

