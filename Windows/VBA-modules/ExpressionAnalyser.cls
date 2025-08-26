VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExpressionAnalyser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public text As String
Public StartBracket As String
Public EndBracket As String
Private SaveStartBracket As String
Private SaveEndBracket As String
Public pos As Long

Public Function GetNextchr() As String
' returns the next letter in the string and removes it
GetNextchr = Left$(text, 1)
text = right$(text, Len(text) - 1)

End Function

Public Property Get NextChr() As Variant
    NextChr = Left$(text, 1)
End Property
Public Property Get Length() As Long
    Length = Len(text)
End Property

Public Function FindNextBracket(StartPos As Long) As Long
Dim pos1 As Long
Dim pos2 As Long

pos1 = InStr(StartPos, text, StartBracket)
pos2 = InStr(StartPos, text, EndBracket)

If pos1 = 0 And pos2 = 0 Then
    FindNextBracket = 0
    Exit Function
End If
If pos1 < 1 Then pos1 = Len(text)
If pos2 < 1 Then pos2 = Len(text)

If pos1 < pos2 Then
    FindNextBracket = pos1
Else
    FindNextBracket = pos2
End If

End Function
Public Function FindPrevBracket(StartPos As Long) As Long
Dim pos1 As Long
Dim c As String
pos1 = StartPos + 1
Do
    pos1 = pos1 - 1
    c = ChrByIndex(pos1)
Loop Until c = StartBracket Or c = EndBracket Or pos1 = 1

    FindPrevBracket = pos1

'If c = StartBracket Or c = EndBracket Then
'    FindPrevBracket = pos1
'Else
'    FindPrevBracket = 0
'End If

End Function
Public Function FindEndOfBracket() As Long
' finds next parenthesis and returns position after closing parenthesis
    GetNextBracketContent
    FindEndOfBracket = pos
End Function
Public Function GetNextBracketContent(Optional ByVal StartPos As Long) As String
' returns the contents of the next parenthesis
' empty if parenthesis do not match
' pos is changed to after the parenthesis

Dim pos1 As Long
Dim pos2 As Long
Dim bracketlevel As Integer
If StartPos < 1 Then
    StartPos = pos
    If StartPos < 1 Then
        StartPos = 1
        pos = 1
    End If
End If
pos1 = StartPos - 1
Do
    pos1 = FindNextBracket(pos1 + 1)
Loop Until ChrByIndex(pos1) = StartBracket Or pos1 = 0

If pos1 = 0 Then
    GetNextBracketContent = ""
    Exit Function
End If

pos2 = pos1
bracketlevel = 1
Do
    pos2 = FindNextBracket(pos2 + 1)
    If ChrByIndex(pos2) = StartBracket Then
        bracketlevel = bracketlevel + 1
    Else
        bracketlevel = bracketlevel - 1
    End If
Loop Until bracketlevel = 0 Or pos2 = 0

If pos2 = 0 Then
    GetNextBracketContent = ""
    Exit Function
End If
    
GetNextBracketContent = Mid$(text, pos1 + 1, pos2 - pos1 - 1)
pos = pos2 + 1
End Function
Public Function GetPrevBracketContent(Optional ByVal StartPos As Long) As String
' returns the contents of the parenthesis before pos
' empty if parenthesis do not match
' pos is changed to the first parenthesis character

Dim pos1 As Long
Dim pos2 As Long
Dim bracketlevel As Integer
If StartPos < 1 Then
    StartPos = pos
    If StartPos < 1 Then
        StartPos = 1
        pos = 1
    End If
End If
pos1 = StartPos + 1
Do
    pos1 = pos1 - 1
Loop Until ChrByIndex(pos1) = EndBracket Or pos1 = 0

If pos1 = 0 Then
    GetPrevBracketContent = ""
    Exit Function
End If

pos2 = pos1
bracketlevel = 1
Do
    pos2 = FindPrevBracket(pos2 - 1)
    If ChrByIndex(pos2) = EndBracket Then
        bracketlevel = bracketlevel + 1
    Else
        bracketlevel = bracketlevel - 1
    End If
Loop Until bracketlevel = 0 Or pos2 = 0

If pos2 = 0 Then
    GetPrevBracketContent = ""
    Exit Function
End If
    
GetPrevBracketContent = Mid$(text, pos2 + 1, pos1 - pos2 - 1)
pos = pos2 + 1
End Function

Public Property Get ChrByIndex(ByVal pos As Long) As String
On Error GoTo fejl
    If pos <= Len(text) And pos > 0 Then
        ChrByIndex = Mid$(text, pos, 1)
    Else
        ChrByIndex = VBA.ChrW$(0)
    End If
GoTo slut
fejl:
        ChrByIndex = VBA.ChrW$(0)
slut:
End Property

Public Function FindNextBreak(Optional stopforbrackets As Boolean = False) As Long
' finds next time where there should be a parenthesis.
' the parenthesis must be after the returned position
' i.e. finds spaces, +-*/^=<>):;,@? parenthesis now terminates before they didn't but it caused problems with e.g. (3+sqrt3)^2
' startparenthesis(c=40) does not terminate as it causes problems with functions e.g. 2^log(3)
' possibly problems with d/dx notation
' comma must be included as there may be maxima functions in the expression solve(2=9^2x,x) e.g.
' stopforbrackets is deactivated as it is assumed it is not necessary, but maybe dx notation for integrals?
' stop for brackets reactivated to work in converttowordsymbols f_a(1) -> f_a (1)
    Dim c As Integer
    
    Do
    c = AscW(ChrByIndex(pos))
    If Not stopforbrackets Then
    If c = 40 Then ' hvis parantesstart hoppes til parantesslut
        GetNextBracketContent
        pos = pos - 1
'    ElseIf c = 41 And pos < Len(text) Then
'        pos = pos + 1
'        c = AscW(ChrByIndex(pos))
    End If
    If c = 91 Then ' hvis [ parantesstart hoppes til parantesslut ]
        SetSquareBrackets
        GetNextBracketContent
        SetNormalBrackets
        pos = pos - 1
    End If
    Else
        If c = 40 Then
            pos = pos + 1
            Exit Do
        End If
    End If
    pos = pos + 1
    Loop Until pos > Len(text) Or c = 32 Or (c >= 41 And c <= 45) Or c = 47 Or (c >= 58 And c <= 64) Or c = 93 Or c = 94 Or c = 9618
'    Loop While c >= 48 Or c >= 57 Or c = 46 Or c = 44 Or c >= 65 Or c >= 90 Or c >= 97 Or c >= 122
    If pos > Len(text) Then
        pos = pos - 1
    Else
        pos = pos - 2
    End If
    FindNextBreak = pos

End Function
Public Function FindHierakiBreak(level As Integer, Optional komma As Boolean = True) As Long
' finds the next operator depending on the specified hierarchy level.
'Hierarki level:
'=<>)] 0 sum and product symbols
':;,@? space    1   only comma if komma=true
'+-  2
'*/  3
'^_  4
'({[|   5
' the returned value is the position of the last character before the operator
' level 0 only stops for operators on level 0
' level 1 only stops for operators on level  0-1
' level 2 only stops for operators on level  0-2
' ...
' findnextbreak corresponds to level 4

    
    Dim c As Integer, i As Integer
    
    Do
    i = i + 1 ' infinite loop security
    c = AscW(ChrByIndex(pos))
    If (c >= 60 And c <= 62) Or c = 41 Or c = 93 Or c = 125 Or c = 8776 Or c = 8719 Or c = 8721 Or c = 8594 Or c = 8592 Or c = 8596 Or c = 8651 Or c = 8652 Then Exit Do ' 41=) , 93=] , 123=} (added 1.24 geogebracas), 60=<, 61== , 62=> , 8776=cirkaligmed, 8719=produkttegn , 8721=sumtegn, 8594=-> , 8592=<-, 8596=<->, 8651,8652=kemidblpil
    
    If level > 0 Then
        If c = 32 Or c = 58 Or c = 59 Or c = 63 Or c = 64 Or c = 38 Then Exit Do    ' 32=space, 58=: , 59=; , 44=, 63=? , 64=@ , 38=&
        If komma And c = 44 Then Exit Do
    End If
    If level > 1 Then
        If c = 43 Or c = 45 Then ' 43=+, 45=-
            If ChrByIndex(pos - 1) <> "E" Then Exit Do ' E- og E+ should not stop
            If pos = 2 Then Exit Do
            If Not (IsNumber(pos - 2)) Then Exit Do
        End If
    End If
    If level > 2 Then
        If c = 42 Or c = 215 Or c = 47 Or c = 183 Then Exit Do ' 42=*, 47=/ , 215=\times, 183=\cdot
    End If
    If level > 3 Then
        If c = 94 Or c = 95 Then Exit Do ' 94=^, 95=_
    End If
    If level > 4 Then
        If c = 40 Or c = 91 Or c = 123 Or c = 124 Or c = 125 Or c = 12310 Then Exit Do ' 40=(, 91=[ , 123={ , 124=|, 125=}
    End If
    
    If c = 40 Then ' if the start of the parentheses is jumped to the end of the parentheses
        SaveBrackets
        SetNormalBrackets
        GetNextBracketContent
        RestoreBrackets
        pos = pos - 1
    End If
    If c = 91 Then ' if [ start of parentheses jump to end of parentheses ]
        SaveBrackets
        SetSquareBrackets
        GetNextBracketContent
        RestoreBrackets
        pos = pos - 1
    End If
    If c = 123 Then ' if {
        SaveBrackets
        SetTuborgBrackets
        GetNextBracketContent
        RestoreBrackets
        pos = pos - 1
    End If
    If c = 12310 Then ' if {
        SaveBrackets
        SetInvisibleBrackets
        GetNextBracketContent
        RestoreBrackets
        pos = pos - 1
    End If
    
    pos = pos + 1
    Loop Until pos > Len(text) Or i > 1000
    
    pos = pos - 1
    FindHierakiBreak = pos

End Function
Public Function FindPrevHierakiBreak(level As Integer, Optional komma As Boolean = True, Optional StopForStartBracket As Boolean = True)
    Dim c As Integer

    Do
    c = AscW(ChrByIndex(pos))
    If c = 41 Or c = 93 Or c = 125 Then ' if parentheses end jump to parentheses start )]}
        SaveBrackets
        SetBrackets (c)
        GetPrevBracketContent
        pos = pos - 1
        RestoreBrackets
    End If
    
    If (c >= 60 And c <= 62) Or c = 41 Or c = 93 Or c = 8776 Or c = 8289 Then Exit Do ' 41=) , 93=] , 60=<, 61== , 62=> , 8776=cirka equal, 8289=function symbol
    
    If level > 0 Then
        If c = 32 Or c = 58 Or c = 59 Or c = 63 Or c = 64 Or c = 38 Then Exit Do ' 32=space, 58=: , 59=; , 44=, 63=? , 64=@, 38=&
        If komma And c = 44 Then Exit Do
    End If
    If level > 1 Then
        If c = 43 Or c = 45 Then Exit Do ' 43=+, 45=-
    End If
    If level > 2 Then
        If c = 42 Or c = 215 Or c = 47 Then Exit Do ' 42=*, 47=/ , 215=\times
    End If
    If level > 3 Then
        If c = 94 Or c = 95 Then Exit Do ' 94=^, 95=_
    End If
    If level > 4 Or StopForStartBracket Then
        If c = 40 Or c = 91 Or c = 123 Or c = 124 Then Exit Do ' 40=(, 91=[ , 123={ , 124=|
    End If
    
    pos = pos - 1
    Loop Until pos < 1
    pos = pos + 1
    If pos = 0 Then
        pos = 1
    End If
    FindPrevHierakiBreak = pos

End Function
Public Function FindPrevBreak() As Long
' finds last time where there should be parentheses
' i.e. finds spaces, +-*/=<>)]}:;,?@ parentheses now end before they didn't but it caused problems with e.g. (3+sqrt3)^2
' 8289 is symbol after function
' ^ removed as it otherwise didn't work for e.g. sin(x)^2/pi

    Dim c As Integer
    
    Do
    c = AscW(ChrByIndex(pos))
    If c = 41 Or c = 93 Or c = 125 Or c = 12311 Then ' if parentheses end jump to parentheses start )]}
        SaveBrackets
        SetBrackets (c)
        GetPrevBracketContent
        pos = pos - 1
        RestoreBrackets
    End If
    pos = pos - 1
    Loop Until pos < 0 Or c = 32 Or c = 38 Or (c >= 42 And c <= 45) Or c = 40 Or c = 47 Or (c >= 58 And c <= 64) Or c = 9618 Or c = 91 Or c = 123 Or c = 12310
    If pos < 0 Then ' v. 1.29 changed to < instead of <= due to problem where [1/.. became ([1)/...
        pos = 1
    Else
        pos = pos + 2
    End If
    FindPrevBreak = pos

End Function

Public Sub InsertUnderstoodBracketPair()
    'inserts parenthesis just before current position and finds where the closing parenthesis should be
    Dim c As Integer
    text = Left$(text, pos - 1) & StartBracket & right$(text, Len(text) - pos + 1)
    pos = pos + 1
    c = AscW(ChrByIndex(pos)) ' +2 instead of +1 as it can give () e.g. at x^-2 ->x^()-2
    If Not (c = 40 Or c = 91 Or c = 123 Or c = 124 Or c = 12310) Then pos = pos + 1 ' 40=(, 91=[ , 123={ , 124=|

    If pos > Length Then
        pos = Length
        text = text & EndBracket
    Else
        FindHierakiBreak 3
        ' inserts parentheses after current position
        text = Left$(text, pos) & EndBracket & right$(text, Len(text) - pos)
    End If
End Sub
Public Sub InsertUnderstoodBracketPairBefore()
' if not already there
    Dim gpos As Long
    gpos = pos
    
' remove any spaces before / or ^
    If ChrByIndex(gpos) = " " Then
        RemoveChar (gpos)
        gpos = gpos - 1
        pos = pos - 1
    End If
    
    FindPrevBreak
    
    If (ChrByIndex(gpos) <> ")" Or ChrByIndex(pos) <> "(") Then
        text = Left$(text, pos - 1) & StartBracket & right$(text, Len(text) - pos + 1)
        text = Left$(text, gpos + 1) & EndBracket & right$(text, Len(text) - gpos - 1)
    End If
    
End Sub
Public Sub InsertBracketAfter(s As String)
' inserts an under-quoted parenthesis after symbol s throughout the text
' is used for ^ and / . If there is already a parenthesis after s, nothing is done
    Dim gempos As Long, t As String
    pos = 1
    Do
    pos = InStr(pos, text, s)
    If pos > 0 Then
        If pos > 3 Then t = Mid$(text, pos - 3, 3)
        If t = "tan" Or t = "sin" Or t = "cos" Then
             pos = pos + 1
        ElseIf Mid$(text, pos + 1, 1) <> "(" Then
'            If Not (IsNumber(pos + 1) And ChrByIndex(pos + 2) = " ") Then ' insert only if not prime number
                pos = pos + 1
                gempos = pos
                InsertUnderstoodBracketPair
                pos = gempos + 1
'            Else
'                text = left$(text, pos + 1) & "*" & right$(text, Len(text) - pos - 2) 'solves problem with 1/5 2 becoming 1/52 instead of 1/5*2. Maybe better to always use parentheses as the output is not visible
'                pos = pos + 3
'            End If
        Else
            gempos = pos
            GetNextBracketContent pos + 1 ' It must be handled that the notation can be e.g. 2x/(x+1)(x-1) in Word. Then divide by both parentheses when a multiplication sign is not specified. Added in version 1.09. The problem occurs a little sporadically as it is not translated consistently in this way to linear form. The same equation that is converted back and forth can be with or without extra parentheses
            If Mid$(text, pos, 1) = "(" Then
                pos = gempos + 1
                InsertUnderstoodBracketPair
            Else
                pos = gempos + 2
            End If
'                pos = pos + 2
        End If
    End If
    Loop While pos > 0

End Sub
Public Sub InsertBracketBefore(s As String)
' inserts an under-quoted parenthesis before the symbol s in the entire text
' is used for ^ and / . If there is already a parenthesis before s, nothing is done
    Dim gempos As Long
    gempos = 1
    Do
    pos = InStr(gempos, text, s)
    gempos = pos + 1
    If pos > 0 Then
'        If mid$(text, pos - 1, 1) <> ")" Then
                pos = pos - 1
                gempos = gempos + 2
                InsertUnderstoodBracketPairBefore
'                pos = gempos - 1
'        Else
'            pos = pos - 2
'        End If
    End If
    Loop While pos > 0

End Sub
Public Function FindChr(c As String, Optional ByVal StartPos As Long) As Long
' finds next occurrence of c after pos and sets position to this
' if it is not found pos is not changed and -1 is returned
    If StartPos < 1 Then
        StartPos = pos
    End If
    FindChr = InStr(StartPos, text, c)
    If FindChr > 0 Then pos = FindChr
End Function
Public Sub InsertBeforePos(s As String, Optional ByVal position As Long)
' inserts string s in the middle of the text string after pos
    If position < 1 Then
        position = pos
    End If
    If position = 1 Then
        text = s & text
    Else
        text = Left$(text, position - 1) & s & right$(text, Len(text) - position + 1)
    End If
End Sub
Public Sub InsertAfterPos(s As String, Optional ByVal position As Long = -1)
    If position < 0 Then
        position = pos
    End If
    If position = Len(text) Then
        text = text & s
    Else
        text = Left$(text, position) & s & right$(text, Len(text) - position)
    End If
End Sub
Public Function GetNextVar(Optional StartPos As Long) As String
' finds next variable or function
'First character must be letter then can be letter, number or _ or [
' name is returned and position is set to last position in variable
Dim posa As Long
Dim sb As String
Dim sb2 As String
    
    If StartPos > 0 Then
        pos = StartPos
    End If
    If pos < 1 Then pos = 1
    Dim c As Integer
    
    Do ' find first letter in variable
        If pos > Len(text) Then
            GetNextVar = ""
            Exit Function
        End If
        c = AscW(ChrByIndex(pos))
        pos = pos + 1
    Loop While Not (IsLetterC(c, False)) Or (c = 95) ' letter, but not underscore
    posa = pos - 1
    
    pos = pos - 1
    Do ' find hvor variabel slutter
        pos = pos + 1
        c = AscW(ChrByIndex(pos))
        If c = 91 Then
            sb = StartBracket
            sb2 = EndBracket
            SetSquareBrackets
            pos = FindEndOfBracket
            StartBracket = sb
            EndBracket = sb2
            Exit Do
        End If
    Loop While pos <= Len(text) And ((c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= 48 And c <= 57) Or c = 95)
    ' here the variable may contain numbers
            
    pos = pos - 1
    
    GetNextVar = Mid$(text, posa, pos - posa + 1)

End Function
Public Function GetPrevVar(Optional StartPos As Long) As String
' finds variable or function before position
' First character must be letter or number then can be letter, number or _
' the name is returned and position is set to the last position in the variable
' last character must not be number (i.e. first character in the variable)

    If StartPos > 0 Then
        pos = StartPos
    End If

    Dim c As Integer
    
    Do ' find letter
        If pos = 0 Then
            GetPrevVar = ""
            Exit Function
        End If
        c = AscW(ChrByIndex(pos))
        pos = pos - 1
    Loop Until (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)
    
    Do ' find where variable ends
        c = AscW(ChrByIndex(pos))
        pos = pos - 1
    Loop While pos > 0 And ((c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= 48 And c <= 57) Or c = 95)
    
    pos = pos + 1
    GetPrevVar = GetNextVar(pos)
            
    
End Function
Public Function GetNextLatexCommand(Optional StartPos As Long) As String
' finds next latex command
'First character must be letter then can be letter, number or _ or [
' name is returned and position is set to last position in variable
Dim posa As Long
Dim sb As String
Dim sb2 As String
    
    If StartPos > 0 Then
        pos = StartPos
    End If
    If pos < 1 Then pos = 1
    Dim c As Integer
    
    Do ' find letter
        If pos > Len(text) Then
            GetNextLatexCommand = ""
            Exit Function
        End If
        c = AscW(ChrByIndex(pos))
        pos = pos + 1
'    Loop Until (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)
    Loop While Not (IsLetterC(c, False))
'    c = AscW(ChrByIndex(pos))
    posa = pos - 1
    
    pos = pos - 1
    Do ' find where variable ends
        pos = pos + 1
        c = AscW(ChrByIndex(pos))
        If c = 91 Then
            sb = StartBracket
            sb2 = EndBracket
            SetSquareBrackets
            pos = FindEndOfBracket
            StartBracket = sb
            EndBracket = sb2
            Exit Do
        End If
'    Loop While pos <= Len(text) And ((c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= 48 And c <= 57) Or c = 95)
    Loop While pos <= Len(text) And ((c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= 48 And c <= 57) Or c = 95)
    ' here the variable may contain numbers
            
    pos = pos - 1
    
    GetNextLatexCommand = Mid$(text, posa, pos - posa + 1)

End Function
Function FindVar(svar As String) As Long
' searches for the given var or function and returns the position
Dim i As Integer
Dim Var As String
    Var = GetNextVar()
    FindVar = -1
    Do While Len(Var) > 0 And i < 200
        If Var = svar Then
            FindVar = pos
            Exit Function
        End If
        Var = GetNextVar()
        i = i + 1
    Loop
    
    If i = 200 Then FindVar = -1
End Function
Sub InsertMultSigns()
' inserts implied multiplication signs
Dim gempos As Long
Dim c As String, Var As String
    On Error Resume Next
    Do
    c = AscW(ChrByIndex(pos))
    
    If ChrByIndex(pos) = "_" Then
            FindNextBreak
    ElseIf IsNumber Then  ' tal eller pi 'And c <> 960
        If (IsLetterPos(pos + 1)) And (Not (IsLetterPos(pos - 1, False)) Or pos = 1) Then  ' c=960 introduced v1.24 because geogebra creates pi2 which should give ->pi*2
            gempos = pos
            Var = GetPrevVar(pos) ' checks whether the number can be a number at the end of a function, then no multiplication sign should be inserted
            If Not (Len(Var) > 0 And pos = gempos) Then
                InsertAfterPos "*", gempos
            End If
            pos = gempos
            pos = pos + 1 'extra
        ElseIf ChrByIndex(pos - 1) = ")" Or ChrByIndex(pos - 1) = ChrW$(960) Then
            InsertBeforePos ("*")
            pos = pos + 1 'extra
        End If
    ElseIf ChrByIndex(pos) = ")" Then
        If IsLetterPos(pos + 1) And ChrByIndex(pos + 1) <> "_" Then
            InsertAfterPos ("*")
            pos = pos + 1 'extra
        End If
    End If
    pos = pos + 1
    Loop Until pos > Len(text)
    
End Sub
Sub ReplaceSpaceByMult()
' replaces spaces with multiplication signs, but only in special cases
' letter space (

Dim p As Long
On Error GoTo slut
p = 0
Do
    p = InStr(p + 1, text, " ")
    If p = 1 Then
        text = right$(text, Len(text) - 1)
    ElseIf p = Len(text) Then
        text = Left$(text, Len(text) - 1)
    ElseIf p > 0 Then
        If IsLetterPos(p - 1, False) Then
            If IsLetterPos(p + 1, True) Or IsNumber(p + 1) Then
                text = Left$(text, p - 1) & "*" & right$(text, Len(text) - p)
            End If
        ElseIf IsNumber(p - 1) Then
            If IsLetterPos(p + 1, True) Then
                text = Left$(text, p - 1) & "*" & right$(text, Len(text) - p)
            End If
        End If
    End If
Loop While p > 0

slut:
End Sub
Sub RemoveMultSign()
' removes multiplication sign
' number*letter and )*( and 2*( or space*all

Dim p As Long
Dim C1 As String, C2 As String, c3 As String
    p = 1
    Do
        p = InStr(p, text, "*")
        If p > 0 And p < Len(text) Then
            C1 = ChrByIndex(p - 1)
            c3 = ChrByIndex(p + 1)
            If (IsNumber(p - 1) And IsLetterPos(p + 1)) Or (C1 = ")" And C2 = "(") Or C1 = " " Then
                text = Left$(text, p - 1) & " " & right$(text, Len(text) - p)
            End If
            p = p + 2
        End If
    Loop While p > 0 And p < Len(text)
    
End Sub
Function IsLetterPos(Optional ByVal position As Long = -100, Optional startbrack As Boolean = True) As Boolean
' returns whether the current position is a letter (uppercase, lowercase, Greek, ae oe aa) underscore _ or opening parenthesis

Dim c As Integer
Dim cs As String
If position = -100 Then
    position = pos ' position is optional
End If
If position < 1 Then
    IsLetterPos = False
    Exit Function
End If
cs = ChrByIndex(position)
If cs = "" Then
    c = 0
Else
    c = AscW(cs)
End If

IsLetterPos = IsLetterC(c, startbrack)

'If (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= &H391 And c <= &H3C9) Or c = 40 Or c = 95 Or c = 230 Or c = 197 Or c = 198 Or c = 216 Or c = 229 Or c = 248 Or c = 8289 Then
'' 95 - underscore_  8289 funktionstegn
'    If c = 40 And Not startbrack Then
'        IsLetterPos = False
'    Else
'        IsLetterPos = True
'    End If
'Else
'    IsLetterPos = False
'End If
End Function
Function IsLatexLetterPos(Optional ByVal position As Long = -100) As Boolean
' returns whether the current position is a letter (uppercase, lowercase, Greek, ae oe aa)
Dim c As Integer
Dim cs As String
If position = -100 Then
    position = pos ' position er kun optional
End If
If position < 1 Then
    IsLatexLetterPos = False
    Exit Function
End If
cs = ChrByIndex(position)
If cs = "" Then
    c = 0
Else
    c = AscW(cs)
End If

IsLatexLetterPos = IsLatexLetterC(c)

End Function
Function IsLetterC(c As Integer, Optional startbrack As Boolean = True) As Boolean
' returns whether character is letter (uppercase, lowercase, Greek, ae oe aa) underscore _ or opening parenthesis
' Can be used as IsLetter(AscW("a"))

If (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= &H391 And c <= &H3C9) Or c = 40 Or c = 95 Or c = 230 Or c = 197 Or c = 198 Or c = 216 Or c = 229 Or c = 248 Or c = 8289 Then
' The following was an attempt to remove pi from this one. It wasn't necessary, but maybe it will turn out to be necessary later.
'If (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= 913 And c <= 959) Or (c >= 961 And c <= 969) Or c = 40 Or c = 95 Or c = 230 Or c = 197 Or c = 198 Or c = 216 Or c = 229 Or c = 248 Or c = 8289 Then
' 95 - underscore_  8289 function symbol
    If c = 40 And Not startbrack Then
        IsLetterC = False
    Else
        IsLetterC = True
    End If
Else
    IsLetterC = False
End If

End Function
Function IsLatexLetterC(c As Integer) As Boolean
' returns whether character is letter (uppercase, lowercase, Greek, ae oe aa) Greek and ae oe aa can probably be removed
' pi removed v.1.24 since e.g. pi2 should be pi*2
' Can be used as IsLetter(AscW("a"))
If (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= 913 And c <= 959) Or (c >= 961 And c <= 969) Or c = 230 Or c = 197 Or c = 198 Or c = 216 Or c = 229 Or c = 248 Or c = 8289 Then
' 95 - underscore_  8289 function symbol
    IsLatexLetterC = True
Else
    IsLatexLetterC = False
End If

End Function
Function IsLetter(s As String, Optional startbrack As Boolean = True) As Boolean
' like IsLetterC, but with string

IsLetter = IsLetterC(AscW(s), startbrack)
End Function
Function IsLatexLetter(s As String) As Boolean
' like IsLatexLetterC, but with string

IsLatexLetter = IsLatexLetterC(AscW(s))
End Function

Function IsSquareRoot(Optional ByVal position As Long = -100) As Boolean
' true if square root, cubic og 4th or ^
Dim c As Integer
Dim cs As String
If position = -100 Then
    position = pos ' position is optional
End If
If position < 1 Then
    IsSquareRoot = False
    Exit Function
End If
cs = ChrByIndex(position)
If cs = "" Then
    c = 0
Else
    c = AscW(cs)
End If

If c = 8730 Or c = 8731 Or c = 8732 Or c = 94 Then '  8730 - squareroot symbol ,8731 cubicroot symbol 8732 4. root, 94 ^
    IsSquareRoot = True
Else
    IsSquareRoot = False
End If

End Function
Function IsNumber(Optional ByVal spos As Long) As Boolean
' returns whether current position is number

If spos = 0 Then
    spos = pos
End If

Dim c As Integer
c = AscW(ChrByIndex(spos))
If (c >= 48 And c <= 57) Or c = 960 Then  '960?  44 is , and 46 is .
    IsNumber = True
Else
    IsNumber = False
End If
End Function
Function IsNumberOrSign(Optional ByVal spos As Long) As Boolean
' returns whether the current position is a number or a sign

If spos = 0 Then
    spos = pos
End If

Dim c As Integer
c = AscW(ChrByIndex(spos))
If (c >= 48 And c <= 57) Or c = 960 Or c = 45 Then '960?  44 is , and 46 is . 45=-
    IsNumberOrSign = True
Else
    IsNumberOrSign = False
End If
End Function
Function IsFunction(text As String) As Boolean
'
If text = "cos" Or text = "sin" Or text = "tan" Or text = "acos" Or text = "asin" Or text = "atan" Or text = "cosh" Or text = "sinh" Or text = "tanh" Or text = "acosh" Or text = "asinh" Or text = "atanh" Or text = "solve" Or text = "integrate" Or text = "diff" Or text = "sqrt" Or text = "log" Or text = "ln" Then
    IsFunction = True
Else
    IsFunction = False
End If

End Function
Function IsOperator(text As String, Optional level As Integer = 0) As Boolean
' is only used for | | determination, and hieraki is not quite like findhierakibreak
'level is hieraki-level
'Hierarchy level:
'=<>| 0
':;,@?     1
'+-  2
'*/  3
'^_  4
'(   5
IsOperator = False
If text = "=" Or text = "<" Or text = ">" Or text = "|" Then
    IsOperator = True
End If

If level > 0 Then
    If text = ":" Or text = ";" Or text = "," Or text = "@" Or text = "?" Then IsOperator = True
End If
If level > 1 Then
    If text = "+" Or text = "-" Then IsOperator = True
End If
If level > 2 Then
    If text = "*" Or text = "/" Then IsOperator = True
End If
If level > 3 Then
    If text = "^" Or text = "_" Then IsOperator = True
End If
If level > 4 Then
    If text = "(" Or text = "[" Or text = "{" Then IsOperator = True
End If

End Function
Function IsAllText() As Boolean
' checks from pos to end and text only consists of letters _ and ( and 0-9
IsAllText = True
Do
    If Not (IsLetterPos(True) Or IsNumber) Then
        IsAllText = False
        Exit Do
    End If
    pos = pos + 1
Loop While pos <= Len(text)

End Function
Public Sub SetNormalBrackets()
    StartBracket = "("
    EndBracket = ")"
End Sub
Public Sub SetSquareBrackets()
    StartBracket = "["
    EndBracket = "]"
End Sub
Public Sub SetTuborgBrackets()
    StartBracket = "{"
    EndBracket = "}"
End Sub
Public Sub SetInvisibleBrackets()
    StartBracket = VBA.ChrW$(12310)
    EndBracket = VBA.ChrW$(12311)
End Sub
Private Sub Class_Initialize()
    StartBracket = "["
    EndBracket = "]"
End Sub
Function GetTextInterval(ByVal p1 As Long, ByVal p2 As Long) As String
' returns the part of the text that starts with p1 and ends at p2
    If p2 > Len(text) Then p2 = Len(text)
    If p1 < 1 Then p1 = 1
    
    If p2 >= p1 Then
        GetTextInterval = Mid$(text, p1, p2 - p1 + 1)
    Else
        GetTextInterval = ""
    End If
End Function

Sub ReplaceVar(ByVal Var As String, ByVal replvar As String, Optional StartPos As Long = 1, Optional Count As Long = -1)
' finds var in text and replaces with replvar

    Dim pos2 As Long
    
    pos2 = InStr(StartPos, text, Var)
    
    Do While pos2 > 0 And Count <> 0
    If pos2 = 1 Then
        If Not (IsLetterPos(pos2 + Len(Var), False) Or IsNumber(pos2 + Len(Var))) Then
            text = Left$(text, pos2 - 1) & Replace(text, Var, replvar, pos2, 1)
            pos2 = pos2 + Len(replvar)
        End If
    ElseIf pos2 + Len(Var) = Len(text) + 1 Then
        If Not (IsLetterPos(pos2 - 1, False)) Then
            text = Left$(text, pos2 - 1) & Replace(text, Var, replvar, pos2, 1)
            pos2 = pos2 + Len(replvar)
        End If
    Else
        If Not (IsLetterPos(pos2 - 1, False) Or IsLetterPos(pos2 + Len(Var), False) Or IsNumber(pos2 + Len(Var))) Then
            text = Left$(text, pos2 - 1) & Replace(text, Var, replvar, pos2, 1)
            pos2 = pos2 + Len(replvar)
        End If
    End If
    pos2 = InStr(pos2 + 1, text, Var)
    Count = Count - 1
    Loop
End Sub
Sub ReplaceTrigPow(ByVal Var As String, ByVal replvar As String, Optional StartPos As Long = 1, Optional Count As Long = -1)
' finds var in text and replaces with replvar
' var must be a trig function. takes into account sin^2 & chrw$(8289)

    Dim pos2 As Long
    
    pos2 = InStr(StartPos, text, Var & "^")
    
    Do While pos2 > 0 And Count <> 0
        FindNextBreak
        If pos2 = 1 Then
            If Not (IsLetterPos(pos2 + Len(Var), False) Or IsNumber(pos2 + Len(Var))) Then
                text = Left$(text, pos2 - 1) & Replace(text, Var, replvar, pos2, 1)
                pos2 = pos2 + Len(replvar)
            End If
        ElseIf pos2 + Len(Var) = Len(text) + 1 Then
            If Not (IsLetterPos(pos2 - 1, False)) Then
                text = Left$(text, pos2 - 1) & Replace(text, Var, replvar, pos2, 1)
                pos2 = pos2 + Len(replvar)
            End If
        Else
            If Not (IsLetterPos(pos2 - 1, False) Or IsLetterPos(pos2 + Len(Var), False) Or IsNumber(pos2 + Len(Var))) Then
                text = Left$(text, pos2 - 1) & Replace(text, Var, replvar, pos2, 1)
                pos2 = pos2 + Len(replvar)
            End If
        End If
        pos2 = InStr(pos2 + 1, text, Var & "^")
        Count = Count - 1
    Loop
End Sub
Function GetNextListItem(Optional ByVal spos As Long = 0, Optional Sep As String = "") As String
' returns the next part of the string up to the first separator that is not in parentheses ([{
' not finished
Dim c As String
Dim cs As String 'startbracket
Dim CE As String 'endbracket
Dim bracketlevel As Integer

If Sep = "" Then
    Sep = ListSeparator
End If

If spos < 1 Then
    spos = pos
    If spos < 1 Then
        spos = 1
        pos = 1
    End If
Else
    pos = spos
End If

If pos > Len(text) Then Exit Function

Do
    Do
        c = ChrByIndex(pos)
        pos = pos + 1
    Loop Until c = "[" Or c = "(" Or c = "{" Or c = Sep Or c = vbTab Or pos >= Len(text)
    
    If c = Sep Or c = vbTab Then
        pos = pos
        GoTo slut
    Else
        cs = c
        If c = "(" Then
            CE = ")"
        ElseIf c = "[" Then
            CE = "]"
        ElseIf c = "{" Then
          If ChrByIndex(pos) = VBA.ChrW$(9608) Then  ' curyl brackets with end bracket for piecewise defined functions
            CE = VBA.ChrW$(9508)
          Else
            CE = "}"
          End If
        End If
    End If
    bracketlevel = 1
    Do
        c = ChrByIndex(pos)
        
        If c = cs Then
            bracketlevel = bracketlevel + 1
        ElseIf c = CE Then
            bracketlevel = bracketlevel - 1
        End If
        pos = pos + 1
    Loop Until bracketlevel = 0 Or pos > Len(text) ' changed from >= to > as there was an example of an error with a string ending in )
Loop Until pos > Len(text)

slut:
GetNextListItem = Mid$(text, spos, pos - spos)
If right$(GetNextListItem, 1) = Sep Then
    GetNextListItem = Left$(GetNextListItem, Len(GetNextListItem) - 1)
End If

End Function
Sub SplitNonBracket(ByRef Arr As Variant, Optional Sep As String = "")
    Dim s As String, t As String
    pos = 1
    Do
        s = GetNextListItem(pos, Sep)
        If s <> "" Then
            t = t & s & "$_q"
        End If
    Loop While s <> ""
    If t <> "" Then t = Left$(t, Len(t) - 3)
    Arr = Split(t, "$_q")
    
End Sub

Sub RemoveFaultySeparator()
' removes 23. -> 23  og .34 -> 0.34  og x*.34 -> x*0.34  og 34.*x -> 34*x
Dim spos As Long
    If pos < 1 Then pos = 1
    spos = InStr(pos, text, ".")
    Do While spos > 0
        If spos = 1 Then
            text = "0" & text
        ElseIf spos = Len(text) Then
            text = Left$(text, Len(text) - 1)
        ElseIf Not (IsNumber(spos + 1)) Then
            text = Left$(text, spos - 1) & right$(text, Len(text) - spos)
        ElseIf Not (IsNumber(spos - 1)) Then
            text = Left$(text, spos - 1) & "0" & right$(text, Len(text) - spos + 1)
        End If
        spos = InStr(spos + 1, text, ".")
    Loop
End Sub
Sub InsertTextAt(it As String, Optional ByVal ipos As Long = -1)
' inserts it at position ipos
    If ipos < 0 Then ipos = pos

    If ipos < 2 Then
        text = it & text
    ElseIf ipos >= Len(text) + 1 Then
        text = text & it
    Else
        text = Left$(text, ipos - 1) & it & right$(text, Len(text) - ipos + 1)
    End If
End Sub
Sub RemoveChar(Optional ind As Long = -1)
' removes char at position in or pos if none specified

    If ind < 0 Then ind = pos

    If ind < 2 Then
        text = right$(text, Len(text) - 1)
    ElseIf ind = Len(text) Then
        text = Left$(text, Len(text) - 1)
    Else
        text = Left$(text, ind - 1) & right$(text, Len(text) - ind)
    End If
End Sub

Function ValidateBrackets() As Boolean
' indicates whether parentheses are specified correctly
' however only simple method that does not catch everything

ValidateBrackets = CountText("(") = CountText(")")

End Function

Function CountText(s As String) As Integer
' counts how many times s appears
Dim p As Long
Dim i As Long
p = 1
Do While p > 0 And p <= Len(text)
    p = InStr(p, text, s)
    i = i + 1
    If p > 0 Then p = p + 1
Loop
i = i - 1
CountText = i
End Function
Sub ReplaceStringAt(s As String, Optional ByVal p As Long = -1)
' overwrites text with s starting at position p

If p < 0 Then p = pos

If p < 2 Then
    text = s & right$(text, Len(text) - Len(s))
ElseIf p > Len(text) Then
    text = text & s
ElseIf p > Len(text) - Len(s) Then
    text = Left$(text, p - 1) & s
Else
    text = Left$(text, p - 1) & s & right$(text, Len(text) - Len(s) - p + 1)
End If

End Sub

Sub ConvertDecSeparator()
' converts , to . where , surrounded by numbers
    Dim p As Long
        p = 0
        Do
            p = InStr(p + 1, text, ",")
            If p > 0 Then
                If IsNumber(p - 1) And IsNumber(p + 1) Then
                    ReplaceStringAt ".", p
                End If
            End If
        Loop While p > 0

End Sub
Sub RemoveSpaceAfterSubSuperscript(tegn As String)
' removes spaces after e.g. f_a (x)
' so you can otherwise replace spaces with *
On Error GoTo slut

Do
    pos = InStr(pos + 1, text, tegn)
    If pos = 1 Or pos = Len(text) Then
    ElseIf pos > 0 Then
'        pos = pos + 1
        FindNextBreak (False)
        If ChrByIndex(pos + 1) = " " And ChrByIndex(pos + 2) = "(" Then
             text = Left$(text, pos) & right$(text, Len(text) - pos - 1)
        End If
    End If
Loop While pos > 0

slut:
End Sub
Sub AddSpaceAfterSubSuperscript(tegn As String)
' adds spaces after e.g. f_a (x)
' so you can otherwise replace spaces with *
Dim p As Long
On Error GoTo slut
p = pos
pos = 1
Do
    pos = InStr(pos, text, tegn)
    If pos = Len(text) Then
        Exit Sub
    ElseIf pos = 1 Then
        pos = pos + 1
    ElseIf pos > 0 Then
        pos = pos + 1
        If ChrByIndex(pos) = StartBracket Then
            GetNextBracketContent
            pos = pos - 1
        Else
            FindHierakiBreak 5
        End If
        If ChrByIndex(pos + 1) = EndBracket Then pos = pos + 1
        
'        FindNextBreak (True)
        If ChrByIndex(pos + 1) = "(" Then
             text = Left$(text, pos) & " " & right$(text, Len(text) - pos)
        End If
    End If
Loop While pos > 0
pos = p
slut:
End Sub

Sub InsertFunctionBrackets()
' inserts parentheses after e.g. log 2 -> log(2)
' special function symbol 8289 is not removed yet as it will be used by other functions
Dim p As Long
Dim c As String
Dim sb As String
Dim eb As String

p = 1
Do
    p = InStr(p, text, VBA.ChrW$(8289))
    If p > 0 Then
        p = p + 1
        c = ChrByIndex(p)
        If c = "(" Then
        ElseIf c = VBA.ChrW$(12310) Then ' if invisible bracket
            If ChrByIndex(p + 1) <> "(" Then ' invisible parentheses may be unwanted. It is assumed if the parentheses immediately after
                pos = p
                sb = StartBracket
                eb = EndBracket
                StartBracket = VBA.ChrW$(12310)
                EndBracket = VBA.ChrW$(12311)
                FindEndOfBracket
                StartBracket = sb
                EndBracket = eb
                ReplaceStringAt "(", p
                ReplaceStringAt ")", pos - 1
            End If
        Else 'hvis ingen parentes
            InsertTextAt "(", p
            pos = p + 1
            InsertTextAt ")", FindHierakiBreak(3) + 1
        End If
    End If

Loop While p > 0

End Sub

Function ContainsVar(Var As String) As Boolean
' determines if variable var is in text
Dim ivar As String

ContainsVar = False
pos = 0

Do
    ivar = GetNextVar(pos + 1)
    If ivar = Var Then
        ContainsVar = True
    End If
Loop While ivar <> "" And pos < Len(text)

End Function

Sub ConvertTrigPow(Trig As String)
' converts e.g. sin^2(x) -> sin(x)^2
Dim ea As ExpressionAnalyser
Dim pow As String
Dim pos2 As Long
pos = 1
Do
    pos = InStr(pos, text, Trig & "^")
    If pos > 0 Then
        pos = pos + Len(Trig) + 1
        pos2 = InStr(pos, text, VBA.ChrW$(8289))
        If pos2 = 0 Then
            Set ea = New ExpressionAnalyser
            ea.SetNormalBrackets
            ea.text = text
            ea.pos = pos
            pos2 = ea.FindEndOfBracket
        End If
        pow = Mid$(text, pos, pos2 - pos)
        text = Left$(text, pos - 2) & right$(text, Len(text) - pos2 + 1)
        InsertTextAt "^" & pow, FindEndOfBracket
        
    End If
Loop While pos > 0

End Sub

Function FindNextGreaterOrLower(Optional StartPos As Long = 0)
' finds next < or >
Dim p1 As Long, p2 As Long
If StartPos = 0 Then
    StartPos = pos
End If

p1 = InStr(StartPos, text, ">")
p2 = InStr(StartPos, text, "<")

If (p1 < p2 And p1 > 0 And p2 > 0) Or (p1 > 0 And p2 < 1) Then
    pos = p1
    FindNextGreaterOrLower = p1
ElseIf (p2 < p1 And p2 > 0 And p1 > 0) Or (p2 > 0 And p1 < 1) Then
    pos = p2
    FindNextGreaterOrLower = p2
Else
    FindNextGreaterOrLower = 0
End If

End Function
Function WhatsNextBracket() As String
' returns next startbracket [,( or {
Dim p1 As Long, p2 As Long, p3 As Long

p1 = InStr(pos, text, "(")
p2 = InStr(pos, text, "[")
p3 = InStr(pos, text, "{")

If p1 > 0 Then
    WhatsNextBracket = "("
End If
If p2 > 0 Then
    If WhatsNextBracket = "" Then WhatsNextBracket = "["
    If p2 < p1 Then WhatsNextBracket = "["
End If
If p3 > 0 Then
    If WhatsNextBracket = "" Then WhatsNextBracket = "{"
    If p3 < p1 And p3 < p2 And p1 > 0 And p2 > 0 Then WhatsNextBracket = "["
    If p3 < p1 And p2 = 0 And p1 > 0 Then WhatsNextBracket = "["
    If p3 < p2 And p1 = 0 And p2 > 0 Then WhatsNextBracket = "["
End If

End Function

Sub SetToNextBracket()
Dim b As String
b = WhatsNextBracket()
    If b = "(" Then
        SetNormalBrackets
    ElseIf b = "[" Then
        SetSquareBrackets
    ElseIf b = "{" Then
        SetTuborgBrackets
    Else
        SetNormalBrackets
    End If
End Sub

Function IsUnit(ByVal E As String) As Boolean
' checks if variable is unit. Only pure units
' used for Latex so units can be written in non-italic
' not finished
Dim p As String

If E = "kWh" Or E = "min" Then ' approve those that can be confused with prefix
    IsUnit = True
    GoTo slut
End If

If Len(E) > 1 Then 'remove any prefix
p = Left$(E, 1)
If p = "k" Or p = "M" Or p = "G" Or p = "T" Or p = "P" Or p = "c" Or p = "m" Or p = VBA.ChrW$(956) Or p = "n" Or p = "p" Or p = "a" Or p = "" Then
    E = right$(E, Len(E) - 1)
End If
End If

If E = "g" Or E = "m" Or E = "s" Or E = "K" Or E = "mol" Or E = "cd" Or E = "A" Or E = "V" Or E = "W" Or E = "J" Or E = "cal" Or E = "N" Or E = "T" Or E = "H" Or E = "timer" Or E = "Ohm" Then
    IsUnit = True
Else
    IsUnit = False
End If

slut:
End Function

Function SubString(StartPos As Long, endpos As Long)

If StartPos < 1 Then StartPos = 1

If endpos > Length Then
    endpos = Length
End If

If endpos < StartPos Then
    SubString = ""
Else
    SubString = Mid$(text, StartPos, endpos - StartPos + 1)
End If

End Function
Sub SetBrackets(b As Integer)
    If b = 40 Or b = 41 Then
        SetNormalBrackets
    ElseIf b = 91 Or b = 93 Then
        SetSquareBrackets
    ElseIf b = 123 Or b = 125 Then
        SetTuborgBrackets
    ElseIf b = 12310 Or b = 12311 Then
        StartBracket = VBA.ChrW$(12310)
        EndBracket = VBA.ChrW$(12311)
    End If

End Sub

Sub SetBracketsS(b As String)
    If b = "(" Or b = ")" Then
        SetNormalBrackets
    ElseIf b = "[" Or b = "]" Then
        SetSquareBrackets
    ElseIf b = "{" Or b = "}" Then
        SetTuborgBrackets
    ElseIf b = VBA.ChrW$(12310) Or b = VBA.ChrW$(12311) Then
        StartBracket = VBA.ChrW$(12310)
        EndBracket = VBA.ChrW$(12311)
    End If
End Sub

Sub SaveBrackets()
    SaveStartBracket = StartBracket
    SaveEndBracket = EndBracket
End Sub
Sub RestoreBrackets()
    StartBracket = SaveStartBracket
    EndBracket = SaveEndBracket
End Sub
Function IsSuperScriptNoC(c As Integer) As Boolean
' returns whether character is superscript number 0,1,2,3 or superscript + - ( )
' Can be used as IsLetter(AscW("a"))
If c = 8314 Or c = 8315 Or c = 8317 Or c = 8318 Or c = 8304 Or c = 185 Or c = 178 Or c = 179 Or (c >= 8308 And c <= 8318) Then
    IsSuperScriptNoC = True
Else
    IsSuperScriptNoC = False
End If
End Function
Function IsSuperScriptNo(s As String) As Boolean
' like IsLetterC, but with string
    IsSuperScriptNo = IsSuperScriptNoC(AscW(s))
End Function
Function IsSuperScriptNoPos(Optional ByVal position As Long = -100) As Boolean
' returns whether the current position is a letter (uppercase, lowercase, Greek, ae oe aa) underscore _ or opening parenthesis
    Dim c As Integer
    Dim cs As String
    If position = -100 Then
        position = pos ' position is optional
    End If
    If position < 1 Then
        IsSuperScriptNoPos = False
        Exit Function
    End If
    cs = ChrByIndex(position)
    If cs = "" Then
        c = 0
    Else
        c = AscW(cs)
    End If

    IsSuperScriptNoPos = IsSuperScriptNoC(c)

End Function
