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
Public pos As Integer

Public Function GetNextchr() As String
' returnerer næste bogstav i strengen og fjerner denne
GetNextchr = Left(text, 1)
text = right(text, Len(text) - 1)

End Function

Public Property Get NextChr() As Variant
    NextChr = Left(text, 1)
End Property
Public Property Get Length() As Integer
    Length = Len(text)
End Property

Public Function FindNextBracket(StartPos As Integer) As Integer
Dim pos1 As Integer
Dim pos2 As Integer

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
Public Function FindPrevBracket(StartPos As Integer) As Integer
Dim pos1 As Integer
Dim pos2 As Integer
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
Public Function FindEndOfBracket() As Integer
' finder næste parantes og returnerer postion efter slutparentes
    GetNextBracketContent
    FindEndOfBracket = pos
End Function
Public Function GetNextBracketContent(Optional ByVal StartPos As Integer) As String
' returnerer indholdet af næste parantes
' tom hvis paranteser ikke matcher
' pos ændres til efter parantesen

Dim pos1 As Integer ' var begge String i v. 1.09 og før
Dim pos2 As Integer
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
    
GetNextBracketContent = Mid(text, pos1 + 1, pos2 - pos1 - 1)
pos = pos2 + 1
End Function
Public Function GetPrevBracketContent(Optional ByVal StartPos As Integer) As String
' returnerer indholdet af parantes før pos
' tom hvis paranteser ikke matcher
' pos ændres til første parantestegn

Dim pos1 As String
Dim pos2 As String
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
    
GetPrevBracketContent = Mid(text, pos2 + 1, pos1 - pos2 - 1)
pos = pos2 + 1
End Function

Public Property Get ChrByIndex(ByVal pos As Integer) As String
On Error GoTo Fejl
    If pos <= Len(text) And pos > 0 Then
        ChrByIndex = Mid(text, pos, 1)
    Else
        ChrByIndex = VBA.ChrW(0)
    End If
GoTo Slut
Fejl:
        ChrByIndex = VBA.ChrW(0)
Slut:
End Property
Public Function FindNextBreak(Optional stopforbrackets As Boolean = False) As Integer
' finder næste gang hvor der burde være parantes.
' parentesen skal være efter den returnerede position
' dvs. finder mellemrum, +-*/^=<>):;,@?   paranteser afslutter nu før gjorde de ikke men det gav problemer med f.eks. (3+sqrt3)^2
' startparantes(c=40) afslutter ikke da det giver problemer med funktioner f.eks. 2^log(3)
' muligvis problemer med d/dx notation
' komma skal med da der kan være maximafunktioner i udtrykket solve(2=9^2x,x) f.eks.
' stopforbrackets er deaktiveret da det formodes det ikke er nødvendigt, men måske dx notation ved integraler?
' stop for brackets reaktiveret for at virke i converttowordsymbols f_a(1) -> f_a (1)
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
Public Function FindHierakiBreak(level As Integer, Optional komma As Boolean = True) As Integer
' finder næste operator afhængig af angivet hierakilevel.
'Hierarki level:
'=<>)] 0 sumtegn og produkttegn
':;,@? mellemrum    1   kun komma hvis komma=true
'+-  2
'*/  3
'^_  4
'({[|   5
' den returnerede værdi er positionen for sidste tegn inden operatoren
' level 0 stopper kun for operatorer på level 0
' level 1 stopper for operatorer på level 0-1
' level 2 stopper for operatorer på level 0-2
' ...
' findnextbreak svarer til level 4
' findnextbreak bracket:=true svarer til level 5
    
    Dim c As Integer, i As Integer
    
    Do
    i = i + 1 ' sikkerhed for intinite loop
    c = AscW(ChrByIndex(pos))
    If (c >= 60 And c <= 62) Or c = 41 Or c = 93 Or c = 8776 Or c = 8719 Or c = 8721 Or c = 8594 Or c = 8592 Or c = 8596 Or c = 8651 Or c = 8652 Then Exit Do ' 41=) , 93=] , 60=<, 61== , 62=> , 8776=cirkaligmed, 8719=produkttegn , 8721=sumtegn, 8594=-> , 8592=<-, 8596=<->, 8651,8652=kemidblpil
    
    If level > 0 Then
        If c = 32 Or c = 58 Or c = 59 Or c = 63 Or c = 64 Or c = 38 Then Exit Do    ' 32=mellemrum, 58=: , 59=; , 44=, 63=? , 64=@ , 38=&
        If komma And c = 44 Then Exit Do
    End If
    If level > 1 Then
        If c = 43 Or c = 45 Then ' 43=+, 45=-
            If ChrByIndex(pos - 1) <> "E" Then Exit Do ' E- og E+ skal ikke stoppe
            If pos = 2 Then Exit Do
            If Not (IsNumber(pos - 2)) Then Exit Do
        End If
    End If
    If level > 2 Then
        If c = 42 Or c = 215 Or c = 47 Or c = 183 Then Exit Do ' 42=*, 47=/ , 215=\times, 183=\cdot (183 tilføjet v. 1.07 pga Latexkonv)
    End If
    If level > 3 Then
        If c = 94 Or c = 95 Then Exit Do ' 94=^, 95=_
    End If
    If level > 4 Then
        If c = 40 Or c = 91 Or c = 123 Or c = 124 Or c = 12310 Then Exit Do ' 40=(, 91=[ , 123={ , 124=|
    End If
    
    If c = 40 Then ' hvis parantesstart hoppes til parantesslut
        SaveBrackets
        SetNormalBrackets
        GetNextBracketContent
        RestoreBrackets
        pos = pos - 1
    End If
    If c = 91 Then ' hvis [ parantesstart hoppes til parantesslut ]
        SaveBrackets
        SetSquareBrackets
        GetNextBracketContent
        RestoreBrackets
        pos = pos - 1
    End If
    If c = 123 Then ' hvis {
        SaveBrackets
        SetTuborgBrackets
        GetNextBracketContent
        RestoreBrackets
        pos = pos - 1
    End If
    If c = 12310 Then ' hvis {
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
    If c = 41 Or c = 93 Or c = 125 Then ' hvis parantesslut hoppes til parantesstart )]}
        SaveBrackets
        SetBrackets (c)
        GetPrevBracketContent
        pos = pos - 1
        RestoreBrackets
    End If
    
    If (c >= 60 And c <= 62) Or c = 41 Or c = 93 Or c = 8776 Or c = 8289 Then Exit Do ' 41=) , 93=] , 60=<, 61== , 62=> , 8776=cirkaligmed, 8289=funktionstegn
    
    If level > 0 Then
        If c = 32 Or c = 58 Or c = 59 Or c = 63 Or c = 64 Or c = 38 Then Exit Do ' 32=mellemrum, 58=: , 59=; , 44=, 63=? , 64=@, 38=&
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
Public Function FindPrevBreak() As Integer
' finder forrige gang hvor der burde være parantes
' dvs. finder mellemrum, +-*/=<>)]}:;,?@   paranteser afslutter nu før gjorde de ikke men det gav problemer med f.eks. (3+sqrt3)^2
' 8289 er symbol efter funktion
' ^ fjernet da det ellers ikke virkede for f.eks. sin(x)^2/pi
    
    Dim c As Integer
    
    
    Do
    c = AscW(ChrByIndex(pos))
    If c = 41 Or c = 93 Or c = 125 Or c = 12311 Then ' hvis parantesslut hoppes til parantesstart )]}
        SaveBrackets
        SetBrackets (c)
        GetPrevBracketContent
        pos = pos - 1
        RestoreBrackets
    End If
    pos = pos - 1
    Loop Until pos <= 0 Or c = 32 Or c = 38 Or (c >= 42 And c <= 45) Or c = 40 Or c = 47 Or (c >= 58 And c <= 64) Or c = 9618 Or c = 91 Or c = 123 Or c = 12310
    If pos <= 0 Then
        pos = 1
    Else
        pos = pos + 2
    End If
    FindPrevBreak = pos

End Function

Public Sub InsertUnderstoodBracketPair()
    'indsætter parantes lige foran nuværende position og finder hvor slutparantes skal være
    Dim c As Integer
    text = Left(text, pos - 1) & StartBracket & right(text, Len(text) - pos + 1)
    pos = pos + 1
    c = AscW(ChrByIndex(pos)) ' +2 istedet for +1 da det kan give () f.eks. ved x^-2 ->x^()-2
    If Not (c = 40 Or c = 91 Or c = 123 Or c = 124 Or c = 12310) Then pos = pos + 1 ' 40=(, 91=[ , 123={ , 124=|

    If pos > Length Then
        pos = Length
        text = text & EndBracket
    Else
        FindHierakiBreak 3
        ' indsætter parantes efter nuværende pos
        text = Left(text, pos) & EndBracket & right(text, Len(text) - pos)
    End If
'    FindNextBreak
    
End Sub
Public Sub InsertUnderstoodBracketPairBefore()
' hvis ikke allerede er der
Dim gpos As Integer
    gpos = pos
    
    ' fjern evt. mellemrum foran / eller ^
    If ChrByIndex(gpos) = " " Then
        RemoveChar (gpos)
        gpos = gpos - 1
        pos = pos - 1
    End If
    
    FindPrevBreak
    
    
    If (ChrByIndex(gpos) <> ")" Or ChrByIndex(pos) <> "(") Then
        text = Left(text, pos - 1) & StartBracket & right(text, Len(text) - pos + 1)
        text = Left(text, gpos + 1) & EndBracket & right(text, Len(text) - gpos - 1)
    End If
    
    
End Sub
Public Sub InsertBracketAfter(s As String)
' indsætter underfortået parantes efter symbol s i hele teksten
' bruges til ^ og / . Hvis der allerede er parantes efter s gøres intet
    Dim gempos As Integer
    pos = 1
    Do
    pos = InStr(pos, text, s)
    If pos > 0 Then
        If Mid(text, pos + 1, 1) <> "(" Then
'            If Not (IsNumber(pos + 1) And ChrByIndex(pos + 2) = " ") Then ' indsæt kun hvis ikke simpel tal
                pos = pos + 1
                gempos = pos
                InsertUnderstoodBracketPair
                pos = gempos + 1
'            Else
'                text = Left(text, pos + 1) & "*" & Right(text, Len(text) - pos - 2) 'løser problem med 1/5 2 bliver til 1/52 istedet for 1/5*2. Måske bedre altid at lave parantes da output ikke ses
'                pos = pos + 3
'            End If
        Else
            gempos = pos
            GetNextBracketContent pos + 1 ' det skal håndteres at notationen kan være fx 2x/(x+1)(x-1)  i Word. Da divideres med begge parenteser når der ikke er angivet et gangetegn. Tilføjet i version 1.09. Problemet optræder lidt sporadisk da der ikke oversættes konsekvent på denne måde til lineær form. Samme lingning der konveretes frem og tilbage kan være med eller uden ekstra parantes
            If Mid(text, pos, 1) = "(" Then
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
' indsætter underfortået parantes før symbol s i hele teksten
' bruges til ^ og / . Hvis der allerede er parantes før s gøres intet
    Dim gempos As Integer
    gempos = 1
    Do
    pos = InStr(gempos, text, s)
    gempos = pos + 1
    If pos > 0 Then
'        If Mid(text, pos - 1, 1) <> ")" Then
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
Public Function FindChr(c As String, Optional ByVal StartPos As Integer) As Integer
    ' finder næste forekomst af c efter pos og sætter position til denne
    ' hvis den ikke findes ændres pos ikke og der returneres -1
    If StartPos < 1 Then
        StartPos = pos
    End If
    FindChr = InStr(StartPos, text, c)
    If FindChr > 0 Then pos = FindChr
End Function
Public Sub InsertBeforePos(s As String, Optional position As Integer)
    ' indsætter streng s midt i text-strengen efter pos
    If position < 1 Then
        position = pos
    End If
    If position = 1 Then
        text = s & text
    Else
        text = Left(text, position - 1) & s & right(text, Len(text) - position + 1)
    End If
End Sub
Public Sub InsertAfterPos(s As String, Optional position As Integer)
    If position < 1 Then
        position = pos
    End If
    If position = Len(text) Then
        text = text & s
    Else
        text = Left(text, position) & s & right(text, Len(text) - position)
    End If
End Sub
Public Function GetNextVar(Optional StartPos As Integer) As String
' finder næste variabel eller funktion
'Første tegn skal være bogstav dernæst kan der være bogstav,tal eller _ eller [
' navnet returneres og position sættes til sidste position i variablen
Dim posa As Integer
Dim sb As String
Dim sb2 As String
    
    If StartPos > 0 Then
        pos = StartPos
    End If
    If pos < 1 Then pos = 1
    Dim c As Integer
    
    Do ' find bogstav
        If pos > Len(text) Then
            GetNextVar = ""
            Exit Function
        End If
        c = AscW(ChrByIndex(pos))
        pos = pos + 1
'    Loop Until (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)
    Loop While Not (IsLetterC(c, False))
'    c = AscW(ChrByIndex(pos))
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
'    Loop While pos <= Len(text) And ((c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= 48 And c <= 57) Or c = 95)
    Loop While pos <= Len(text) And ((c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= 48 And c <= 57) Or c = 95)
    ' her må variabel gerne indehold tal
            
    pos = pos - 1
    
    GetNextVar = Mid(text, posa, pos - posa + 1)

End Function
Public Function GetPrevVar(Optional StartPos As Integer) As String
' finder variabel eller funktion før positionen
'Første tegn skal være bogstav eller tal dernæst kan der være bogstav,tal eller _
' navnet returneres og position sættes til sidste position i variablen
' sidste tegn må ikke være tal (altså første tegn i variablen)
    
    If StartPos > 0 Then
        pos = StartPos
    End If

    Dim c As Integer
    
    Do ' find bogstav
        If pos = 0 Then
            GetPrevVar = ""
            Exit Function
        End If
        c = AscW(ChrByIndex(pos))
        pos = pos - 1
    Loop Until (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)
    
    Do ' find hvor variabel slutter
        c = AscW(ChrByIndex(pos))
        pos = pos - 1
    Loop While pos > 0 And ((c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= 48 And c <= 57) Or c = 95)
    
    pos = pos + 1
    GetPrevVar = GetNextVar(pos)
            
    
End Function
Public Function GetNextLatexCommand(Optional StartPos As Integer) As String
' finder næste latex kommando
'Første tegn skal være bogstav dernæst kan der være bogstav,tal eller _ eller [
' navnet returneres og position sættes til sidste position i variablen
Dim posa As Integer
Dim sb As String
Dim sb2 As String
    
    If StartPos > 0 Then
        pos = StartPos
    End If
    If pos < 1 Then pos = 1
    Dim c As Integer
    
    Do ' find bogstav
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
'    Loop While pos <= Len(text) And ((c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= 48 And c <= 57) Or c = 95)
    Loop While pos <= Len(text) And ((c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= 48 And c <= 57) Or c = 95)
    ' her må variabel gerne indehold tal
            
    pos = pos - 1
    
    GetNextLatexCommand = Mid(text, posa, pos - posa + 1)

End Function
Function FindVar(svar As String) As Integer
' searches for the given var or function and returns the position
Dim i As Integer
Dim var As String
    var = GetNextVar()
    FindVar = -1
    Do While Len(var) > 0 And i < 200
        If var = svar Then
            FindVar = pos
            Exit Function
        End If
        var = GetNextVar()
        i = i + 1
    Loop
    
    If i = 200 Then FindVar = -1
End Function
Sub InsertMultSigns()
' indsætter underforståede gangetegn
Dim navn As String
Dim c1 As Integer, c2 As Integer, gempos As Integer
Dim c As String, var As String
    On Error Resume Next
    Do
    c = AscW(ChrByIndex(pos))
    
    If ChrByIndex(pos) = "_" Then
            FindNextBreak
    ElseIf IsNumber Then ' tal eller pi
        If IsLetterPos(pos + 1) And (Not (IsLetterPos(pos - 1, False)) Or pos = 1) Then
            gempos = pos
            var = GetPrevVar(pos) ' checker om tallet kan være en tal i slutningen af en funktion, så skal der ikke indsættes gangetegn
            If Not (Len(var) > 0 And pos = gempos) Then
                InsertAfterPos "*", gempos
            End If
            pos = gempos
            pos = pos + 1 'extra
        ElseIf ChrByIndex(pos - 1) = ")" Then
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
' erstatter mellemrum med gangetegn, men kun i specielle tilfælde
' bogstav mellemrum (

Dim p As Integer
On Error GoTo Slut
p = 0
Do
    p = InStr(p + 1, text, " ")
    If p = 1 Then
        text = right(text, Len(text) - 1)
    ElseIf p = Len(text) Then
        text = Left(text, Len(text) - 1)
    ElseIf p > 0 Then
        If IsLetterPos(p - 1, False) Then
            If IsLetterPos(p + 1, True) Or IsNumber(p + 1) Then
                text = Left(text, p - 1) & "*" & right(text, Len(text) - p)
            End If
        ElseIf IsNumber(p - 1) Then
            If IsLetterPos(p + 1, True) Then
                text = Left(text, p - 1) & "*" & right(text, Len(text) - p)
            End If
        End If
    End If
Loop While p > 0

Slut:
End Sub
Sub RemoveMultSign()
' fjerner gangetegn
' tal*bogstav samt )*( og 2*( eller mellemrum*alt

Dim p As Integer
Dim c1 As String, c2 As String, c3 As String
    p = 1
    Do
        p = InStr(p, text, "*")
        If p > 0 And p < Len(text) Then
            c1 = ChrByIndex(p - 1)
            c3 = ChrByIndex(p + 1)
            If (IsNumber(p - 1) And IsLetterPos(p + 1)) Or (c1 = ")" And c2 = "(") Or c1 = " " Then
                text = Left(text, p - 1) & " " & right(text, Len(text) - p)
            End If
            p = p + 2
        End If
    Loop While p > 0 And p < Len(text)
    
End Sub
Function IsLetterPos(Optional ByVal position As Integer = -100, Optional startbrack As Boolean = True) As Boolean
' returnerer om nuværende position er bogstav (stort, småt, græsk, æøå) underscore _ eller startparantes
Dim c As Integer
Dim cs As String
If position = -100 Then
    position = pos ' position er kun optional
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
Function IsLatexLetterPos(Optional ByVal position As Integer = -100) As Boolean
' returnerer om nuværende position er bogstav (stort, småt, græsk, æøå)
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
' returnerer om tegn er bogstav (stort, småt, græsk, æøå) underscore _ eller startparantes
' Kan bruges som IsLetter(AscW("a"))
If (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= &H391 And c <= &H3C9) Or c = 40 Or c = 95 Or c = 230 Or c = 197 Or c = 198 Or c = 216 Or c = 229 Or c = 248 Or c = 8289 Then
' 95 - underscore_  8289 funktionstegn
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
' returnerer om tegn er bogstav (stort, småt, græsk, æøå) græsk og æøå kan vel fjernes
' Kan bruges som IsLetter(AscW("a"))
If (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= &H391 And c <= &H3C9) Or c = 230 Or c = 197 Or c = 198 Or c = 216 Or c = 229 Or c = 248 Or c = 8289 Then
' 95 - underscore_  8289 funktionstegn
    IsLatexLetterC = True
Else
    IsLatexLetterC = False
End If

End Function
Function IsLetter(s As String, Optional startbrack As Boolean = True) As Boolean
' som IsLetterC, men med streng

IsLetter = IsLetterC(AscW(s), startbrack)
End Function
Function IsLatexLetter(s As String) As Boolean
' som IsLatexLetterC, men med streng

IsLatexLetter = IsLatexLetterC(AscW(s))
End Function

Function IsSquareRoot(Optional ByVal position As Integer = -100) As Boolean
' true if square root, cubic og 4th or ^
Dim c As Integer
Dim cs As String
If position = -100 Then
    position = pos ' position er kun optional
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

If c = 8730 Or c = 8731 Or c = 8732 Or c = 94 Then '  8730 - kvadratrodstegn ,8731 kubikrodstegn 8732 4. rod, 94 ^
    IsSquareRoot = True
Else
    IsSquareRoot = False
End If

End Function
Function IsNumber(Optional ByVal spos As Integer) As Boolean
' returnerer om nuværende position er tal

If spos = 0 Then
    spos = pos
End If

Dim c As Integer
c = AscW(ChrByIndex(spos))
If (c >= 48 And c <= 57) Or c = 960 Then  '960?  44 er , og 46 er .
    IsNumber = True
Else
    IsNumber = False
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
' bruges kun til | | bestemmelse, og hieraki er ikke helt som findhierakibreak
'level er hieraki-level
'Hierarki level:
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
' checker fra pos til slut og text kun består af bogstaver _ og ( og 0-9
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
    StartBracket = VBA.ChrW(12310)
    EndBracket = VBA.ChrW(12311)
End Sub
Private Sub Class_Initialize()
    StartBracket = "["
    EndBracket = "]"
End Sub
Function GetTextInterval(ByVal p1 As Integer, ByVal p2 As Integer) As String
' giver den af teksten der starter med p1 og slutter ved p2
    If p2 > Len(text) Then p2 = Len(text)
    If p1 < 1 Then p1 = 1
    
    If p2 >= p1 Then
        GetTextInterval = Mid(text, p1, p2 - p1 + 1)
    Else
        GetTextInterval = ""
    End If
End Function

Sub ReplaceVar(ByVal var As String, ByVal replvar As String, Optional StartPos As Integer = 1, Optional Count As Integer = -1)
' finder var i text og erstatter med replvar

    Dim pos2 As Integer
    
    pos2 = InStr(StartPos, text, var)
    
    Do While pos2 > 0 And Count <> 0
    If pos2 = 1 Then
        If Not (IsLetterPos(pos2 + Len(var), False) Or IsNumber(pos2 + Len(var))) Then
            text = Left(text, pos2 - 1) & Replace(text, var, replvar, pos2, 1)
            pos2 = pos2 + Len(replvar)
        End If
    ElseIf pos2 + Len(var) = Len(text) + 1 Then
        If Not (IsLetterPos(pos2 - 1, False)) Then
            text = Left(text, pos2 - 1) & Replace(text, var, replvar, pos2, 1)
            pos2 = pos2 + Len(replvar)
        End If
    Else
        If Not (IsLetterPos(pos2 - 1, False) Or IsLetterPos(pos2 + Len(var), False) Or IsNumber(pos2 + Len(var))) Then
            text = Left(text, pos2 - 1) & Replace(text, var, replvar, pos2, 1)
            pos2 = pos2 + Len(replvar)
        End If
    End If
    pos2 = InStr(pos2 + 1, text, var)
    Count = Count - 1
    Loop
End Sub

Function GetNextListItem(Optional spos As Integer, Optional sep As String = "") As String
' returnerer næste del af strengen indtil første separator som ikke er i parantes ([{
' ikke færdig
Dim c As String
Dim cs As String 'startbracket
Dim ce As String 'endbracket
Dim bracketlevel As Integer

If sep = "" Then
    sep = ListSeparator
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

Do

Do
    c = ChrByIndex(pos)
    pos = pos + 1
Loop Until c = "[" Or c = "(" Or c = "{" Or c = sep Or c = vbTab Or pos >= Len(text)

If c = sep Or c = vbTab Then
    GoTo Slut
Else
    cs = c
    If c = "(" Then
        ce = ")"
    ElseIf c = "[" Then
        ce = "]"
    ElseIf c = "{" Then
      If ChrByIndex(pos) = VBA.ChrW(9608) Then  ' tuborgparentes uden ende til gaffelforskrifter
        ce = VBA.ChrW(9508)
      Else
        ce = "}"
      End If
    End If
End If
bracketlevel = 1
Do
    c = ChrByIndex(pos)
    
    If c = cs Then
        bracketlevel = bracketlevel + 1
    ElseIf c = ce Then
        bracketlevel = bracketlevel - 1
    End If
    pos = pos + 1
Loop Until bracketlevel = 0 Or pos > Len(text) ' ændret fra >= til > da der var eksempel på fejl med streng der slutter på )

Loop Until pos > Len(text)

Slut:
GetNextListItem = Mid(text, spos, pos - spos)
If right(GetNextListItem, 1) = sep Then
    GetNextListItem = Left(GetNextListItem, Len(GetNextListItem) - 1)
End If

End Function

Sub RemoveFaultySeparator()
' fjerner 23. -> 23  og .34 -> 0.34  og x*.34 -> x*0.34  og 34.*x -> 34*x
Dim spos As Integer
    If pos < 1 Then pos = 1
    spos = InStr(pos, text, ".")
    Do While spos > 0
        If spos = 1 Then
            text = "0" & text
        ElseIf spos = Len(text) Then
            text = Left(text, Len(text) - 1)
        ElseIf Not (IsNumber(spos + 1)) Then
            text = Left(text, spos - 1) & right(text, Len(text) - spos)
        ElseIf Not (IsNumber(spos - 1)) Then
            text = Left(text, spos - 1) & "0" & right(text, Len(text) - spos + 1)
        End If
        spos = InStr(spos + 1, text, ".")
    Loop
End Sub
Sub InsertTextAt(it As String, Optional ipos As Integer = -1)
' indsætter it på position ipos
    If ipos < 0 Then ipos = pos

    If ipos < 2 Then
        text = it & text
    ElseIf ipos = Len(text) + 1 Then
        text = text & it
    Else
        text = Left(text, ipos - 1) & it & right(text, Len(text) - ipos + 1)
    End If
End Sub
Sub RemoveChar(Optional ind As Integer = -1)
' fjerner char ved posistion ind eller pos hvis ingen angivet

    If ind < 0 Then ind = pos

    If ind < 2 Then
        text = right(text, Len(text) - 1)
    ElseIf ind = Len(text) Then
        text = Left(text, Len(text) - 1)
    Else
        text = Left(text, ind - 1) & right(text, Len(text) - ind)
    End If
End Sub

Function ValidateBrackets() As Boolean
' angiver om paranteser er angivet korrekt
' dog kun simpel metode der ikke fanger alt

ValidateBrackets = CountText("(") = CountText(")")

End Function

Function CountText(s As String) As Integer
' tæller hvor mange gange s optræder
Dim p As Integer
Dim i As Integer
p = 1
Do While p > 0 And p <= Len(text)
    p = InStr(p, text, s)
    i = i + 1
    If p > 0 Then p = p + 1
Loop
i = i - 1
CountText = i
End Function
Sub ReplaceStringAt(s As String, Optional p As Integer = -1)
' overskriver text med s startende ved pos p

If p < 0 Then p = pos

If p < 2 Then
    text = s & right(text, Len(text) - Len(s))
ElseIf p > Len(text) Then
    text = text & s
ElseIf p > Len(text) - Len(s) Then
    text = Left(text, p - 1) & s
Else
    text = Left(text, p - 1) & s & right(text, Len(text) - Len(s) - p + 1)
End If

End Sub

Sub ConvertDecSeparator()
' konverterer , til . hvor , omgivet af tal
    Dim p As Integer
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
' fjerner mellemrum efter f.eks. f_a (x)
' man så ellers kan erstatte mellemrum med *
Dim p As Integer
On Error GoTo Slut

Do
    pos = InStr(pos + 1, text, tegn)
    If pos = 1 Or pos = Len(text) Then
    ElseIf pos > 0 Then
'        pos = pos + 1
        FindNextBreak (False)
        If ChrByIndex(pos + 1) = " " And ChrByIndex(pos + 2) = "(" Then
             text = Left(text, pos) & right(text, Len(text) - pos - 1)
        End If
    End If
Loop While pos > 0

Slut:
End Sub
Sub AddSpaceAfterSubSuperscript(tegn As String)
' tilføjer mellemrum efter f.eks. f_a (x)
' man så ellers kan erstatte mellemrum med *
Dim p As Integer
'On Error GoTo slut
p = pos
pos = 1
Do
    pos = InStr(pos, text, tegn)
    If pos = 1 Or pos = Len(text) Then
    ElseIf pos > 0 Then
        pos = pos + 1
        If ChrByIndex(pos) = "(" Then
            GetNextBracketContent
            pos = pos - 1
        Else
            FindHierakiBreak 5
        End If
        If ChrByIndex(pos + 1) = ")" Then pos = pos + 1
        
'        FindNextBreak (True)
        If ChrByIndex(pos + 1) = "(" Then
             text = Left(text, pos) & " " & right(text, Len(text) - pos)
        End If
    End If
Loop While pos > 0
pos = p
Slut:
End Sub

Sub InsertFunctionBrackets()
' indsætter parenteser efter f.eks. log 2 -> log(2)
'  specielt funktionssymbol 8289 fjernes ikke endnu da det skal bruges af andre funktioner
Dim p As Integer
Dim p2 As Integer
Dim c As String
Dim sb As String
Dim eb As String

p = 1
Do
    p = InStr(p, text, VBA.ChrW(8289))
    If p > 0 Then
        p = p + 1
        c = ChrByIndex(p)
        If c = "(" Then
        ElseIf c = VBA.ChrW(12310) Then ' hvis usynlig parentes
            If ChrByIndex(p + 1) <> "(" Then ' usynlige parenteser kan være uønskede. Det antages hvis parentes lige efter
                pos = p
                sb = StartBracket
                eb = EndBracket
                StartBracket = VBA.ChrW(12310)
                EndBracket = VBA.ChrW(12311)
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

Function ContainsVar(var As String) As Boolean
' afgør om variabel var er i text
Dim ivar As String

ContainsVar = False
pos = 0

Do
    ivar = GetNextVar(pos + 1)
    If ivar = var Then
        ContainsVar = True
    End If
Loop While ivar <> "" And pos < Len(text)

End Function

Sub ConvertTrigPow(trig As String)
' omdanner f.eks. sin^2(x) -> sin(x)^2
Dim ea As ExpressionAnalyser
Dim pow As String
Dim pos2 As Integer
pos = 1
Do
    pos = InStr(pos, text, trig & "^")
    If pos > 0 Then
        pos = pos + Len(trig) + 1
        pos2 = InStr(pos, text, VBA.ChrW(8289))
        If pos2 = 0 Then
            Set ea = New ExpressionAnalyser
            ea.SetNormalBrackets
            ea.text = text
            ea.pos = pos
            pos2 = ea.FindEndOfBracket
        End If
        pow = Mid(text, pos, pos2 - pos)
        text = Left(text, pos - 2) & right(text, Len(text) - pos2 + 1)
        InsertTextAt "^" & pow, FindEndOfBracket
        
    End If
Loop While pos > 0

End Sub

Function FindNextGreaterOrLower(Optional StartPos As Integer = 0)
' finder næste < eller >
Dim p1 As Integer, p2 As Integer
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
Dim p1 As Integer, p2 As Integer, p3 As Integer

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
' checker om variabel er enhed. Kun rene enheder
' bruges til Latex så enheder kan skrives i ikke-kursiv
' ikke færdig
Dim p As String

If E = "kWh" Or E = "min" Then ' godkend dem der kan forveksles med præfiks
    IsUnit = True
    GoTo Slut
End If

If Len(E) > 1 Then ' fjern evt. præfiks
p = Left(E, 1)
If p = "k" Or p = "M" Or p = "G" Or p = "T" Or p = "P" Or p = "c" Or p = "m" Or p = VBA.ChrW(956) Or p = "n" Or p = "p" Or p = "a" Or p = "" Then
    E = right(E, Len(E) - 1)
End If
End If

If E = "g" Or E = "m" Or E = "s" Or E = "K" Or E = "mol" Or E = "cd" Or E = "A" Or E = "V" Or E = "W" Or E = "J" Or E = "cal" Or E = "N" Or E = "T" Or E = "H" Or E = "timer" Or E = "Ohm" Then
    IsUnit = True
Else
    IsUnit = False
End If

Slut:
End Function

Function SubString(StartPos As Integer, endpos As Integer)

If StartPos < 1 Then StartPos = 1

If endpos > Length Then
    endpos = Length
End If

If endpos < StartPos Then
    SubString = ""
Else
    SubString = Mid(text, StartPos, endpos - StartPos + 1)
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
        StartBracket = VBA.ChrW(12310)
        EndBracket = VBA.ChrW(12311)
    End If

End Sub

Sub SetBracketsS(b As String)
    If b = "(" Or b = ")" Then
        SetNormalBrackets
    ElseIf b = "[" Or b = "]" Then
        SetSquareBrackets
    ElseIf b = "{" Or b = "}" Then
        SetTuborgBrackets
    ElseIf b = VBA.ChrW(12310) Or b = VBA.ChrW(12311) Then
        StartBracket = VBA.ChrW(12310)
        EndBracket = VBA.ChrW(12311)
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

Function GetFirstMatrix() As CMatrix
' finder næste matrix fra pos
Dim p As Integer
Dim s As String
Dim arr As Variant
Dim arr2 As Variant
Dim i As Integer, j As Integer
Dim M As New CMatrix
    
    SetNormalBrackets
    text = Replace(text, VBA.ChrW(9632), VBA.ChrW(9608))
    p = InStr(pos, text, VBA.ChrW(9608)) ' vektor
    s = GetNextBracketContent(p)
    If Len(s) > 0 Then
    arr = Split(s, "@")
    For i = 0 To UBound(arr)
        arr2 = Split(arr(i), "&")
        For j = 0 To UBound(arr2)
            M.Element(i, j) = arr2(j)
        Next
    Next
    End If
    Set GetFirstMatrix = M
End Function
Function GetAllMatrices() As Collection
Dim M As CMatrix
    pos = 1
    
    M = GetFirstMatrix()
    Do While Not M Is Nothing
        GetAllMatrices.Add M
        M = GetFirstMatrix()
    Loop

End Function

