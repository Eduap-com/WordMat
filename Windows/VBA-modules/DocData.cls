VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Tabel() As Double
Private tabels() As String
Private mboxes() As String
Private antalmboxes As Integer
Public nrows As Integer
Public ncolumns As Integer
Public Function ReadSelection() As Boolean
    Dim cellval As String
    Dim r As Integer
    Dim i As Integer, j As Integer, k As Integer, l As Integer, c As Integer
    Dim startrow As Integer
    Dim startcol As Integer
    Dim D As Double
    On Error GoTo fejl

    r = 1
    c = 1
    If Selection.OMaths.Count > 0 Then
        ReadMathBoxes
    End If
    If Selection.Tables.Count = 0 Then
        GetSetData
    End If
    If Selection.Tables.Count > 0 Then
        On Error GoTo hop
        r = Selection.Tables(1).Rows.Count
        c = Selection.Tables(1).Columns.Count
hop:
        On Error GoTo fejl
        If r < 2 Or c < 2 Then
            If Selection.Tables.Count < 1 Then GoTo fejl
            Selection.Tables(1).Select
            r = Selection.Rows.Count
            c = Selection.Columns.Count
            If r < 2 Or c < 2 Then
                GoTo fejl
            End If
        End If
        On Error Resume Next
        startrow = 0
        startcol = Selection.Cells(1).Column.Index
        If c > 1 Then
            cellval = Selection.Rows(1).Cells(startcol + 1).Range.Text
            cellval = Left(cellval, Len(cellval) - 1)
            cellval = Replace(cellval, "*10^", "E")
            D = CDbl(cellval)
            If Err.Number > 0 Then
                startrow = 1
                r = r - 1
                Err.Clear
            End If
        End If
        If r > 1 Then
            cellval = Selection.Rows(2).Cells(startcol).Range.Text
            cellval = Left(cellval, Len(cellval) - 1)
            cellval = Replace(cellval, "*10^", "E")
            D = CDbl(cellval)
            If Err.Number > 0 Then
                startcol = startcol + 1
                c = c - 1
                Err.Clear
            End If
        End If
    
        ReDim Tabel(1 To r, 1 To c) As Double

        For j = 1 To r
            i = i + 1
            For k = 1 To c
                l = l + 1
                cellval = Replace(Selection.Rows(i + startrow).Cells(l + startcol - 1).Range.Text, ",", ".")
                cellval = Left(cellval, Len(cellval) - 1)
                cellval = Replace(cellval, "*10^", "E")
                Tabel(i, l) = CDbl(cellval)
                Tabel(i, l) = val(cellval)
                If Err.Number > 0 Then
                    '           l = l - 1
                    Err.Clear
                End If
            Next
            l = 0
        Next
        ReDim Preserve Tabel(1 To i, 1 To c) As Double
        nrows = i
        ncolumns = c
        Selection.Tables(1).Select
        Selection.Collapse (wdCollapseEnd)
        Selection.TypeParagraph
        ReadSelection = True
    End If
    GoTo slut
fejl:
    ReadSelection = False
slut:

End Function
Public Function ReadSelectionS() As Boolean
' reads as strings
Dim cellval As String
Dim r As Integer
Dim i As Integer, j As Integer, k As Integer, l As Integer, c As Integer
Dim startrow As Integer
Dim startcol As Integer

r = 1
c = 1
If Selection.Tables.Count = 0 And Selection.OMaths.Count > 0 Then
    GetSetData
Else
On Error GoTo hop
r = Selection.Rows.Count
c = Selection.Columns.Count
hop:
On Error GoTo fejl
If r < 1 Or c < 1 Or (r = 1 And c = 1) Then
    If Selection.Tables.Count < 1 Then GoTo fejl
    Selection.Tables(1).Select
    r = Selection.Rows.Count
    c = Selection.Columns.Count
    If r < 1 Or c < 1 Then
        GoTo fejl
    End If
End If
    On Error Resume Next
    startrow = 0
    startcol = Selection.Cells(1).Column.Index
    
    ReDim tabels(1 To r, 1 To c) As String

    For j = 1 To r
        i = i + 1
        For k = 1 To c
        l = l + 1
        cellval = Replace(Selection.Rows(i + startrow).Cells(l + startcol - 1).Range.Text, ",", ".")
        cellval = Left(cellval, Len(cellval) - 1)
        cellval = Replace(cellval, vbCrLf, "")
        cellval = Replace(cellval, vbCr, "")
        cellval = Replace(cellval, vbLf, "")
        tabels(i, l) = cellval
        Next
        l = 0
    Next
    ReDim Preserve tabels(1 To i, 1 To c) As String
    nrows = i
    ncolumns = c
    Selection.Tables(1).Select
    Selection.Collapse (wdCollapseEnd)
    Selection.TypeParagraph
    ReadSelectionS = True ' ingen fejl
End If
GoTo slut
fejl:
    ReadSelectionS = False
slut:

End Function

Public Function TabelCelle(r, c) As Double
On Error GoTo fejl
If c > ncolumns Or r > nrows Then GoTo slut

TabelCelle = Tabel(r, c)

GoTo slut
fejl:
slut:

End Function
Public Function TabelsCelle(r, c) As String
On Error GoTo fejl
If c > ncolumns Or r > nrows Then GoTo slut

TabelsCelle = tabels(r, c)

GoTo slut
fejl:
slut:

End Function

Public Function GetSetForm(Optional separator As Integer = 0) As String
' returns data as set {(1,2),(3,4)}
' 0 period decseparator
' 1 comma decseparator

Dim t As String
Dim i As Integer, j As Integer
t = "{ "
For i = 1 To nrows
    t = t & "(" & ConvertSeparator(Tabel(i, 1), separator)
    For j = 2 To ncolumns
        t = t & " , " & ConvertSeparator(Tabel(i, j), separator)
    Next
    t = t & "),"
Next

If Right(t, 1) = "," Then
    t = Left(t, Len(t) - 1)
End If
t = t & " }"

GetSetForm = t

End Function
Public Function GetSetFormS(Optional separator As Integer = 0) As String
' returns data as set {(1,2),(3,4)}
' 0 period decseparator
' 1 comma decseparator

Dim t As String
Dim i As Integer, j As Integer
t = "{ "
For i = 1 To nrows
    t = t & "(" & ConvertSeparator(tabels(i, 1), separator)
    For j = 2 To ncolumns
        t = t & " , " & ConvertSeparator(tabels(i, j), separator)
    Next
    t = t & "),"
Next

If Right(t, 1) = "," Then
    t = Left(t, Len(t) - 1)
End If
t = t & " }"

GetSetFormS = t

End Function
Public Function GetListFormS(Optional separator As Integer = 0) As String
' returns data as set {(1,2),(3,4)}
' 0 period decseparator
' 1 comma decseparator

Dim t As String
Dim i As Integer, j As Integer
t = "[ "
For i = 1 To nrows
    t = t & "[" & ConvertSeparator(tabels(i, 1), separator)
    For j = 2 To ncolumns
        t = t & " , " & ConvertSeparator(tabels(i, j), separator)
    Next
    t = t & "],"
Next

If Right(t, 1) = "," Then
    t = Left(t, Len(t) - 1)
End If
t = t & " ]"

GetListFormS = t

End Function
Function ConvertSeparator(ByVal t As String, n As Integer)
If n = 0 Then
    ConvertSeparator = ConvertToPunktum(t)
Else
    ConvertSeparator = ConvertToKomma(t)
End If
End Function

Function ConvertToPunktum(ByVal t As String) As String
    ConvertToPunktum = Replace(t, ",", ".")
End Function
Function ConvertToKomma(ByVal t As String) As String
    ConvertToKomma = Replace(t, ".", ",")
End Function

Public Sub GetSetData()
' reads set {(2,1),(3,4)} or list from mathbox into table
Dim p As String
Dim Text As String
Dim n As Integer
Dim ea As New ExpressionAnalyser
Dim r() As String
Dim i As Integer
On Error GoTo slut

ReDim r(1 To 1000) As String
If Selection.OMaths.Count > 0 Then
    Text = mboxes(0)
'    Selection.OMaths.Linearize
'    text = Selection.OMaths(1).Range.text
Else
    Text = Selection.Range.Text
End If
ea.Text = Replace(Text, VBA.ChrW(8722), "-") ' required

If InStr(Text, "[") > 0 Then
    ea.SetSquareBrackets
    ea.Text = ea.GetNextBracketContent()
ElseIf InStr(Text, "{") > 0 Then
    ea.SetTuborgBrackets
    ea.Text = ea.GetNextBracketContent()
Else
    ea.Text = Text
End If
ea.SetNormalBrackets
ea.Pos = 1
nrows = 0
    ea.Text = Replace(ea.Text, ";", " , ")
    ea.ConvertDecSeparator


Do
'    If ea.WhatsNextBracket = "" Then
'        p = ea.text
'    Else
'        ea.SetToNextBracket
'        p = ea.GetNextBracketContent()
'    End If
    p = Trim(ea.GetNextListItem(ea.Pos, ","))
    If Left(p, 1) = "[" And Right(p, 1) = "]" Then
        If Len(p) > 2 Then p = Mid(p, 2, Len(p) - 2)
    End If
'    ea2.text = Replace(p, ";", " , ")
'    ea2.ConvertDecSeparator
    
    If Len(p) > 0 Then
        nrows = nrows + 1
        r(nrows) = p
    End If
Loop While Not (p = "") And nrows < 1000


'nrows = nrows - 1
If nrows > 0 Then
    ReDim tabels(1 To nrows, 1 To 100)
Else
    ReDim tabels(1 To 1, 1 To 100)
End If

For i = 1 To nrows
    ea.Text = r(i)
    ea.Pos = 1
    n = 0
    Do
        p = ea.GetNextListItem(0, ",")
        n = n + 1
        If p <> "" And n > ncolumns Then
            ncolumns = n
        End If
        If Not (MaximaSeparator) Then p = Replace(p, ".", ",")
        tabels(i, n) = omax.ConvertToAscii(p)
    Loop While ea.Pos < Len(ea.Text) And n < 1000
Next
    
If nrows = 1 And ncolumns = 1 Then ' if only one field then it is probably just a math field and not a list or table
    ReDim Preserve tabels(1 To 1, 1 To 1) As String
    nrows = 0
    ncolumns = 0
ElseIf nrows > 0 Then
    ReDim Preserve tabels(1 To nrows, 1 To ncolumns) As String
End If

slut:
End Sub

Sub ReadMathBoxes()
' inserts the selected equation boxes as mboxes array
Dim antal As Integer
Dim j As Integer
On Error GoTo slut

' if ball gives it buildup problems
For j = 1 To Selection.OMaths.Count
    Selection.OMaths(j).Range.Font.Bold = False
Next

'sstart = Selection.start
'sslut = Selection.End
antal = Selection.OMaths.Count
Dim sr As Range
'    With tempDoc
'        .Content.Font.Name = "Tahoma"
'        .SaveAs FileName:="Sample.doc"
'    End With
    
If antal = 0 Then
'    kommando = Selection.text
    Selection.Collapse (wdCollapseEnd)
Else
    ReDim mboxes(0 To 100)
    Set sr = Selection.Range
    
    If Len(Selection.Text) < 3 Then
        Selection.OMaths(1).Range.Select
        Set sr = Selection.Range
    End If

    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord
    With sr
        .OMaths.BuildUp
        .OMaths.Linearize
    End With

'    Set tempDoc = Documents.Add(, , , False)
'    Set r = tempDoc.Range
'    r.Paste
'    With r
'    .OMaths.BuildUp
'    .OMaths.Linearize
'    End With
    
    Dim mathcol As OMaths
    Set mathcol = sr.OMaths
    For j = 0 To antal - 1
        mathcol(j + 1).ConvertToNormalText
        mboxes(j) = Replace(mathcol(j + 1).Range.Text, ChrW(11), "") ' remove shift-enter
    Next
    ReDim Preserve mboxes(0 To antal - 1)
    antalmboxes = antal
    sr.Select
    Oundo.EndCustomRecord
    ActiveDocument.Undo
End If

slut:
    
    sr.Select
' Selection.start = sstart
' Selection.End = sslut

End Sub


Public Property Get MathBoxes(Index As Integer) As String
    On Error Resume Next
    If Index <= UBound(mboxes) Then
        MathBoxes = mboxes(Index)
    Else
        MathBoxes = ""
    End If
End Property

Public Property Get AntalMathBoxes() As Variant
    On Error Resume Next
    AntalMathBoxes = antalmboxes
End Property

