VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRegression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private xval() As Double
Private yval() As Double
Private Coeff() As Double ' coefficients for polregr Coeff(0)=a_0 ...
Public A As Double
Public b As Double
Public c As Double
Public r As Double
Public Ligning As String
Private RegressionText As String
Private regrtype As Integer

Public Property Get XValues() As Variant
    XValues = xval
End Property

Public Property Let XValues(ByVal Value As Variant)
    xval = Value
End Property
Public Property Get YValues() As Variant
    YValues = yval
End Property

Public Property Let YValues(ByVal Value As Variant)
    yval = Value
End Property
Public Property Get Datatext() As Variant

End Property

Public Property Let Datatext(ByVal Text As Variant)
    Dim i As Integer
    Dim j As Integer
    Dim lsep As String
    Dim Arr As Variant
    Dim Arr2 As Variant
    Dim n As Integer
    Arr = Split(Text, VbCrLfMac)
    n = UBound(Arr)
    ReDim xval(1 To n + 1) As Double
    ReDim yval(1 To n + 1) As Double
    For j = 0 To UBound(Arr)
        If InStr(Arr(j), ";") Then
            lsep = ";"
        Else
            lsep = ","
        End If
        Arr2 = Split(Arr(j), lsep)
        If UBound(Arr2) > 0 Then
        i = i + 1
        xval(i) = Replace(Arr2(0), ".", ",")
        yval(i) = Replace(Arr2(1), ".", ",")
        End If
    Next
    ReDim Preserve xval(1 To i) As Double
    ReDim Preserve yval(1 To i) As Double
    
End Property

Public Sub GetTableData()
Dim rws As Variant
Dim cellval As String
Dim tal As Double
Dim i As Integer
Dim j As Integer
Dim n As Integer
Dim oFS As Integer

On Error GoTo Fejl

If Selection.Tables.Count = 0 Then Exit Sub

If Selection.Range.OMaths.Count > 0 Or Selection.Tables(1).Range.OMaths.Count > 1 Then PrepareMaxima

If Selection.Range.start < Selection.Tables(1).Range.start Then
   Selection.start = Selection.Tables(1).Range.start
End If

If Selection.Range.End > Selection.Tables(1).Range.End Then
    Selection.End = Selection.Tables(1).Range.End
End If

If Selection.Rows.Count >= Selection.Columns.Count Then ' fejler hvis andet end tabel også er valgt
    Set rws = Selection.Rows
    oFS = Selection.Cells(1).Column.Index - 1 ' kolonne hvor selection starter
Else
    Set rws = Selection.Columns
    oFS = Selection.Cells(1).Row.Index - 1 ' række hvor selection starter
End If
n = rws.Count

If n < 2 Then ' hvis intet valgt på forhånd
    If Selection.Tables(1).Rows.Count >= Selection.Tables(1).Columns.Count Then
        Set rws = Selection.Tables(1).Rows
        oFS = 0
    Else
        Set rws = Selection.Tables(1).Columns
        oFS = 0
    End If
    n = rws.Count
    On Error Resume Next
    For i = 1 To n ' bestem hvor der ligger tal
        For j = 1 To rws(1).Cells.Count
        cellval = rws(i).Cells(j).Range.Text
        cellval = Left(cellval, Len(cellval) - 1)
        cellval = Replace(cellval, "*10^", "E")
        tal = CDbl(Replace(Replace(cellval, ".", ""), ",", ""))
'        tal = val(Replace(cellval, ",", "."))
        If Err.Number = 0 Then
           oFS = j - 1
           Exit For
        Else
           Err.Clear
        End If
        Next
    Next
End If

If n < 2 Then
    GoTo Fejl
End If

    ReDim xval(1 To n) As Double
    ReDim yval(1 To n) As Double
'    ReDim Preserve x(1 To n) As Double
'    ReDim Preserve y(1 To n) As Double
    On Error Resume Next
    i = 0
    For j = 1 To n
        i = i + 1
        If rws(j).Cells(1 + oFS).Range.OMaths.Count > 0 Then
'            PrepareMaxima
            rws(j).Cells(1 + oFS).Range.Select
            omax.ReadSelection
            omax.beregn
            cellval = Replace(omax.MaximaOutput, "=", "")
        Else
            cellval = Replace(rws(j).Cells(1 + oFS).Range.Text, ",", ".")
            cellval = Left(cellval, Len(cellval) - 1)
        End If
'        cellval = tabel.Cell(j, 1).Range.text
        cellval = Replace(cellval, "*10^", "E")
        If Not IsNumeric(cellval) Then GoTo NotNumber
        xval(i) = CDbl(Replace(Replace(cellval, ".", ""), ",", "")) ' skal med for at melde fejl , men alene duer ikke da den er afh af sprog indstill. De to replace fjerner . og , da disse ikke genkendes af alle sprog fx norsk bruger ikke . så alle tal med . ville blive ignoreret
        xval(i) = val(Replace(cellval, ",", ".")) 'melder ikke fejl ved bogstaver
        
        If rws(j).Cells(2 + oFS).Range.OMaths.Count > 0 Then
'            PrepareMaxima
            rws(j).Cells(2 + oFS).Range.Select
            omax.ReadSelection
            omax.beregn
            cellval = Replace(omax.MaximaOutput, "=", "")
        Else
            cellval = Replace(rws(j).Cells(2 + oFS).Range.Text, ",", ".")
            cellval = Left(cellval, Len(cellval) - 1)
        End If

'        cellval = tabel.Cell(j, 2).Range.text
        cellval = Replace(cellval, "*10^", "E")
        If Not IsNumeric(cellval) Then GoTo NotNumber
        yval(i) = CDbl(Replace(Replace(cellval, ".", ""), ",", ""))
        yval(i) = val(Replace(cellval, ",", "."))
        If Err.Number > 0 Then ' hvis enten x -eller y-værdien ikke er et tal (fejl fra cdbl) så skal linjen ignoreres
NotNumber:
           i = i - 1
            Err.Clear
        End If
    Next
    ReDim Preserve xval(1 To i) As Double
    ReDim Preserve yval(1 To i) As Double
    Selection.Tables(1).Select
    Selection.Collapse (wdCollapseEnd)
    Selection.TypeParagraph
GoTo slut
Fejl:
    MsgBox Sprog.A(27), vbOKOnly, Sprog.Error
slut:

End Sub
Public Sub GetTableDataOld()
Dim rws As Variant
Dim cellval As String
Dim i, j, n, oFS As Integer
'On Error GoTo fejl
'    If Selection.Tables.Count = 0 Then
'        Call MsgBox("Du skal lave en tabel med to kolonner. x-værdier i venstre kolonne og y-værdier i højrekolonne")
'        Exit Sub
'    End If
If Selection.Tables.Count = 0 Then Exit Sub

tableonly:
On Error GoTo heletabel

If Selection.Rows.Count >= Selection.Columns.Count Then ' fejler hvis andet end tabel også er valgt
    Set rws = Selection.Rows
    oFS = Selection.Cells(1).Column.Index - 1 ' kolonne hvor selection starter
Else
    Set rws = Selection.Columns
    oFS = Selection.Cells(1).Row.Index - 1 ' række hvor selection starter
End If
n = rws.Count
GoTo Prepare
heletabel:

'Set r = Selection.Range
If Selection.Range.start < Selection.Tables(1).Range.start Then
   Selection.start = Selection.Tables(1).Range.start
End If

If Selection.Range.End > Selection.Tables(1).Range.End Then
    Selection.End = Selection.Tables(1).Range.End
End If
GoTo tableonly

Prepare:
If n < 2 Then
    If Selection.Tables(1).Rows.Count >= Selection.Tables(1).Columns.Count Then
        Set rws = Selection.Tables(1).Rows
        oFS = Selection.Cells(1).Column.Index - 1 ' kolonne hvor selection starter
    Else
        Set rws = Selection.Tables(1).Columns
        oFS = Selection.Cells(1).Row.Index - 1 ' række hvor selection starter
    End If
    n = rws.Count
End If
If n < 2 Then
    GoTo Fejl
End If

    ReDim xval(1 To n) As Double
    ReDim yval(1 To n) As Double
'    ReDim Preserve x(1 To n) As Double
'    ReDim Preserve y(1 To n) As Double
    On Error Resume Next
    For j = 1 To n
        i = i + 1
        
        cellval = Replace(rws(j).Cells(1 + oFS).Range.Text, ",", ".")
'        cellval = tabel.Cell(j, 1).Range.text
        cellval = Left(cellval, Len(cellval) - 1)
        xval(i) = val(cellval)
        cellval = Replace(rws(j).Cells(2 + oFS).Range.Text, ",", ".")
'        cellval = tabel.Cell(j, 2).Range.text
        cellval = Left(cellval, Len(cellval) - 1)
        yval(i) = val(cellval)
        If Err.Number > 0 Then
           i = i - 1
            Err.Clear
        End If
    Next
    ReDim Preserve xval(1 To i) As Double
    ReDim Preserve yval(1 To i) As Double
    Selection.Tables(1).Select
    Selection.Collapse (wdCollapseEnd)
    Selection.TypeParagraph
GoTo slut
Fejl:
    MsgBox Sprog.A(27), vbOKOnly, Sprog.Error
slut:

End Sub
Public Sub GetSetData()
' læser mængde {(2,1),(3,4)} fra mathbox ind i xval og yval
Dim p As String
Dim Text As String
Dim n As Integer
Dim ea As New ExpressionAnalyser
Dim Arr As Variant

    ReDim xval(1 To 1000) As Double
    ReDim yval(1 To 1000) As Double
ea.StartBracket = "{"
ea.EndBracket = "}"
Selection.OMaths.Linearize
Text = Selection.OMaths(1).Range.Text
ea.Text = Replace(Text, VBA.ChrW(8722), "-") ' nødvendigt
ea.Text = ea.GetNextBracketContent()
ea.SetNormalBrackets
ea.Pos = 1
Do
p = ea.GetNextBracketContent()
If InStr(p, ";") Then
    Arr = Split(p, ";")
Else
    Arr = Split(p, ",")
End If
If UBound(Arr) > 0 Then
    n = n + 1
'    xval(n) = CDbl(arr(0))
'    yval(n) = CDbl(arr(1))
    xval(n) = val(Replace(Arr(0), ",", "."))
    yval(n) = val(Replace(Arr(1), ",", "."))
End If
Loop While Not (p = "") And n < 1000

    ReDim Preserve xval(1 To n) As Double
    ReDim Preserve yval(1 To n) As Double

End Sub

Private Function ComputeRegr(X() As Double, Y() As Double) As Boolean
'Udfører lineær regression på to arrays og returnerer a,b og R^2

Dim sumx As Double, sumy As Double, sumxy As Double, sumxx As Double, sumyy As Double
Dim n As Integer
Dim i As Integer
On Error GoTo Fejl

n = UBound(X) - LBound(X) + 1
               
For i = LBound(X) To UBound(X)
    sumx = sumx + X(i)
    sumy = sumy + Y(i)
    sumxy = sumxy + X(i) * Y(i)
    sumxx = sumxx + X(i) ^ 2
    sumyy = sumyy + Y(i) ^ 2
Next

A = (n * sumxy - sumx * sumy) / (n * sumxx - sumx ^ 2)
A = betcif(A, 16)   ' sikrer at man ikke får små værdier for b der egentlig er 0
b = (sumy - A * sumx) / n
If n * sumxx - sumx ^ 2 = 0 Or (n * sumyy - sumy ^ 2) = 0 Then
    r = 1
Else
    r = (n * sumxy - sumx * sumy) ^ 2 / (n * sumxx - sumx ^ 2) / (n * sumyy - sumy ^ 2)
    r = betcif(r, MaximaCifre)
End If
If r > 1 Then r = 1

ComputeRegr = True
GoTo slut
Fejl:
    ComputeRegr = False
    MsgBox Sprog.A(28), vbOKOnly, Sprog.Error
slut:
End Function

Function Logv(val() As Double) As Variant
'tager log af alle elementer i array
    Dim n As Integer, j As Integer, i As Integer
    
    ReDim arrlog(LBound(val) To UBound(val)) As Double
    
    n = UBound(val) - LBound(val) + 1
    For i = 1 To n
        If val(i) > 0 Then
            j = j + 1
            arrlog(j) = Log(val(i))
        Else
            MsgBox Sprog.A(672), vbOKOnly, Sprog.Error
            GoTo slut
        End If
    Next
slut:
    ReDim Preserve arrlog(LBound(val) To j) As Double

    Logv = arrlog
End Function
Sub ComputeLinRegr()
    
    If ComputeRegr(xval, yval) Then
    
    Ligning = "y="
    If A = 1 Then
        Ligning = Ligning & "x"
    Else
        Ligning = Ligning & ConvertNumberToString(betcif(A, MaximaCifre)) & "x"
    End If
    If b <> 0 Then
        If Left(b, 1) <> "-" Then
            Ligning = Ligning & "+"
        End If
        Ligning = Ligning & ConvertNumberToString(betcif(b, MaximaCifre))
    End If
    RegressionText = Sprog.A(29)
'    regrtype = 1
    Else
        Ligning = ""
        A = 0
        b = 0
        c = 0
        r = 0
    End If

End Sub
Sub ComputeExpRegr()
    Dim astr As String
    ReDim logy(LBound(yval) To UBound(yval)) As Double
    
    logy = Logv(yval)
    If ComputeRegr(xval, logy) Then
        
    Ligning = "y="
    If b <> 1 Then
        Ligning = Ligning & ConvertNumberToString(betcif(Exp(b), MaximaCifre)) & VBA.ChrW(183)
    End If
    If A <> 1 Then
        astr = ConvertNumberToString(betcif(Exp(A), MaximaCifre))
        If InStr(astr, "10^(") Then astr = "(" & astr & ")"
'        ligning = ligning & VBA.ChrW(12310) & astr & VBA.ChrW(12311) & "^x" ' usynlige paranteser
        Ligning = Ligning & astr & "^x"
    End If

    RegressionText = Sprog.A(30)
'    regrtype = 2
    Else
        Ligning = ""
        A = 0
        b = 0
        c = 0
        r = 0
    End If

End Sub

Sub ComputePowRegr()
On Error GoTo Fejl
    ReDim logx(LBound(xval) To UBound(xval)) As Double
    ReDim logy(LBound(yval) To UBound(yval)) As Double
    
    logx = Logv(xval)
    logy = Logv(yval)
    If ComputeRegr(logx, logy) Then
    
    Ligning = "y=" & ConvertNumberToString(betcif(Exp(b), MaximaCifre))
    If A <> 0 Then
        Ligning = Ligning & VBA.ChrW(183) & "x^(" & ConvertNumberToString(betcif(A, MaximaCifre)) & ")"
    End If
    
    RegressionText = Sprog.A(31)
'    regrtype = 3
    Else
        Ligning = ""
        A = 0
        b = 0
        c = 0
        r = 0
    End If
GoTo slut
Fejl:
    MsgBox Sprog.A(32), vbOKOnly, Sprog.Error
slut:
End Sub
Sub ComputeUserRegr()
    Dim fkt As String
    Dim Arr As Variant
    Dim xlist As String, ylist As String
    Dim i As Long, guess As String, Var As String, p As Integer, eps As String
    On Error GoTo Fejl
    xlist = "["
    For i = LBound(xval) To UBound(xval)
        xlist = xlist & Replace((xval(i)), ",", ".") & ","
    Next
    xlist = Left(xlist, Len(xlist) - 1) & "]"
    
    ylist = "["
    For i = LBound(yval) To UBound(yval)
        ylist = ylist & Replace(yval(i), ",", ".") & ","
    Next
    ylist = Left(ylist, Len(ylist) - 1) & "]"
    
    
    fkt = Replace(omax.Kommando, vbCrLf, "")
    If fkt <> vbNullString Then
        If InStr(fkt, "=") Then
            Arr = Split(fkt, "=")
            If UBound(Arr) > 0 Then fkt = Arr(1)
        End If
        UserFormUserRegression.TextBox_function.Text = fkt
    End If
    
    If InStr(fkt, "x") > 0 Then
        Var = "x"
    ElseIf InStr(fkt, "t") > 0 Then
        Var = "t"
    ElseIf InStr(fkt, "a") > 0 Then
        Var = "a"
    ElseIf InStr(fkt, "b") > 0 Then
        Var = "b"
    ElseIf InStr(fkt, "c") > 0 Then
        Var = "c"
    ElseIf InStr(fkt, "y") > 0 Then
        Var = "y"
    ElseIf fkt = vbNullString Then
        Var = "x"
    Else
        Var = ""
    End If
    UserFormUserRegression.TextBox_var.Text = Var
    
    UserFormUserRegression.TextBox_eps.Text = MaximaCifre
    
    UserFormUserRegression.Show
    
    eps = UserFormUserRegression.TextBox_eps.Text
    fkt = UserFormUserRegression.TextBox_function.Text
    fkt = Replace(fkt, "*", MaximaGangeTegn)
    guess = "[" & UserFormUserRegression.TextBox_guess & "]"
    Var = UserFormUserRegression.TextBox_var.Text
    If Var <> vbNullString And fkt <> vbNullString Then
        omax.UserRegression ConvertSep(fkt), Var, xlist, ylist, CInt(eps), ConvertSep(guess)
        p = InStr(omax.KommentarOutput, "R^2=")
        If p > 0 Then
            r = ConvertStringToNumber(right(omax.KommentarOutput, Len(omax.KommentarOutput) - p - 3))
        End If
        RegressionText = Sprog.A(33) & fkt & Sprog.A(34) & ": "
        Ligning = omax.MaximaOutput
    End If
    
GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Function ConvertSep(Text As String) As String
    ' kommaer og listeseparator
    Dim ea As New ExpressionAnalyser
    If Not MaximaSeparator Then
        ea.Text = Text
        ea.ConvertDecSeparator
        Text = ea.Text
        '        text = Replace(text, ",", ".")
    End If
    ConvertSep = Replace(Text, ";", ",")

End Function
Sub ComputePolRegrKvad()
    If polregrkvad(xval, yval) Then
    A = betcif(A, MaximaCifre)
    b = betcif(b, MaximaCifre)
    c = betcif(c, MaximaCifre)
    r = betcif(r, MaximaCifre)
    Ligning = "y=" & ConvertNumberToString(A) & VBA.ChrW(183) & "x^2"
    If b > 0 Then
        Ligning = Ligning & "+" & ConvertNumberToString(b) & VBA.ChrW(183) & "x"
    ElseIf b < 0 Then
        Ligning = Ligning & ConvertNumberToString(b) & VBA.ChrW(183) & "x"
    End If
    
    If c > 0 Then
        Ligning = Ligning & "+" & ConvertNumberToString(c)
    ElseIf c < 0 Then
        Ligning = Ligning & ConvertNumberToString(c)
    End If
    
    regrtype = 4
    Else
        Ligning = ""
        A = 0
        b = 0
        c = 0
        r = 0
    End If
End Sub
Sub ComputePolRegr()
    Dim i As Integer, orden As Integer, n As Integer
    On Error Resume Next
    n = UBound(xval) - LBound(xval)
    orden = InputBox(Sprog.A(35) & n & ")" & vbCrLf & vbCrLf & Sprog.A(36), Sprog.A(37), "2")
    On Error GoTo Fejl
    If orden < 1 Or orden > n Then
        MsgBox "You have stated an invalid order. it must be between 1 and " & n, vbOKOnly, Sprog.Error
        Exit Sub
    End If
    
    If polregr(xval, yval, orden) Then
    r = betcif(r, MaximaCifre)
    If Coeff(UBound(Coeff)) = 1 Then
        Ligning = "y=x^" & UBound(Coeff)
    ElseIf Coeff(UBound(Coeff)) = -1 Then
        Ligning = "y=-x^" & UBound(Coeff)
    ElseIf Coeff(UBound(Coeff)) = 0 Then
        Ligning = "y="
    Else
        Ligning = "y=" & ConvertNumberToString(betcif(Coeff(UBound(Coeff)), MaximaCifre)) & VBA.ChrW(183) & "x^" & UBound(Coeff)
    End If
    
    For i = UBound(Coeff) - 1 To 0 Step -1
        If Coeff(i) > 0 Then
            Ligning = Ligning & "+"
        End If
        If Not Coeff(i) = 0 Then
           If i = 0 Then
               Ligning = Ligning & ConvertNumberToString(betcif(Coeff(i), MaximaCifre))
           ElseIf i = 1 Then
               If Coeff(i) = 1 Then
                   Ligning = Ligning & "x"
               ElseIf Coeff(i) = -1 Then
                   Ligning = Ligning & "-x"
               Else
                   Ligning = Ligning & ConvertNumberToString(betcif(Coeff(i), MaximaCifre)) & VBA.ChrW(183) & "x"
               End If
           Else
               If Coeff(i) = 1 Then
                   Ligning = Ligning & "x^" & i
               ElseIf Coeff(i) = -1 Then
                   Ligning = Ligning & "-x^" & i
               Else
                   Ligning = Ligning & ConvertNumberToString(betcif(Coeff(i), MaximaCifre)) & VBA.ChrW(183) & "x^" & i
               End If
           End If
        End If
    Next
    
    RegressionText = Sprog.A(38) & orden & Sprog.A(34) & ": " & vbTab
'    regrtype = 4
    Else
        Ligning = ""
        ReDim Coeff(0, 0)
        A = 0
        b = 0
        c = 0
        r = 0
    End If
GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Function polregr(X() As Double, Y() As Double, orden As Integer, Optional ByVal tving As Boolean = False) As Boolean
'Udfører polynomisk regression på to arrays og returnerer a,b,c og R^2
' tving betyder tving gennem (0,0)
'  ax^2+bx+c
Dim n As Integer, i As Integer, j As Integer, q As Integer, ymean As Double
Dim SStotal As Double, SSregr As Double
Dim m() As Double, s() As Double, t() As Double
Dim np As Integer ' number of points
On Error GoTo Fejl

If Not (UBound(X) = UBound(Y)) Then
    polregr = False
    Exit Function
End If

np = UBound(X) - LBound(X) + 1

If np <= orden Then
    polregr = False
    Exit Function
End If

ReDim m(1 To orden + 1, 1 To orden + 2)
ReDim s(2 * orden)
ReDim t(orden)

n = UBound(X) - LBound(X) + 1
               
For j = 0 To orden ' T(i) sættes op (højre kolonne af matrix)
    For i = 0 To np - 1
        t(j) = t(j) + Y(i + 1) * X(i + 1) ^ j
    Next
Next
               
For j = 0 To 2 * orden ' S(i) sættes op (elementer til det indre af matrix)
    For i = 0 To np - 1
        s(j) = s(j) + X(i + 1) ^ j
    Next
Next

               

If tving = False Then
' sæt indre af matrix op
For i = 1 To orden + 1
    For j = 1 To orden + 1
        m(i, j) = s(orden + i - j)
    Next
Next
'sæt højre kolonne i matrix op
For i = 0 To orden
    m(i + 1, orden + 2) = t(i)
Next

' Der udføres gaussisk eliminination
For i = 1 To orden + 1
' dividerer række i med værdien i celle (i,i)
  For j = orden + 2 To 1 Step -1
    If m(i, i) = 0 Then
        MsgBox "Divide by zero", vbOKOnly, Sprog.Error
        Exit Function
    Else
        m(i, j) = m(i, j) / m(i, i)
    End If
  Next
' laver nuller i kolonne i undtagen celle i,i hvor der skal stå et 1
  For q = 1 To orden + 1
    If q <> i Then
      For j = orden + 2 To 1 Step -1
        m(q, j) = m(q, j) - m(i, j) * m(q, i)
      Next
    End If
  Next
Next

ReDim Coeff(orden)
For i = 0 To orden
    Coeff(i) = m(orden + 1 - i, orden + 2)
Next
            
Else
tving:

End If
             
ymean = t(0) / n
Dim FV As Double
For i = LBound(X) To UBound(X)
  SStotal = SStotal + (Y(i) - ymean) ^ 2
  FV = -ymean
  For j = 0 To orden
    FV = FV + Coeff(j) * X(i) ^ j
  Next
  SSregr = SSregr + FV ^ 2
Next
             
r = SSregr / SStotal
             
polregr = True
GoTo slut
Fejl:
  polregr = False
'  MsgBox Err.Number
slut:
End Function
Function polregrkvad(X() As Double, Y() As Double, Optional ByVal tving As Boolean = False) As Boolean
'Udfører polynomisk regression på to arrays og returnerer a,b,c og R^2 via userdefined datatype
' tving betyder tving gennem (0,0)
'  ax^2+bx+c
Dim sumx As Double, sumy As Double, sumxy As Double, sumxx As Double
Dim n As Integer, i As Integer, j As Integer, q As Integer, ymean As Double
Dim sumx2y As Double, sumx3 As Double, sumx4 As Double, SStotal As Double, SSregr As Double
Dim m(3, 4)

On Error GoTo Fejl

n = UBound(X) - LBound(X) + 1
               
For i = LBound(X) To UBound(X)
    sumx = sumx + X(i)
    sumy = sumy + Y(i)
    sumxy = sumxy + X(i) * Y(i)
    sumx2y = sumx2y + X(i) ^ 2 * Y(i)
    sumxx = sumxx + X(i) ^ 2
    sumx3 = sumx3 + X(i) ^ 3
    sumx4 = sumx4 + X(i) ^ 4
Next

If tving = False Then

m(1, 1) = sumxx
m(1, 2) = sumx
m(1, 3) = n
m(1, 4) = sumy
m(2, 1) = sumx3
m(2, 2) = sumxx
m(2, 3) = sumx
m(2, 4) = sumxy
m(3, 1) = sumx4
m(3, 2) = sumx3
m(3, 3) = sumxx
m(3, 4) = sumx2y

' Der udføres gaussisk eliminination
For i = 1 To 3
' dividerer række i med værdien i celle (i,i)
  For j = 4 To 1 Step -1
    If m(i, i) = 0 Then
'        MsgBox "nul"
        GoTo tving
    Else
        m(i, j) = m(i, j) / m(i, i)
    End If
  Next
' laver nuller i kolonne i undtagen celle i,i hvor der skal stå et 1
  For q = 1 To 3
    If q <> i Then
      For j = 4 To 1 Step -1
        m(q, j) = m(q, j) - m(i, j) * m(q, i)
      Next
    End If
  Next
Next
A = m(1, 4)
b = m(2, 4)
c = m(3, 4)
            
Else
tving:
A = (sumx2y * sumxx - sumx3 * sumxy) / (sumx4 * sumxx - sumx3 ^ 2)
b = (sumxy - sumx3 * A) / sumxx
c = 0

End If
             
ymean = sumy / n

For i = LBound(X) To UBound(X)
  SStotal = SStotal + (Y(i) - ymean) ^ 2
  SSregr = SSregr + (A * X(i) ^ 2 + b * X(i) + c - ymean) ^ 2
Next
             
r = SSregr / SStotal
             
polregrkvad = True
GoTo slut
Fejl:
  polregrkvad = False
'  MsgBox Err.Number
slut:
End Function


Sub InsertEquation()
    If Ligning = "" Then Exit Sub
    Application.ScreenUpdating = False
#If Mac Then
#Else
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
#End If
    Selection.TypeText RegressionText

    Selection.TypeText Text:="R"
    Selection.Font.Superscript = wdToggle
    Selection.TypeText Text:="2"
    Selection.Font.Superscript = wdToggle
    Selection.TypeText Text:=" = " & ConvertNumberToString(r)
    Selection.TypeParagraph
    Selection.Collapse
    Selection.OMaths.Add Range:=Selection.Range
    Selection.TypeText Text:=Ligning
    Selection.OMaths.BuildUp
    Selection.TypeParagraph
'    Call Indsæt2DPlot(ligning, xval, yval)
#If Mac Then
#Else
        Oundo.EndCustomRecord
#End If

End Sub


