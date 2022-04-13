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

Public Property Let Datatext(ByVal text As Variant)
    Dim i As Integer
    Dim j As Integer
    Dim lsep As String
    Dim arr As Variant
    Dim arr2 As Variant
    Dim n As Integer
    arr = Split(text, VbCrLfMac)
    n = UBound(arr)
    ReDim xval(1 To n + 1) As Double
    ReDim yval(1 To n + 1) As Double
    For j = 0 To UBound(arr)
        If InStr(arr(j), ";") Then
            lsep = ";"
        Else
            lsep = ","
        End If
        arr2 = Split(arr(j), lsep)
        If UBound(arr2) > 0 Then
        i = i + 1
        xval(i) = Replace(arr2(0), ".", ",")
        yval(i) = Replace(arr2(1), ".", ",")
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
Dim r As Range

On Error GoTo fejl

If Selection.Tables.Count = 0 Then Exit Sub

If Selection.Range.start < Selection.Tables(1).Range.start Then
   Selection.start = Selection.Tables(1).Range.start
End If

If Selection.Range.End > Selection.Tables(1).Range.End Then
    Selection.End = Selection.Tables(1).Range.End
End If

If Selection.Rows.Count >= Selection.Columns.Count Then ' fejler hvis andet end tabel ogs*aa* er valgt
    Set rws = Selection.Rows
    oFS = Selection.Cells(1).Column.Index - 1 ' kolonne hvor selection starter
Else
    Set rws = Selection.Columns
    oFS = Selection.Cells(1).Row.Index - 1 ' r*ae*kke hvor selection starter
End If
n = rws.Count

If n < 2 Then ' hvis intet valgt p*aa* forh*aa*nd
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
        cellval = rws(i).Cells(j).Range.text
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
    GoTo fejl
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
            PrepareMaxima
            rws(j).Cells(1 + oFS).Range.Select
            omax.ReadSelection
            omax.beregn
            cellval = Replace(omax.MaximaOutput, "=", "")
        Else
            cellval = Replace(rws(j).Cells(1 + oFS).Range.text, ",", ".")
            cellval = Left(cellval, Len(cellval) - 1)
        End If
'        cellval = tabel.Cell(j, 1).Range.text
        cellval = Replace(cellval, "*10^", "E")
        xval(i) = CDbl(Replace(Replace(cellval, ".", ""), ",", "")) ' skal med for at melde fejl , men alene duer ikke da den er afh af sprog indstill. De to replace fjerner . og , da disse ikke genkendes af alle sprog fx norsk bruger ikke . s*aa* alle tal med . ville blive ignoreret
        xval(i) = val(cellval) 'melder ikke fejl ved bogstaver
        
        If rws(j).Cells(2 + oFS).Range.OMaths.Count > 0 Then
            PrepareMaxima
            rws(j).Cells(2 + oFS).Range.Select
            omax.ReadSelection
            omax.beregn
            cellval = Replace(omax.MaximaOutput, "=", "")
        Else
            cellval = Replace(rws(j).Cells(2 + oFS).Range.text, ",", ".")
            cellval = Left(cellval, Len(cellval) - 1)
        End If

'        cellval = tabel.Cell(j, 2).Range.text
        cellval = Replace(cellval, "*10^", "E")
        yval(i) = CDbl(Replace(Replace(cellval, ".", ""), ",", ""))
        yval(i) = val(cellval)
        If Err.Number > 0 Then ' hvis enten x -eller y-v*ae*rdien ikke er et tal (fejl fra cdbl) s*aa* skal linjen ignoreres
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
fejl:
    MsgBox Sprog.A(27), vbOKOnly, Sprog.Error
slut:

End Sub
Public Sub GetTableDataOld()
Dim rws As Variant
Dim cellval As String
Dim i, j, n, oFS As Integer
'On Error GoTo fejl
'    If Selection.Tables.Count = 0 Then
'        Call MsgBox("Du skal lave en tabel med to kolonner. x-v*ae*rdier i venstre kolonne og y-v*ae*rdier i h*oe*jrekolonne")
'        Exit Sub
'    End If
Dim r As Range
If Selection.Tables.Count = 0 Then Exit Sub

tableonly:
On Error GoTo heletabel

If Selection.Rows.Count >= Selection.Columns.Count Then ' fejler hvis andet end tabel ogs*aa* er valgt
    Set rws = Selection.Rows
    oFS = Selection.Cells(1).Column.Index - 1 ' kolonne hvor selection starter
Else
    Set rws = Selection.Columns
    oFS = Selection.Cells(1).Row.Index - 1 ' r*ae*kke hvor selection starter
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
        oFS = Selection.Cells(1).Row.Index - 1 ' r*ae*kke hvor selection starter
    End If
    n = rws.Count
End If
If n < 2 Then
    GoTo fejl
End If

    ReDim xval(1 To n) As Double
    ReDim yval(1 To n) As Double
'    ReDim Preserve x(1 To n) As Double
'    ReDim Preserve y(1 To n) As Double
    On Error Resume Next
    For j = 1 To n
        i = i + 1
        
        cellval = Replace(rws(j).Cells(1 + oFS).Range.text, ",", ".")
'        cellval = tabel.Cell(j, 1).Range.text
        cellval = Left(cellval, Len(cellval) - 1)
        xval(i) = val(cellval)
        cellval = Replace(rws(j).Cells(2 + oFS).Range.text, ",", ".")
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
fejl:
    MsgBox Sprog.A(27), vbOKOnly, Sprog.Error
slut:

End Sub
Public Sub GetSetData()
' l*ae*ser m*ae*ngde {(2,1),(3,4)} fra mathbox ind i xval og yval
Dim p As String
Dim xs As String
Dim ys As String
Dim text As String
Dim n As Integer
Dim ea As New ExpressionAnalyser
Dim arr As Variant

    ReDim xval(1 To 1000) As Double
    ReDim yval(1 To 1000) As Double
ea.StartBracket = "{"
ea.EndBracket = "}"
Selection.OMaths.Linearize
text = Selection.OMaths(1).Range.text
ea.text = Replace(text, VBA.ChrW(8722), "-") ' n*oe*dvendigt
ea.text = ea.GetNextBracketContent()
ea.SetNormalBrackets
ea.pos = 1
Do
p = ea.GetNextBracketContent()
If InStr(p, ";") Then
    arr = Split(p, ";")
Else
    arr = Split(p, ",")
End If
If UBound(arr) > 0 Then
    n = n + 1
'    xval(n) = CDbl(arr(0))
'    yval(n) = CDbl(arr(1))
    xval(n) = val(Replace(arr(0), ",", "."))
    yval(n) = val(Replace(arr(1), ",", "."))
End If
Loop While Not (p = "") And n < 1000

    ReDim Preserve xval(1 To n) As Double
    ReDim Preserve yval(1 To n) As Double

End Sub

Private Function ComputeRegr(x() As Double, Y() As Double) As Boolean
'Udf*oe*rer line*ae*r regression p*aa* to arrays og returnerer a,b og R^2

Dim sumx As Double, sumy As Double, sumxy As Double, sumxx As Double, sumyy As Double
Dim n As Integer
Dim i As Integer
On Error GoTo fejl

n = UBound(x) - LBound(x) + 1
               
For i = LBound(x) To UBound(x)
    sumx = sumx + x(i)
    sumy = sumy + Y(i)
    sumxy = sumxy + x(i) * Y(i)
    sumxx = sumxx + x(i) ^ 2
    sumyy = sumyy + Y(i) ^ 2
Next

A = (n * sumxy - sumx * sumy) / (n * sumxx - sumx ^ 2)
A = betcif(A, 16)   ' sikrer at man ikke f*aa*r sm*aa* v*ae*rdier for b der egentlig er 0
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
fejl:
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
    Dim n As Integer
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
On Error GoTo fejl
    Dim n As Integer
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
fejl:
    MsgBox Sprog.A(32), vbOKOnly, Sprog.Error
slut:
End Sub
Sub ComputeUserRegr()
    Dim fkt As String
    Dim ea As New ExpressionAnalyser
    Dim arr As Variant
    Dim xlist As String, ylist As String
    Dim i As Long, guess As String, var As String, p As Integer, eps As String
    On Error GoTo fejl
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
            arr = Split(fkt, "=")
            If UBound(arr) > 0 Then fkt = arr(1)
        End If
        UserFormUserRegression.TextBox_function.text = fkt
    End If
    
    If InStr(fkt, "x") > 0 Then
        var = "x"
    ElseIf InStr(fkt, "t") > 0 Then
        var = "t"
    ElseIf InStr(fkt, "a") > 0 Then
        var = "a"
    ElseIf InStr(fkt, "b") > 0 Then
        var = "b"
    ElseIf InStr(fkt, "c") > 0 Then
        var = "c"
    ElseIf InStr(fkt, "y") > 0 Then
        var = "y"
    ElseIf fkt = vbNullString Then
        var = "x"
    Else
        var = ""
    End If
    UserFormUserRegression.TextBox_var.text = var
    
    UserFormUserRegression.TextBox_eps.text = MaximaCifre
    
    UserFormUserRegression.Show
    
    eps = UserFormUserRegression.TextBox_eps.text
    fkt = UserFormUserRegression.TextBox_function.text
    fkt = Replace(fkt, "*", MaximaGangeTegn)
    guess = "[" & UserFormUserRegression.TextBox_guess & "]"
    var = UserFormUserRegression.TextBox_var.text
    If var <> vbNullString And fkt <> vbNullString Then
        omax.UserRegression ConvertSep(fkt), var, xlist, ylist, CInt(eps), ConvertSep(guess)
        p = InStr(omax.KommentarOutput, "R^2=")
        If p > 0 Then
            r = ConvertStringToNumber(right(omax.KommentarOutput, Len(omax.KommentarOutput) - p - 3))
        End If
        RegressionText = Sprog.A(33) & fkt & Sprog.A(34) & ": "
        Ligning = omax.MaximaOutput
    End If
    
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Function ConvertSep(text As String) As String
    ' kommaer og listeseparator
    Dim ea As New ExpressionAnalyser
    If Not MaximaSeparator Then
        ea.text = text
        ea.ConvertDecSeparator
        text = ea.text
        '        text = Replace(text, ",", ".")
    End If
    ConvertSep = Replace(text, ";", ",")

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
    On Error GoTo fejl
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
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Function polregr(x() As Double, Y() As Double, orden As Integer, Optional ByVal tving As Boolean = False) As Boolean
'Udf*oe*rer polynomisk regression p*aa* to arrays og returnerer a,b,c og R^2
' tving betyder tving gennem (0,0)
'  ax^2+bx+c
Dim sumx As Double, sumy As Double, sumxy As Double, sumxx As Double, sumyy As Double
Dim n As Integer, i As Integer, j As Integer, q As Integer, ymean As Double
Dim sumx2y As Double, sumx3 As Double, sumx4 As Double, SStotal As Double, SSregr As Double
Dim M() As Double, s() As Double, t() As Double
Dim np As Integer ' number of points
On Error GoTo fejl

If Not (UBound(x) = UBound(Y)) Then
    polregr = False
    Exit Function
End If

np = UBound(x) - LBound(x) + 1

If np <= orden Then
    polregr = False
    Exit Function
End If

ReDim M(1 To orden + 1, 1 To orden + 2)
ReDim s(2 * orden)
ReDim t(orden)

n = UBound(x) - LBound(x) + 1
               
For j = 0 To orden ' T(i) s*ae*ttes op (h*oe*jre kolonne af matrix)
    For i = 0 To np - 1
        t(j) = t(j) + Y(i + 1) * x(i + 1) ^ j
    Next
Next
               
For j = 0 To 2 * orden ' S(i) s*ae*ttes op (elementer til det indre af matrix)
    For i = 0 To np - 1
        s(j) = s(j) + x(i + 1) ^ j
    Next
Next

               

If tving = False Then
' s*ae*t indre af matrix op
For i = 1 To orden + 1
    For j = 1 To orden + 1
        M(i, j) = s(orden + i - j)
    Next
Next
's*ae*t h*oe*jre kolonne i matrix op
For i = 0 To orden
    M(i + 1, orden + 2) = t(i)
Next

' Der udf*oe*res gaussisk eliminination
For i = 1 To orden + 1
' dividerer r*ae*kke i med v*ae*rdien i celle (i,i)
  For j = orden + 2 To 1 Step -1
    If M(i, i) = 0 Then
        MsgBox "Divide by zero", vbOKOnly, Sprog.Error
        Exit Function
    Else
        M(i, j) = M(i, j) / M(i, i)
    End If
  Next
' laver nuller i kolonne i undtagen celle i,i hvor der skal st*aa* et 1
  For q = 1 To orden + 1
    If q <> i Then
      For j = orden + 2 To 1 Step -1
        M(q, j) = M(q, j) - M(i, j) * M(q, i)
      Next
    End If
  Next
Next

ReDim Coeff(orden)
For i = 0 To orden
    Coeff(i) = M(orden + 1 - i, orden + 2)
Next
            
Else
tving:

End If
             
ymean = t(0) / n
Dim FV As Double
For i = LBound(x) To UBound(x)
  SStotal = SStotal + (Y(i) - ymean) ^ 2
  FV = -ymean
  For j = 0 To orden
    FV = FV + Coeff(j) * x(i) ^ j
  Next
  SSregr = SSregr + FV ^ 2
Next
             
r = SSregr / SStotal
             
polregr = True
GoTo slut
fejl:
  polregr = False
'  MsgBox Err.Number
slut:
End Function
Function polregrkvad(x() As Double, Y() As Double, Optional ByVal tving As Boolean = False) As Boolean
'Udf*oe*rer polynomisk regression p*aa* to arrays og returnerer a,b,c og R^2 via userdefined datatype
' tving betyder tving gennem (0,0)
'  ax^2+bx+c
Dim sumx As Double, sumy As Double, sumxy As Double, sumxx As Double, sumyy As Double
Dim n As Integer, i As Integer, j As Integer, q As Integer, ymean As Double
Dim sumx2y As Double, sumx3 As Double, sumx4 As Double, SStotal As Double, SSregr As Double
Dim M(3, 4)

On Error GoTo fejl

n = UBound(x) - LBound(x) + 1
               
For i = LBound(x) To UBound(x)
    sumx = sumx + x(i)
    sumy = sumy + Y(i)
    sumxy = sumxy + x(i) * Y(i)
    sumx2y = sumx2y + x(i) ^ 2 * Y(i)
    sumxx = sumxx + x(i) ^ 2
    sumx3 = sumx3 + x(i) ^ 3
    sumx4 = sumx4 + x(i) ^ 4
Next

If tving = False Then

M(1, 1) = sumxx
M(1, 2) = sumx
M(1, 3) = n
M(1, 4) = sumy
M(2, 1) = sumx3
M(2, 2) = sumxx
M(2, 3) = sumx
M(2, 4) = sumxy
M(3, 1) = sumx4
M(3, 2) = sumx3
M(3, 3) = sumxx
M(3, 4) = sumx2y

' Der udf*oe*res gaussisk eliminination
For i = 1 To 3
' dividerer r*ae*kke i med v*ae*rdien i celle (i,i)
  For j = 4 To 1 Step -1
    If M(i, i) = 0 Then
'        MsgBox "nul"
        GoTo tving
    Else
        M(i, j) = M(i, j) / M(i, i)
    End If
  Next
' laver nuller i kolonne i undtagen celle i,i hvor der skal st*aa* et 1
  For q = 1 To 3
    If q <> i Then
      For j = 4 To 1 Step -1
        M(q, j) = M(q, j) - M(i, j) * M(q, i)
      Next
    End If
  Next
Next
A = M(1, 4)
b = M(2, 4)
c = M(3, 4)
            
Else
tving:
A = (sumx2y * sumxx - sumx3 * sumxy) / (sumx4 * sumxx - sumx3 ^ 2)
b = (sumxy - sumx3 * A) / sumxx
c = 0

End If
             
ymean = sumy / n

For i = LBound(x) To UBound(x)
  SStotal = SStotal + (Y(i) - ymean) ^ 2
  SSregr = SSregr + (A * x(i) ^ 2 + b * x(i) + c - ymean) ^ 2
Next
             
r = SSregr / SStotal
             
polregrkvad = True
GoTo slut
fejl:
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

    Selection.TypeText text:="R"
    Selection.Font.Superscript = wdToggle
    Selection.TypeText text:="2"
    Selection.Font.Superscript = wdToggle
    Selection.TypeText text:=" = " & ConvertNumberToString(r)
    Selection.TypeParagraph
    Selection.Collapse
    Selection.OMaths.Add Range:=Selection.Range
    Selection.TypeText text:=Ligning
    Selection.OMaths.BuildUp
    Selection.TypeParagraph
'    Call Inds*ae*t2DPlot(ligning, xval, yval)
#If Mac Then
#Else
        Oundo.EndCustomRecord
#End If

End Sub


