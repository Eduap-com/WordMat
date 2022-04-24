VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormDeSolveNumeric 
   Caption         =   "Løs differentialligning(er) numerisk"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   150
   ClientWidth     =   16050
   OleObjectBlob   =   "UserFormDeSolveNumeric.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormDeSolveNumeric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public luk As Boolean
Public InsertType As Integer
Public ListOutput As String
Private PointArr() As String
Public xlwb As Object
'Public xlwb As Excel.Workbook


Private Sub CommandButton_cancel_Click()
    luk = True
    On Error Resume Next
    If MaxProc.Finished = 0 Then
        MaxProc.CloseProcess
        MaxProc.StartMaximaProcess
    End If
    Unload Me
End Sub

Private Sub CommandButton_geogebra_Click()
    Dim s As String, i As Long, xl As String, yl As String, j As Long
    If Not SolveDE Then
        MsgBox Err.Description, vbOKOnly, "Error calculating points"
        Exit Sub
    End If
'    s = "{"
'    For i = 0 To UBound(PointArr)
'        s = s & "(" & Replace(PointArr(i, 1), ",", ".") & "," & Replace(PointArr(i, 2), ",", ".") & "),"
'    Next
'    s = Left(s, Len(s) - 1)
'    s = s & "}"
    For i = 0 To UBound(PointArr)
        xl = xl & Trim(Replace(Replace(PointArr(i, 0), ",", "."), ChrW(183), "*")) & ","
    Next
    If Len(xl) > 1 Then xl = Left(xl, Len(xl) - 1)
    For j = 1 To UBound(PointArr, 2)
        yl = ""
        For i = 0 To UBound(PointArr)
            yl = yl & Trim(Replace(Replace(PointArr(i, j), ",", "."), ChrW(183), "*")) & ","
        Next
        yl = Left(yl, Len(yl) - 1)
        s = s & "LineGraph({" & xl & "},{" & yl & "});"
    Next
    s = Left(s, Len(s) - 1)
    If Len(xl) > 1 Then
        OpenGeoGebraWeb s, "", False, False
        Label_wait.Caption = "GeoGebra opened"
    Else
        Label_wait.Caption = "No point calculated"
    End If
End Sub

Private Sub CommandButton_insertgraph_Click()
Dim ils As InlineShape
Dim Sep As String, s As String
Dim pointText As String, i As Long
Dim pointText2 As String
    On Error GoTo fejl
    InsertType = 1
    If ListOutput = vbNullString Then SolveDE
    PlotOutput 3
    
    For i = 0 To UBound(PointArr)
        pointText = pointText & PointArr(i, 0) & ListSeparator & PointArr(i, 1) & vbCrLf
    Next
    If UBound(PointArr, 2) > 1 Then
    For i = 0 To UBound(PointArr)
        pointText2 = pointText2 & PointArr(i, 0) & ListSeparator & PointArr(i, 2) & vbCrLf
    Next
    End If
    
    If Selection.OMaths.Count > 0 Then
        omax.GoToEndOfSelectedMaths
    End If
    If Selection.Tables.Count > 0 Then
        Selection.Tables(Selection.Tables.Count).Select
        Selection.Collapse wdCollapseEnd
    End If
    Selection.MoveRight wdCharacter, 1
    Selection.TypeParagraph

#If Mac Then
    Set ils = Selection.InlineShapes.AddPicture(GetTempDir() & "WordMatGraf.pdf", False, True)
#Else
    Set ils = Selection.InlineShapes.AddPicture(GetTempDir() & "WordMatGraf.gif", False, True)
#End If
Sep = "|"
s = "WordMat" & Sep & AppVersion & Sep & TextBox_definitioner.text & Sep & "" & Sep & TextBox_varx.text & Sep & TextBox_var1.text & Sep
s = s & TextBox_xmin.text & Sep & TextBox_xmax.text & Sep & "" & Sep & "" & Sep
s = s & "" & Sep & "" & Sep & "" & Sep & TextBox_ymin.text & Sep & TextBox_ymax.text & Sep
s = s & "" & Sep & "" & Sep & "" & Sep & "" & Sep & "" & Sep
s = s & "" & Sep & "" & Sep & "" & Sep & "" & Sep & "" & Sep
s = s & "" & Sep & "" & Sep & "" & Sep & "" & Sep & "" & Sep
s = s & "" & Sep & "" & Sep & "" & Sep & "" & Sep & "" & Sep
s = s & "" & Sep & "" & Sep & "" & Sep & "" & Sep & "" & Sep
s = s & "" & Sep & "" & Sep & "" & Sep
s = s & "" & Sep & "" & Sep & "" & Sep & "" & Sep
s = s & "" & Sep & "" & Sep & "" & Sep & "" & Sep
s = s & "" & Sep & "" & Sep & "" & Sep & "" & Sep
s = s & pointText & Sep & pointText2 & Sep & "" & Sep & CheckBox_pointsjoined.Value & Sep & CheckBox_pointsjoined.Value & Sep & "2" & Sep & "2" & Sep
s = s & "" & Sep
s = s & "" & Sep
s = s & "true" & Sep & "false" & Sep & "false" & Sep & "false" & Sep


ils.AlternativeText = s
Unload Me
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton_inserttabel_Click()
Dim Tabel As Table
Dim i As Long, j As Integer
    On Error GoTo fejl
    If ListOutput = vbNullString Then SolveDE
    InsertType = 2
        Application.ScreenUpdating = False
        Selection.Collapse wdCollapseEnd
                
        
        Set Tabel = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=UBound(PointArr, 1) + 2, NumColumns:= _
        UBound(PointArr, 2) + 1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed)
        With Tabel
'            .Style = WdBuiltinStyle.WdBuiltinStyle.wdStyleNormalTable ' p*aa* 2013 giver det ingen kanter
'        If .Style <> "Tabel - Gitter" And InStr(.Style, "Table") < 0 Then
'            On Error Resume Next
'            .Style = "Tabel - Gitter" ' duer ikke p*aa* udenlandsk
'        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
        .Cell(1, 1).Range.text = TextBox_varx.text
        .Cell(1, 1).Range.Bold = True
        .Columns(1).Width = 65
        i = 2
        If CheckBox1.Value Then
            .Cell(1, i).Range.text = TextBox_var1.text
            .Cell(1, i).Range.Bold = True
            .Columns(i).Width = 65
            i = i + 1
        End If
        If CheckBox2.Value Then
            .Cell(1, i).Range.text = TextBox_var2.text
            .Cell(1, i).Range.Bold = True
            .Columns(i).Width = 65
            i = i + 1
        End If
        If CheckBox3.Value Then
            .Cell(1, i).Range.text = TextBox_var3.text
            .Cell(1, i).Range.Bold = True
            .Columns(i).Width = 65
            i = i + 1
        End If
        If CheckBox4.Value Then
            .Cell(1, i).Range.text = TextBox_var4.text
            .Cell(1, i).Range.Bold = True
            .Columns(i).Width = 65
            i = i + 1
        End If
        If CheckBox5.Value Then
            .Cell(1, i).Range.text = TextBox_var5.text
            .Cell(1, i).Range.Bold = True
            .Columns(i).Width = 65
            i = i + 1
        End If
        If CheckBox6.Value Then
            .Cell(1, i).Range.text = TextBox_var6.text
            .Cell(1, i).Range.Bold = True
            .Columns(i).Width = 65
            i = i + 1
        End If
        If CheckBox7.Value Then
            .Cell(1, i).Range.text = TextBox_var7.text
            .Cell(1, i).Range.Bold = True
            .Columns(i).Width = 65
            i = i + 1
        End If
        If CheckBox8.Value Then
            .Cell(1, i).Range.text = TextBox_var8.text
            .Cell(1, i).Range.Bold = True
            .Columns(i).Width = 65
            i = i + 1
        End If
        If CheckBox9.Value Then
            .Cell(1, i).Range.text = TextBox_var9.text
            .Cell(1, i).Range.Bold = True
            .Columns(i).Width = 65
            i = i + 1
        End If
        
        End With
        
    For i = 0 To UBound(PointArr, 1)
        For j = 0 To UBound(PointArr, 2)
            Tabel.Cell(i + 2, j + 1).Range.text = PointArr(i, j)
        Next
    Next
    
    Unload Me
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton_opdater_Click()
    SolveDE
    PlotOutput
End Sub

Private Sub CommandButton_toExcel_Click()
'    Dim ws As excel.Worksheet
    Dim ws As Object 'excel.Worksheet
    Dim i As Long, j As Integer
    On Error Resume Next
    InsertType = 4
If XLapp Is Nothing Then
'    Set XLapp = New excel.Application
    Set XLapp = CreateObject("Excel.Application") 'New excel.Application
End If
XLapp.visible = True
    Set xlwb = XLapp.Workbooks.Add
    
    Set ws = xlwb.worksheets(1)
    
    ws.Cells(2, 1) = TextBox_varx.text
    i = 2
    If TextBox_var1.text <> vbNullString And TextBox_eq1.text <> vbNullString And TextBox_init1.text <> vbNullString Then
        ws.Cells(2, i) = TextBox_var1.text
        i = i + 1
    End If
    If TextBox_var2.text <> vbNullString And TextBox_eq2.text <> vbNullString And TextBox_init2.text <> vbNullString Then
        ws.Cells(2, i) = TextBox_var2.text
        i = i + 1
    End If
    If TextBox_var3.text <> vbNullString And TextBox_eq3.text <> vbNullString And TextBox_init3.text <> vbNullString Then
        ws.Cells(2, i) = TextBox_var3.text
        i = i + 1
    End If
    If TextBox_var4.text <> vbNullString And TextBox_eq4.text <> vbNullString And TextBox_init4.text <> vbNullString Then
        ws.Cells(2, i) = TextBox_var4.text
        i = i + 1
    End If
    If TextBox_var5.text <> vbNullString And TextBox_eq5.text <> vbNullString And TextBox_init5.text <> vbNullString Then
        ws.Cells(2, i) = TextBox_var5.text
        i = i + 1
    End If
    If TextBox_var6.text <> vbNullString And TextBox_eq6.text <> vbNullString And TextBox_init6.text <> vbNullString Then
        ws.Cells(2, i) = TextBox_var6.text
        i = i + 1
    End If
    If TextBox_var7.text <> vbNullString And TextBox_eq7.text <> vbNullString And TextBox_init7.text <> vbNullString Then
        ws.Cells(2, i) = TextBox_var7.text
        i = i + 1
    End If
    If TextBox_var8.text <> vbNullString And TextBox_eq8.text <> vbNullString And TextBox_init8.text <> vbNullString Then
        ws.Cells(2, i) = TextBox_var8.text
        i = i + 1
    End If
    If TextBox_var9.text <> vbNullString And TextBox_eq9.text <> vbNullString And TextBox_init9.text <> vbNullString Then
        ws.Cells(2, i) = TextBox_var9.text
        i = i + 1
    End If
    
    For i = 0 To UBound(PointArr, 1)
        For j = 0 To UBound(PointArr, 2)
            ws.Cells(i + 3, j + 1) = "=" & ConvertNumberToExcel(PointArr(i, j))
        Next
    Next
    Unload Me
End Sub
Function ConvertNumberToExcel(n As String) As String
    n = Replace(n, ",", ".")
    n = Replace(n, VBA.ChrW(183), "*")
    ConvertNumberToExcel = n
End Function
Private Sub CommandButton_tolist_Click()
    InsertType = 3
    Unload Me
End Sub

Private Sub TextBox_var2_AfterUpdate()
    OpdaterDefinitioner
End Sub
Private Sub TextBox_var3_AfterUpdate()
    OpdaterDefinitioner
End Sub

Private Sub UserForm_Activate()
On Error Resume Next
InsertType = 0
    SetCaptions
    Label_wait.visible = False
#If Mac Then
    Me.Left = 0
    Me.top = 350
    CommandButton_opdater.visible = False
    CommandButton_toExcel.visible = False
    CommandButton_insertgraph.visible = False
    CheckBox_pointsjoined.visible = False
    CheckBox_visforklaring.visible = False
    TextBox_ymin.visible = False
    TextBox_ymax.visible = False
    Label16.visible = False
    Label17.visible = False
    Label_wait.Caption = ""
    Kill GetTempDir() & "WordMatGraf.pdf"
#Else
    Kill GetTempDir() & "\WordMatGraf.gif"
#End If
    OpdaterDefinitioner
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    luk = True
End Sub

Function SolveDE() As Boolean
    Dim variabel As String, xmin As String, xmax As String, xstep As String, DElist As String, varlist As String, guesslist As String
    Dim ea As New ExpressionAnalyser
    Dim n As Integer, Npoints As Long
    On Error GoTo fejl
    variabel = TextBox_varx.text
    xmin = Replace(TextBox_xmin.text, ",", ".")
    xmax = Replace(TextBox_xmax.text, ",", ".")
    xstep = Replace(TextBox_step.text, ",", ".")
    varlist = "["
    guesslist = "["
    DElist = "["
    If TextBox_var1.text = vbNullString Or TextBox_eq1.text = vbNullString Or TextBox_init1.text = vbNullString Then
        MsgBox "Der mangler data", vbOKOnly, Sprog.Error
        GoTo slut
    Else
        n = n + 1
        varlist = varlist & TextBox_var1.text & ","
        guesslist = guesslist & Replace(TextBox_init1.text, ",", ".") & " ,"
        DElist = DElist & TextBox_eq1.text & " ,"
    End If
    If TextBox_var2.text <> vbNullString And TextBox_eq2.text <> vbNullString And TextBox_init2.text <> vbNullString Then
        n = n + 1
        varlist = varlist & TextBox_var2.text & ","
        guesslist = guesslist & Replace(TextBox_init2.text, ",", ".") & " ,"
        DElist = DElist & TextBox_eq2.text & " ,"
    End If
    If TextBox_var3.text <> vbNullString And TextBox_eq3.text <> vbNullString And TextBox_init3.text <> vbNullString Then
        n = n + 1
        varlist = varlist & TextBox_var3.text & ","
        guesslist = guesslist & Replace(TextBox_init3.text, ",", ".") & " ,"
        DElist = DElist & TextBox_eq3.text & " ,"
    End If
    If TextBox_var4.text <> vbNullString And TextBox_eq4.text <> vbNullString And TextBox_init4.text <> vbNullString Then
        n = n + 1
        varlist = varlist & TextBox_var4.text & ","
        guesslist = guesslist & Replace(TextBox_init4.text, ",", ".") & " ,"
        DElist = DElist & TextBox_eq4.text & " ,"
    End If
    If TextBox_var5.text <> vbNullString And TextBox_eq5.text <> vbNullString And TextBox_init5.text <> vbNullString Then
        n = n + 1
        varlist = varlist & TextBox_var5.text & ","
        guesslist = guesslist & Replace(TextBox_init5.text, ",", ".") & " ,"
        DElist = DElist & TextBox_eq5.text & " ,"
    End If
    If TextBox_var6.text <> vbNullString And TextBox_eq6.text <> vbNullString And TextBox_init6.text <> vbNullString Then
        n = n + 1
        varlist = varlist & TextBox_var6.text & ","
        guesslist = guesslist & Replace(TextBox_init6.text, ",", ".") & " ,"
        DElist = DElist & TextBox_eq6.text & " ,"
    End If
    If TextBox_var7.text <> vbNullString And TextBox_eq7.text <> vbNullString And TextBox_init7.text <> vbNullString Then
        n = n + 1
        varlist = varlist & TextBox_var7.text & ","
        guesslist = guesslist & Replace(TextBox_init7.text, ",", ".") & " ,"
        DElist = DElist & TextBox_eq7.text & " ,"
    End If
    If TextBox_var8.text <> vbNullString And TextBox_eq8.text <> vbNullString And TextBox_init8.text <> vbNullString Then
        n = n + 1
        varlist = varlist & TextBox_var8.text & ","
        guesslist = guesslist & Replace(TextBox_init8.text, ",", ".") & " ,"
        DElist = DElist & TextBox_eq8.text & " ,"
    End If
    If TextBox_var9.text <> vbNullString And TextBox_eq9.text <> vbNullString And TextBox_init9.text <> vbNullString Then
        n = n + 1
        varlist = varlist & TextBox_var9.text & ","
        guesslist = guesslist & Replace(TextBox_init9.text, ",", ".") & " ,"
        DElist = DElist & TextBox_eq9.text & " ,"
    End If
    
    Npoints = (val(Replace(TextBox_xmax.text, ",", ".")) - val(Replace(TextBox_xmin.text, ",", "."))) / val(Replace(TextBox_step.text, ",", "."))
    varlist = Left(varlist, Len(varlist) - 1) & "]"
    guesslist = Left(guesslist, Len(guesslist) - 1) & "]"
    DElist = Left(DElist, Len(DElist) - 1) & "]"
    
    omax.PrepareNewCommand finddef:=False  ' uden at s*oe*ge efter definitioner i dokument
    InsertDefinitioner
    omax.SolveDENumeric variabel, xmin, xmax, xstep, varlist, guesslist, DElist
    ListOutput = omax.MaximaOutput
    
    Dim s As String, i As Long, j As Integer
    Dim Arr As Variant
    ReDim PointArr(Npoints, n)
    ea.text = ListOutput
    ea.SetSquareBrackets
    If ea.Length > 2 Then
        ea.text = Mid(ea.text, 2, ea.Length - 2)
    End If
    Do
        s = ea.GetNextBracketContent(0)
        Arr = Split(s, ListSeparator)
        For j = 0 To n 'UBound(Arr)
            PointArr(i, j) = Arr(j)
        Next
        i = i + 1
    Loop While ea.Pos < ea.Length - 1 And i < 1000
SolveDE = True
GoTo slut
fejl:
    SolveDE = False
slut:
End Function

Sub PlotOutput(Optional highres As Double = 1)
Dim text As String, yAxislabel As String
On Error GoTo fejl
    Label_wait.Caption = Sprog.Wait & "!"
    Label_wait.Font.Size = 36
    Label_wait.visible = True
    omax.PrepareNewCommand finddef:=False  ' uden at s*oe*ge efter definitioner i dokument
    
'    text = "explicit(x^2,x,-1,1)"
    If Len(TextBox_ymin.text) > 0 And Len(TextBox_ymax.text) > 0 Then
        text = text & "yrange=[" & ConvertNumberToMaxima(TextBox_ymin.text) & "," & ConvertNumberToMaxima(TextBox_ymax.text) & "],"
    End If
    colindex = 0
    text = text & "color=" & GetNextColor & ","
    If Not CheckBox_pointsjoined.Value Then
        text = text & "point_size=" & Replace(highres * 1, ",", ".") & ","
    Else
#If Mac Then
        text = text & "point_size=0.1," ' fejler med 0 p*aa* mac
#Else
        text = text & "point_size=0,"
#End If
    End If
    text = text & "point_type=filled_circle,points_joined=" & VBA.LCase(CheckBox_pointsjoined.Value) & ","
    If CheckBox1.Value Then
        If CheckBox_visforklaring.Value Then
            text = text & "key=""" & omax.ConvertToAscii(TextBox_var1.text) & ""","
        Else
            text = text & "key="""","
        End If
        text = text & "points(makelist([pq[1],pq[2]],pq,qDElist)),"
        yAxislabel = yAxislabel & TextBox_var1.text & ","
    End If
    If CheckBox2.Value Then
        If CheckBox_visforklaring.Value Then
            text = text & "key=""" & omax.ConvertToAscii(TextBox_var2.text) & ""","
        Else
            text = text & "key="""","
        End If
        text = text & "color=" & GetNextColor & ","
        text = text & "points(makelist([pq[1],pq[3]],pq,qDElist)),"
        yAxislabel = yAxislabel & TextBox_var2.text & ","
    End If
    If CheckBox3.Value Then
        If CheckBox_visforklaring.Value Then
            text = text & "key=""" & omax.ConvertToAscii(TextBox_var3.text) & ""","
        Else
            text = text & "key="""","
        End If
        text = text & "color=" & GetNextColor & ","
        text = text & "points(makelist([pq[1],pq[4]],pq,qDElist)),"
        yAxislabel = yAxislabel & TextBox_var3.text & ","
    End If
    If CheckBox4.Value Then
        If CheckBox_visforklaring.Value Then
            text = text & "key=""" & omax.ConvertToAscii(TextBox_var4.text) & ""","
        Else
            text = text & "key="""","
        End If
        text = text & "color=" & GetNextColor & ","
        text = text & "points(makelist([pq[1],pq[5]],pq,qDElist)),"
        yAxislabel = yAxislabel & TextBox_var4.text & ","
    End If
    If CheckBox5.Value Then
        If CheckBox_visforklaring.Value Then
            text = text & "key=""" & omax.ConvertToAscii(TextBox_var5.text) & ""","
        Else
            text = text & "key="""","
        End If
        text = text & "color=" & GetNextColor & ","
        text = text & "points(makelist([pq[1],pq[6]],pq,qDElist)),"
        yAxislabel = yAxislabel & TextBox_var5.text & ","
    End If
    If CheckBox6.Value Then
        If CheckBox_visforklaring.Value Then
            text = text & "key=""" & omax.ConvertToAscii(TextBox_var6.text) & ""","
        Else
            text = text & "key="""","
        End If
        text = text & "color=" & GetNextColor & ","
        text = text & "points(makelist([pq[1],pq[7]],pq,qDElist)),"
        yAxislabel = yAxislabel & TextBox_var6.text & ","
    End If
    If CheckBox7.Value Then
        If CheckBox_visforklaring.Value Then
            text = text & "key=""" & omax.ConvertToAscii(TextBox_var7.text) & ""","
        Else
            text = text & "key="""","
        End If
        text = text & "color=" & GetNextColor & ","
        text = text & "points(makelist([pq[1],pq[8]],pq,qDElist)),"
        yAxislabel = yAxislabel & TextBox_var7.text & ","
    End If
    If CheckBox8.Value Then
        If CheckBox_visforklaring.Value Then
            text = text & "key=""" & omax.ConvertToAscii(TextBox_var8.text) & ""","
        Else
            text = text & "key="""","
        End If
        text = text & "color=" & GetNextColor & ","
        text = text & "points(makelist([pq[1],pq[9]],pq,qDElist)),"
        yAxislabel = yAxislabel & TextBox_var8.text & ","
    End If
    If CheckBox9.Value Then
        If CheckBox_visforklaring.Value Then
            text = text & "key=""" & omax.ConvertToAscii(TextBox_var9.text) & ""","
        Else
            text = text & "key="""","
        End If
        text = text & "color=" & GetNextColor & ","
        text = text & "points(makelist([pq[1],pq[10]],pq,qDElist)),"
        yAxislabel = yAxislabel & TextBox_var9.text & ","
    End If
    text = Left(text, Len(text) - 1)
    yAxislabel = Left(yAxislabel, Len(yAxislabel) - 1)
'    text = text & ",[xlabel,""" & TextBox_varx.text & """]"
'    text = text & ",[ylabel,""" & TextBox_var1.text & """]"
    
    If Len(text) > 0 Then
        Call omax.Draw2D(text, "", TextBox_varx.text, yAxislabel, True, True, 1)
        If omax.MaximaOutput = "" Then
            Label_wait.Caption = "Fejl!"
            Label_wait.visible = True
            GoTo slut
        Else
            DoEvents
#If Mac Then
'            If highres <> 3 Then Image1.Picture = LoadPicture(GetTempDir() & "WordMatGraf.pdf")
            ShowPreviewMac
#Else
            If highres <> 3 Then Image1.Picture = LoadPicture(GetTempDir() & "WordMatGraf.gif")
#End If
        End If
    Else
'        Label_wait.Caption = " indtast funktion og Tryk opdater"
        Label_wait.visible = False
    End If
    Label_wait.visible = False
GoTo slut
fejl:
    On Error Resume Next
    Label_wait.Caption = Sprog.A(94)
    Label_wait.Font.Size = 12
    Label_wait.Width = 150
    Label_wait.visible = True
    Image1.Picture = Nothing
slut:

End Sub

Sub InsertDefinitioner()
' inds*ae*tter definitioner fra textboxen i maximainputstring
Dim DefString As String

omax.InsertKillDef

DefString = GetDefString

If Len(DefString) > 0 Then
'defstring = Replace(defstring, ",", ".")
'defstring = Replace(defstring, ";", ",")
'defstring = Replace(defstring, "=", ":")
If right(DefString, 1) = "," Then DefString = Left(DefString, Len(DefString) - 1)

'omax.MaximaInputStreng = omax.MaximaInputStreng & "[" & defstring & "]$"
omax.MaximaInputStreng = omax.MaximaInputStreng & DefString
End If
End Sub
Function GetDefString()
Dim DefString As String
omax.ResetDefinitions
DefString = TextBox_definitioner.text
If Len(DefString) > 0 Then
DefString = Replace(DefString, vbCrLf, ListSeparator)
    DefString = TrimB(DefString, ListSeparator)
Do While InStr(DefString, ListSeparator & ListSeparator) > 0
    DefString = Replace(DefString, ListSeparator & ListSeparator, ListSeparator) ' dobbelt ;; fjernes
Loop
DefString = omax.AddDefinition("definer:" & DefString)
GetDefString = DefString
End If
End Function

Sub OpdaterDefinitioner()
' ser efter variable i textboxene og inds*ae*tter under definitioner
Dim vars As String
Dim var As String, var2 As String
Dim ea As New ExpressionAnalyser
Dim ea2 As New ExpressionAnalyser
Dim Arr As Variant
Dim arr2 As Variant
Dim i As Integer
    
    
    vars = vars & GetTextboxVars(TextBox_eq1, TextBox_varx)
    vars = vars & GetTextboxVars(TextBox_eq2, TextBox_varx)
    vars = vars & GetTextboxVars(TextBox_eq3, TextBox_varx)
    vars = vars & GetTextboxVars(TextBox_eq4, TextBox_varx)
    vars = vars & GetTextboxVars(TextBox_eq5, TextBox_varx)
    vars = vars & GetTextboxVars(TextBox_eq6, TextBox_varx)
    vars = vars & GetTextboxVars(TextBox_eq7, TextBox_varx)
    vars = vars & GetTextboxVars(TextBox_eq8, TextBox_varx)
    vars = vars & GetTextboxVars(TextBox_eq9, TextBox_varx)
    
    omax.FindVariable vars, False ' fjerner dobbelte
    vars = omax.vars
    vars = RemoveVar(vars, TextBox_var1.text)
    vars = RemoveVar(vars, TextBox_var2.text)
    vars = RemoveVar(vars, TextBox_var3.text)
    vars = RemoveVar(vars, TextBox_var4.text)
    vars = RemoveVar(vars, TextBox_var5.text)
    vars = RemoveVar(vars, TextBox_var6.text)
    vars = RemoveVar(vars, TextBox_var7.text)
    vars = RemoveVar(vars, TextBox_var8.text)
    vars = RemoveVar(vars, TextBox_var9.text)
    
    If Left(vars, 1) = ";" Then vars = right(vars, Len(vars) - 1)
    
    ea.text = vars
    Do While right(TextBox_definitioner.text, 2) = vbCrLf
        TextBox_definitioner.text = Left(TextBox_definitioner.text, Len(TextBox_definitioner.text) - 2)
    Loop
    Arr = Split(TextBox_definitioner.text, vbCrLf)
    
    Do
    var = ea.GetNextListItem
    var = Replace(var, vbCrLf, "")
    For i = 0 To UBound(Arr)
        If Arr(i) <> "" Then
        var2 = Split(Arr(i), "=")(0)
        If var2 = var Then
            var = ""
            Exit For
        End If
        End If
    Next
    If var <> "" Then
'        If Right(TextBox_definitioner.text, 2) <> vbCrLf Then
        If Len(TextBox_definitioner.text) > 0 Then
            TextBox_definitioner.text = TextBox_definitioner.text & vbCrLf
        End If
        TextBox_definitioner.text = TextBox_definitioner.text & var & "=1"
    End If
    Loop While ea.Pos <= Len(ea.text)

    
End Sub
Function GetTextboxVars(tb As TextBox, tbvar As TextBox) As String
Dim ea As New ExpressionAnalyser
    If Len(tb.text) > 0 Then
        omax.vars = ""
        omax.FindVariable tb.text, False
        omax.vars = RemoveVar(omax.vars, tbvar.text)
        If Len(omax.vars) > 0 Then
            GetTextboxVars = ";" & omax.vars
        End If
    End If
End Function

Function RemoveVar(text As String, var As String)
' fjerner var fra string
Dim ea As New ExpressionAnalyser
If var = vbNullString Then
    RemoveVar = text
    Exit Function
End If
ea.text = text
Call ea.ReplaceVar(var, "")
text = Replace(ea.text, ";;", ";")
If Left(text, 1) = ";" Then text = right(text, Len(text) - 1)
If right(text, 1) = ";" Then text = Left(text, Len(text) - 1)

RemoveVar = text
End Function

Sub SetCaptions()
    Me.Caption = Sprog.A(85)
    Label6.Caption = Sprog.A(86)
    Label7.Caption = Sprog.A(87)
    Label_Graf.Caption = Sprog.Graph
    CommandButton_opdater.Caption = Sprog.Update
    CommandButton_cancel.Caption = Sprog.Cancel
    Label_var.Caption = Sprog.IndepVar
    Label3.Caption = Sprog.A(88)
    Label5.Caption = Sprog.Definitions
    Label_wait.Caption = Sprog.Wait & "!"
    CheckBox_pointsjoined.Caption = Sprog.A(89)
    CheckBox_visforklaring.Caption = Sprog.A(90)
    CommandButton_tolist.Caption = Sprog.A(91)
    CommandButton_inserttabel.Caption = Sprog.A(92)
    CommandButton_insertgraph.Caption = Sprog.A(93)
    
End Sub
Sub ShowPreviewMac()
#If Mac Then
    RunScript "OpenPreview", GetTempDir() & "WordMatGraf.pdf"
#End If
End Sub

