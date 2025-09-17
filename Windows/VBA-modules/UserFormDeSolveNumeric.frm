VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormDeSolveNumeric 
   Caption         =   "Solve differential equations numerically"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   150
   ClientWidth     =   16725
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
Private EventsCol As New Collection
Sub SetEscEvents(ControlColl As Controls)
' SetEscEvents Me.Controls     in Initialize
    Dim CE As CEvents, c As control, TN As String, F As MSForms.Frame
    On Error Resume Next
    For Each c In ControlColl ' Me.Controls
        TN = TypeName(c)
        If TN = "CheckBox" Then
            Set CE = New CEvents: Set CE.CheckBoxControl = c: EventsCol.Add CE
        ElseIf TN = "OptionButton" Then
            Set CE = New CEvents: Set CE.OptionButtonControl = c: EventsCol.Add CE
        ElseIf TN = "ComboBox" Then
            Set CE = New CEvents: Set CE.ComboBoxControl = c: EventsCol.Add CE
        ElseIf TN = "Label" Then
            Set CE = New CEvents: Set CE.LabelControl = c: EventsCol.Add CE
        ElseIf TN = "TextBox" Then
            Set CE = New CEvents: Set CE.TextBoxControl = c: EventsCol.Add CE
        ElseIf TN = "CommandButton" Then
            Set CE = New CEvents: Set CE.CommandButtonControl = c: EventsCol.Add CE
        ElseIf TN = "ListBox" Then
            Set CE = New CEvents: Set CE.ListBoxControl = c: EventsCol.Add CE
        ElseIf TN = "Frame" Then
            Set F = c
            SetEscEvents F.Controls
        End If
    Next
End Sub
Private Sub CheckBox_autostep_Click()
   If CheckBox_autostep.Value Then
      UpdateStep
   End If
End Sub

Private Sub ComboBox_graphapp_Change()
   If ComboBox_graphapp.ListIndex > 0 Then
      Label_insertgraph.visible = False
      CheckBox_pointsjoined.visible = False
      CheckBox_visforklaring.visible = False
      Me.Width = 347
   Else
      Label_insertgraph.visible = True
      CheckBox_pointsjoined.visible = True
      CheckBox_visforklaring.visible = True
      Me.Width = 848
   End If
   Validate
End Sub

Private Sub Label_cancel_Click()
    luk = True
    On Error Resume Next
#If Mac Then
#Else
    If MaxProc.Finished = 0 Then
        MaxProc.CloseProcess
        MaxProc.StartMaximaProcess
    End If
#End If
    Unload Me
End Sub

Private Sub GeoGebraPlot()
    Dim s As String, i As Long, xl As String, yl As String, j As Long
    Dim Y As Double, ymax As Double, ymin As Double
    ymax = -10000000
    ymin = 10000000
    Erase PointArr
    If Not SolveDE Then ' first calculate points with Maxima
        MsgBox Err.Description, vbOKOnly, "Error calculating points"
        Exit Sub
    End If

'    s = "{"
'    For i = 0 To UBound(PointArr)
'        s = s & "(" & Replace(PointArr(i, 1), ",", ".") & "," & Replace(PointArr(i, 2), ",", ".") & "),"
'    Next
'    s = left$(s, Len(s) - 1)
'    s = s & "}"
    For i = 0 To UBound(PointArr)
        xl = xl & Trim$(Replace(Replace(PointArr(i, 0), ",", "."), ChrW$(183), "*")) & ","
    Next
    If Len(xl) > 1 Then xl = Left$(xl, Len(xl) - 1)
    For j = 1 To UBound(PointArr, 2)
        yl = ""
        If (j = 1 And CheckBox1.Value) Or (j = 2 And CheckBox2.Value) Or (j = 3 And CheckBox3.Value) Then
        For i = 0 To UBound(PointArr)
            Y = val(Trim$(Replace(Replace(PointArr(i, j), ",", "."), ChrW$(183), "*")))
            If Y > ymax Then
               ymax = Y
            End If
            If Y < ymin Then
               ymin = Y
            End If
            yl = yl & Replace(Y, ",", ".") & ","
        Next
        yl = Left$(yl, Len(yl) - 1)
        s = s & "LineGraph({" & xl & "},{" & yl & "});"
        End If
    Next
    s = Left$(s, Len(s) - 1)
    If Len(s) > 30000 Then
        Label_wait.Caption = "Too many points for GeoGebra. Decrease no of. steps."
        MsgBox "Too many points for GeoGebra. Decrease no of. steps.", vbOKOnly, "Error"
    ElseIf Len(xl) > 1 Then
      If TextBox_ymin.text <> "" And TextBox_ymax.text <> "" Then
         s = s & ";ZoomIn(" & Replace(TextBox_xmin.text, ",", ".") & "," & Replace(TextBox_ymin.text, ",", ".") & "," & Replace(TextBox_xmax.text, ",", ".") & "," & Replace(TextBox_ymax.text, ",", ".") & ");ZoomIn(0.9)" 'ggbApplet.setCoordinateSystem(0,1000,0,1000)
      Else
         If ymin > 0 And (ymax - ymin) > ymin Then ymin = 0
         s = s & ";ZoomIn(" & Replace(TextBox_xmin.text, ",", ".") & "," & Replace(ymin, ",", ".") & "," & Replace(TextBox_xmax.text, ",", ".") & "," & Replace(ymax, ",", ".") & ");ZoomIn(0.9)" 'ggbApplet.setCoordinateSystem(0,1000,0,1000)
      End If
        OpenGeoGebraWeb s, "", False, False
        Label_wait.Caption = "GeoGebra opened"
    Else
        Label_wait.Caption = "No point calculated"
    End If
End Sub

Private Sub Label_insertgraph_Click()
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
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
Application.ScreenUpdating = True
End Sub

Private Sub Label_inserttabel_Click()
Dim Tabel As Table
Dim i As Long, j As Integer
'    On Error GoTo Fejl
    If ListOutput = vbNullString Then SolveDE
    InsertType = 2
        Application.ScreenUpdating = False
        Selection.Collapse wdCollapseEnd
                
        GoToEndOfMath
        Selection.TypeParagraph
        Set Tabel = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=UBound(PointArr, 1) + 2, NumColumns:= _
        UBound(PointArr, 2) + 1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed)
        With Tabel
'            .Style = WdBuiltinStyle.WdBuiltinStyle.wdStyleNormalTable ' on 2013 no edges
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
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
Application.ScreenUpdating = True
End Sub

Private Sub Label_opdater_Click()
#If Mac Then
   GeoGebraPlot
#Else
   If ComboBox_graphapp.ListIndex = 0 Then
      GnuPlotOpdater
   Else
      GeoGebraPlot
   End If
#End If
End Sub

Private Sub GnuPlotOpdater()
    SolveDE
    PlotOutput
End Sub

Private Sub Label_toExcel_Click()
'    Dim ws As excel.Worksheet
    Dim ws As Object 'excel.Worksheet
    Dim i As Long, j As Integer
    
    Erase PointArr
    If Not SolveDE Then ' first calculate points with Maxima
        MsgBox Err.Description, vbOKOnly, "Error calculating points"
        Exit Sub
    End If
    
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
    n = Replace(n, VBA.ChrW$(183), "*")
    ConvertNumberToExcel = n
End Function
Private Sub Label_tolist_Click()
    InsertType = 3
    Unload Me
End Sub

Private Sub TextBox_eq1_AfterUpdate()
    OpdaterDefinitioner
End Sub

Private Sub TextBox_eq2_AfterUpdate()
    OpdaterDefinitioner
    If TextBox_eq2.text <> vbNullString And TextBox_init2.text = vbNullString Then
      TextBox_init2.text = "1"
    End If
End Sub

Private Sub TextBox_eq3_AfterUpdate()
    OpdaterDefinitioner
    If TextBox_eq3.text <> vbNullString And TextBox_init3.text = vbNullString Then
      TextBox_init3.text = "1"
    End If
End Sub

Private Sub TextBox_eq4_AfterUpdate()
    OpdaterDefinitioner
    If TextBox_eq4.text <> vbNullString And TextBox_init4.text = vbNullString Then
      TextBox_init4.text = "1"
    End If
End Sub

Private Sub TextBox_eq5_AfterUpdate()
    OpdaterDefinitioner
    If TextBox_eq5.text <> vbNullString And TextBox_init5.text = vbNullString Then
      TextBox_init5.text = "1"
    End If
End Sub

Private Sub TextBox_eq6_AfterUpdate()
    OpdaterDefinitioner
    If TextBox_eq6.text <> vbNullString And TextBox_init6.text = vbNullString Then
      TextBox_init6.text = "1"
    End If
End Sub

Private Sub TextBox_eq7_AfterUpdate()
    OpdaterDefinitioner
    If TextBox_eq7.text <> vbNullString And TextBox_init7.text = vbNullString Then
      TextBox_init7.text = "1"
    End If
End Sub

Private Sub TextBox_eq8_AfterUpdate()
    OpdaterDefinitioner
    If TextBox_eq8.text <> vbNullString And TextBox_init8.text = vbNullString Then
      TextBox_init8.text = "1"
    End If
End Sub

Private Sub TextBox_eq9_AfterUpdate()
    OpdaterDefinitioner
    If TextBox_eq9.text <> vbNullString And TextBox_init9.text = vbNullString Then
      TextBox_init9.text = "1"
    End If
End Sub

Private Sub TextBox_step_Change()
   Validate
End Sub

Private Sub TextBox_var2_AfterUpdate()
    OpdaterDefinitioner
End Sub
Private Sub TextBox_var3_AfterUpdate()
    OpdaterDefinitioner
End Sub

Private Sub TextBox_varx_AfterUpdate()
   OpdaterDefinitioner
End Sub
Private Sub TextBox_xmin_Change()
   UpdateStep
End Sub

Private Sub TextBox_xmax_Change()
   UpdateStep
End Sub

Private Sub UpdateStep()
Dim st As Double
   Validate
   If CheckBox_autostep.Value And IsNumeric(TextBox_xmin.text) And IsNumeric(TextBox_xmax.text) Then
      st = (StrToDbl(TextBox_xmax.text) - StrToDbl(TextBox_xmin.text)) / 500
      TextBox_step.text = st
   End If
End Sub

Private Sub Validate()
On Error GoTo slut
   Dim st As Double
   Label_validate.Caption = ""
   Label_validate.visible = False
   If Not IsNumeric(TextBox_xmin.text) Then Label_validate.Caption = "xmin is not a number"
   If Not IsNumeric(TextBox_xmax.text) Then Label_validate.Caption = "xmax is not a number"
   If Not IsNumeric(TextBox_step.text) Then Label_validate.Caption = "Stepsize is not a number"
#If Mac Then
#Else
   If ComboBox_graphapp.ListIndex > 0 Then
#End If
      If IsNumeric(TextBox_xmin.text) And IsNumeric(TextBox_xmax.text) And IsNumeric(TextBox_step.text) Then
         st = Round((StrToDbl(TextBox_xmax.text) - StrToDbl(TextBox_xmin.text)) / StrToDbl(TextBox_step.text), 0)
         If st > 1000 Then Label_validate.Caption = "No of steps is " & st & ". It will probably not work with GeoGebra with that many steps."
      End If
#If Mac Then
#Else
   End If
#End If
slut:
   If Label_validate.Caption <> vbNullString Then Label_validate.visible = True
End Sub
Function StrToDbl(s As String) As Double
   If IsNumeric(s) Then
      s = Replace(s, ",", ".")
      StrToDbl = val(s)
   Else
      StrToDbl = Null
   End If
End Function
Private Sub UserForm_Activate()
 On Error Resume Next
   InsertType = 0
    SetCaptions
    Label_wait.visible = False
#If Mac Then
    Me.Left = 0
    Me.Top = 350
    Label_toExcel.visible = False
    Label_insertgraph.visible = False
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

Private Sub UserForm_Initialize()
#If Mac Then
   Image1.visible = False
   Label_wait.visible = False
   Me.Width = 345
#End If
    SetEscEvents Me.Controls
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    luk = True
    Label_cancel_Click
End Sub

Function SolveDE() As Boolean
    Dim variabel As String, xmin As String, xmax As String, xstep As String, DElist As String, VarList As String, guesslist As String
    Dim ea As New ExpressionAnalyser
    Dim n As Integer, Npoints As Long
    On Error GoTo fejl
    variabel = TextBox_varx.text
    xmin = Replace(TextBox_xmin.text, ",", ".")
    xmax = Replace(TextBox_xmax.text, ",", ".")
    xstep = Replace(TextBox_step.text, ",", ".")
    VarList = "["
    guesslist = "["
    DElist = "["
    If TextBox_var1.text = vbNullString Or TextBox_eq1.text = vbNullString Or TextBox_init1.text = vbNullString Then
        MsgBox "Der mangler data", vbOKOnly, TT.Error
        GoTo slut
    Else
        n = n + 1
        VarList = VarList & TextBox_var1.text & ","
        guesslist = guesslist & Replace(TextBox_init1.text, ",", ".") & " ,"
        DElist = DElist & TextBox_eq1.text & " ,"
    End If
    If TextBox_var2.text <> vbNullString And TextBox_eq2.text <> vbNullString And TextBox_init2.text <> vbNullString Then
        n = n + 1
        VarList = VarList & TextBox_var2.text & ","
        guesslist = guesslist & Replace(TextBox_init2.text, ",", ".") & " ,"
        DElist = DElist & TextBox_eq2.text & " ,"
    End If
    If TextBox_var3.text <> vbNullString And TextBox_eq3.text <> vbNullString And TextBox_init3.text <> vbNullString Then
        n = n + 1
        VarList = VarList & TextBox_var3.text & ","
        guesslist = guesslist & Replace(TextBox_init3.text, ",", ".") & " ,"
        DElist = DElist & TextBox_eq3.text & " ,"
    End If
    If TextBox_var4.text <> vbNullString And TextBox_eq4.text <> vbNullString And TextBox_init4.text <> vbNullString Then
        n = n + 1
        VarList = VarList & TextBox_var4.text & ","
        guesslist = guesslist & Replace(TextBox_init4.text, ",", ".") & " ,"
        DElist = DElist & TextBox_eq4.text & " ,"
    End If
    If TextBox_var5.text <> vbNullString And TextBox_eq5.text <> vbNullString And TextBox_init5.text <> vbNullString Then
        n = n + 1
        VarList = VarList & TextBox_var5.text & ","
        guesslist = guesslist & Replace(TextBox_init5.text, ",", ".") & " ,"
        DElist = DElist & TextBox_eq5.text & " ,"
    End If
    If TextBox_var6.text <> vbNullString And TextBox_eq6.text <> vbNullString And TextBox_init6.text <> vbNullString Then
        n = n + 1
        VarList = VarList & TextBox_var6.text & ","
        guesslist = guesslist & Replace(TextBox_init6.text, ",", ".") & " ,"
        DElist = DElist & TextBox_eq6.text & " ,"
    End If
    If TextBox_var7.text <> vbNullString And TextBox_eq7.text <> vbNullString And TextBox_init7.text <> vbNullString Then
        n = n + 1
        VarList = VarList & TextBox_var7.text & ","
        guesslist = guesslist & Replace(TextBox_init7.text, ",", ".") & " ,"
        DElist = DElist & TextBox_eq7.text & " ,"
    End If
    If TextBox_var8.text <> vbNullString And TextBox_eq8.text <> vbNullString And TextBox_init8.text <> vbNullString Then
        n = n + 1
        VarList = VarList & TextBox_var8.text & ","
        guesslist = guesslist & Replace(TextBox_init8.text, ",", ".") & " ,"
        DElist = DElist & TextBox_eq8.text & " ,"
    End If
    If TextBox_var9.text <> vbNullString And TextBox_eq9.text <> vbNullString And TextBox_init9.text <> vbNullString Then
        n = n + 1
        VarList = VarList & TextBox_var9.text & ","
        guesslist = guesslist & Replace(TextBox_init9.text, ",", ".") & " ,"
        DElist = DElist & TextBox_eq9.text & " ,"
    End If
    
    Npoints = (val(Replace(TextBox_xmax.text, ",", ".")) - val(Replace(TextBox_xmin.text, ",", "."))) / val(Replace(TextBox_step.text, ",", "."))
    VarList = Left$(VarList, Len(VarList) - 1) & "]"
    guesslist = Left$(guesslist, Len(guesslist) - 1) & "]"
    DElist = Left$(DElist, Len(DElist) - 1) & "]"
    
    omax.PrepareNewCommand FindDef:=False  ' without searching for definitions in document
    InsertDefinitioner
    omax.SolveDENumeric variabel, xmin, xmax, xstep, VarList, guesslist, DElist
    ListOutput = omax.MaximaOutput
    
    Dim s As String, i As Long, j As Integer
    Dim arr As Variant
    ReDim PointArr(Npoints, n)
    ea.text = ListOutput
    ea.SetSquareBrackets
    If ea.Length > 2 Then
        ea.text = Mid$(ea.text, 2, ea.Length - 2)
    End If
    Do
        s = ea.GetNextBracketContent(0)
        arr = Split(s, ListSeparator)
        For j = 0 To n 'UBound(Arr)
            PointArr(i, j) = arr(j)
        Next
        i = i + 1
    Loop While ea.pos < ea.Length - 1 And i < 10000
SolveDE = True
GoTo slut
fejl:
   If i >= Npoints Then
    SolveDE = True
   Else
    SolveDE = False
    End If
slut:
End Function

Sub PlotOutput(Optional highres As Double = 1)
Dim text As String, yAxislabel As String
On Error GoTo fejl
    Label_wait.Caption = TT.A(826) & "!"
    Label_wait.Font.Size = 36
    Label_wait.visible = True
    omax.PrepareNewCommand FindDef:=False
    
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
        text = text & "point_size=0.1," ' fails with 0 on mac
#Else
        text = text & "point_size=0,"
#End If
    End If
    text = text & "point_type=filled_circle,points_joined=" & VBA.LCase$(CheckBox_pointsjoined.Value) & ","
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
    text = Left$(text, Len(text) - 1)
    yAxislabel = Left$(yAxislabel, Len(yAxislabel) - 1)
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
    Label_wait.Caption = TT.A(94)
    Label_wait.Font.Size = 12
    Label_wait.Width = 150
    Label_wait.visible = True
    Image1.Picture = Nothing
slut:

End Sub

Sub InsertDefinitioner()
    ' inserts definitions from the textbox into the maximainputstring
    Dim DefString As String

    omax.InsertKillDef

    DefString = GetDefString

    If Len(DefString) > 0 Then
        'defstring = Replace(defstring, ",", ".")
        'defstring = Replace(defstring, ";", ",")
        'defstring = Replace(defstring, "=", ":")
        If Right$(DefString, 1) = "," Then DefString = Left$(DefString, Len(DefString) - 1)

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
    DefString = Replace(DefString, ListSeparator & ListSeparator, ListSeparator) ' double ;; removed
Loop
DefString = omax.AddDefinition("definer:" & DefString)
GetDefString = DefString
End If
End Function

Sub OpdaterDefinitioner()
' looks for variables in the textboxes and inserts under definitions
    Dim Vars As String
   Dim Var As String, var2 As String
   Dim ea As New ExpressionAnalyser
   Dim arr As Variant
   Dim i As Integer, s As String
   Validate
    
   Vars = Vars & GetTextboxVars(TextBox_eq1, TextBox_varx)
   Vars = Vars & GetTextboxVars(TextBox_eq2, TextBox_varx)
   Vars = Vars & GetTextboxVars(TextBox_eq3, TextBox_varx)
   Vars = Vars & GetTextboxVars(TextBox_eq4, TextBox_varx)
   Vars = Vars & GetTextboxVars(TextBox_eq5, TextBox_varx)
   Vars = Vars & GetTextboxVars(TextBox_eq6, TextBox_varx)
   Vars = Vars & GetTextboxVars(TextBox_eq7, TextBox_varx)
   Vars = Vars & GetTextboxVars(TextBox_eq8, TextBox_varx)
   Vars = Vars & GetTextboxVars(TextBox_eq9, TextBox_varx)
    
   omax.FindVariable Vars, False ' fjerner dobbelte
   Vars = omax.Vars
   Vars = RemoveVar(Vars, TextBox_var1.text)
   Vars = RemoveVar(Vars, TextBox_var2.text)
   Vars = RemoveVar(Vars, TextBox_var3.text)
   Vars = RemoveVar(Vars, TextBox_var4.text)
   Vars = RemoveVar(Vars, TextBox_var5.text)
   Vars = RemoveVar(Vars, TextBox_var6.text)
   Vars = RemoveVar(Vars, TextBox_var7.text)
   Vars = RemoveVar(Vars, TextBox_var8.text)
   Vars = RemoveVar(Vars, TextBox_var9.text)
    
   If Left$(Vars, 1) = ";" Then Vars = Right$(Vars, Len(Vars) - 1)
    
   ea.text = Vars
   Do While Right$(TextBox_definitioner.text, 2) = VbCrLfMac
      TextBox_definitioner.text = Left$(TextBox_definitioner.text, Len(TextBox_definitioner.text) - 2)
   Loop
   arr = Split(TextBox_definitioner.text, VbCrLfMac)
   
   For i = 0 To UBound(arr) ' If variable is included in def, it must be removed
      If arr(i) <> "" Then
         var2 = Split(arr(i), "=")(0)
         If var2 = TextBox_varx.text Then
            arr(i) = ""
         End If
         If arr(i) <> "" Then s = s & arr(i) & VbCrLfMac
      End If
   Next
   Do While Right$(s, 2) = vbCrLf
      s = Left$(s, Len(s) - 2)
   Loop
   TextBox_definitioner.text = s
   
   arr = Split(TextBox_definitioner.text, VbCrLfMac)
   Do
      Var = ea.GetNextListItem(ea.pos)
      Var = Replace(Var, vbCrLf, "")
      For i = 0 To UBound(arr)
         If arr(i) <> "" Then
            var2 = Split(arr(i), "=")(0)
            If var2 = Var Then
               Var = ""
               Exit For
            End If
         End If
      Next
      If Var <> "" Then
         '        If right$(TextBox_definitioner.text, 2) <> vbCrLf Then
         If Len(TextBox_definitioner.text) > 0 Then
            TextBox_definitioner.text = TextBox_definitioner.text & VbCrLfMac
         End If
         TextBox_definitioner.text = TextBox_definitioner.text & Var & "=1"
      End If
   Loop While ea.pos <= Len(ea.text)

    
End Sub
Function GetTextboxVars(tb As TextBox, tbvar As TextBox) As String
    If Len(tb.text) > 0 Then
        omax.Vars = ""
        omax.FindVariable tb.text, False
        omax.Vars = RemoveVar(omax.Vars, tbvar.text)
        If Len(omax.Vars) > 0 Then
            GetTextboxVars = ";" & omax.Vars
        End If
    End If
End Function

Function RemoveVar(text As String, Var As String)
    ' removes var from string
    Dim ea As New ExpressionAnalyser
    If Var = vbNullString Then
        RemoveVar = text
        Exit Function
    End If
    ea.text = text
    Call ea.ReplaceVar(Var, "")
    text = Replace(ea.text, ";;", ";")
    If Left$(text, 1) = ";" Then text = Right$(text, Len(text) - 1)
    If Right$(text, 1) = ";" Then text = Left$(text, Len(text) - 1)

    RemoveVar = text
End Function

Sub SetCaptions()
    Me.Caption = TT.A(85)
    Label6.Caption = TT.A(86)
    Label7.Caption = TT.A(87)
    Label_Graf.Caption = TT.A(667)
    Label_opdater.Caption = TT.A(461)
    Label_cancel.Caption = TT.A(661)
    Label_var.Caption = TT.A(746)
    Label3.Caption = TT.A(88)
    Label5.Caption = TT.A(823)
    Label_wait.Caption = TT.A(826) & "!"
    CheckBox_pointsjoined.Caption = TT.A(89)
    CheckBox_visforklaring.Caption = TT.A(90)
    Label_tolist.Caption = TT.A(91)
    Label_inserttabel.Caption = TT.A(92)
    Label_insertgraph.Caption = TT.A(93)
    Label_toExcel.Caption = ChrW$(&H2192) & " Excel"
    
#If Mac Then
    ComboBox_graphapp.visible = False
#Else
    ComboBox_graphapp.Clear
    ComboBox_graphapp.AddItem "GnuPlot"
    ComboBox_graphapp.AddItem "GeoGebra"
    If GraphApp = 0 Then
       ComboBox_graphapp.ListIndex = 0
    Else
       ComboBox_graphapp.ListIndex = 1
    End If
#End If
End Sub
Sub ShowPreviewMac()
#If Mac Then
    RunScript "OpenPreview", GetTempDir() & "WordMatGraf.pdf"
#End If
End Sub

Private Sub Label_opdater_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_opdater.BackColor = LBColorPress
End Sub
Private Sub Label_opdater_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetLabelsInactive
    Label_opdater.BackColor = LBColorHover
End Sub
Private Sub Label_cancel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorPress
End Sub
Private Sub Label_cancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetLabelsInactive
    Label_cancel.BackColor = LBColorHover
End Sub
Private Sub Label_insertgraph_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_insertgraph.BackColor = LBColorPress
End Sub
Private Sub Label_insertgraph_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetLabelsInactive
    Label_insertgraph.BackColor = LBColorHover
End Sub
Private Sub Label_tolist_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_tolist.BackColor = LBColorPress
End Sub
Private Sub Label_tolist_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetLabelsInactive
    Label_tolist.BackColor = LBColorHover
End Sub
Private Sub Label_toExcel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_toExcel.BackColor = LBColorPress
End Sub
Private Sub Label_toExcel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetLabelsInactive
    Label_toExcel.BackColor = LBColorHover
End Sub
Private Sub Label_inserttabel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_inserttabel.BackColor = LBColorPress
End Sub
Private Sub Label_inserttabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetLabelsInactive
    Label_inserttabel.BackColor = LBColorHover
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetLabelsInactive
End Sub

Sub SetLabelsInactive()
    Label_opdater.BackColor = LBColorInactive
    Label_cancel.BackColor = LBColorInactive
    Label_insertgraph.BackColor = LBColorInactive
    Label_tolist.BackColor = LBColorInactive
    Label_inserttabel.BackColor = LBColorInactive
    Label_toExcel.BackColor = LBColorInactive
End Sub
