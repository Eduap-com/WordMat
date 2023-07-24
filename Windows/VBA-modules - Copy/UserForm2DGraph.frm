VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2DGraph 
   Caption         =   "Plot af grafer og punkter i planen"
   ClientHeight    =   7170
   ClientLeft      =   -30
   ClientTop       =   45
   ClientWidth     =   15945
   OleObjectBlob   =   "UserForm2DGraph.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2DGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private gemx As Single
Private gemy As Single
Private embed As Boolean
Private etikettext As String
Private nytpunkt As Boolean
Private nytmarkerpunkt As Boolean
Private DisableEvents As Boolean
Private Opened As Boolean

Private Sub CommandButton_cancel_Click()
    On Error Resume Next
    PicOpen = False
#If Mac Then
#Else
    If MaxProc.Finished = 0 Then
        MaxProc.CloseProcess
        MaxProc.StartMaximaProcess
    End If
#End If
    Unload Me
End Sub

Private Sub Label_helpmarker_Click()
MsgBox Sprog.A(195), vbOKOnly, Sprog.Help
End Sub

Private Sub Label_punkter2_Click()
MsgBox Sprog.A(196), vbOKOnly, Sprog.Help
End Sub

Private Sub Label_symbol_Click()
Dim ctrl As control
On Error GoTo Fejl
Set ctrl = Me.ActiveControl
If Left(ctrl.Name, 7) <> "TextBox" Then Set ctrl = TextBox_titel
UserFormSymbol.Show
ctrl.Text = ctrl.Text & UserFormSymbol.tegn
Fejl:
End Sub


Private Sub TextBox_definitioner_AfterUpdate()
    CheckForAssume

End Sub

Private Sub TextBox_xmin_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If ToggleButton_propto.Value Then
        TextBox_ymin.Text = TextBox_xmin.Text
    End If
End Sub
Private Sub TextBox_xmax_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If ToggleButton_propto.Value Then
        TextBox_ymax.Text = TextBox_xmax.Text
    End If
End Sub
Private Sub ToggleButton_auto_Click()
If DisableEvents Then Exit Sub
DisableEvents = True
ToggleButton_propto.Value = False
ToggleButton_manuel.Value = False
TextBox_ymin.Text = ""
TextBox_ymax.Text = ""
DisableEvents = False
OpdaterGraf
Me.Repaint
End Sub

Private Sub ToggleButton_manuel_Click()
If DisableEvents Then Exit Sub
DisableEvents = True
ToggleButton_auto.Value = False
ToggleButton_propto.Value = False
If TextBox_ymin.Text = "" Then
    TextBox_ymin.Text = TextBox_xmin.Text
End If
If TextBox_ymax.Text = "" Then
    TextBox_ymax.Text = TextBox_xmax.Text
End If
DisableEvents = False
End Sub

Private Sub ToggleButton_propto_Click()
If DisableEvents Then Exit Sub
DisableEvents = True
ToggleButton_auto.Value = False
ToggleButton_manuel.Value = False
If TextBox_ymin.Text = "" Then
    TextBox_ymin.Text = TextBox_xmin.Text
Else
End If
TextBox_ymax.Text = TextBox_ymin.Text + (TextBox_xmax.Text - TextBox_xmin.Text) * 23 / 33

DisableEvents = False
OpdaterGraf
Me.Repaint
End Sub

Private Sub UserForm_Activate()
Dim d As String, nd As String
Dim ea As New ExpressionAnalyser
Dim s As String
Dim xmin As String, xmax As String
Dim Arr As Variant, i As Integer
On Error Resume Next
    SetCaptions
#If Mac Then
    If Opened Then Exit Sub
    Opened = True
    Me.Left = 10
    Me.top = 80
    Label_wait.Left = 180
    Label_wait.top = 270
    Label_zoom.visible = False
    Kill GetTempDir() & "WordMatGraf.pdf"
#Else
    Kill GetTempDir() & "\WordMatGraf.gif"
#End If

If Not PicOpen Then
    omax.PrepareNewCommand '
    If Len(omax.DefString) > 1 Then
    d = omax.defstringtext

    d = Replace(d, "assume", "")
    d = Replace(d, ":=", "=")
    d = Replace(d, ":", "=")
'    d = omax.ConvertToAscii(omax.ConvertToWordSymbols(d)) ' fjernet efter defstringtext anvendes
    d = Trim(d)
    
    ' reverse definition order
    Arr = Split(d, "$")
    nd = Arr(0)
    For i = 1 To UBound(Arr)
        If Len(Arr(i)) > 0 Then nd = Trim(Arr(i)) & VbCrLfMac & nd
    Next
    
'    d = Replace(d, "$", vbCrLf)
    
    TextBox_definitioner.Text = nd
    
    End If
End If
    OpdaterDefinitioner
    CheckForAssume
    ' indsæt xmin og xmax hvis der var definerede
    xmin = TextBox_xmin1.Text
    If Len(TextBox_xmin2.Text) And ConvertStringToNumber(TextBox_xmin2.Text) < ConvertStringToNumber(xmin) Then xmin = TextBox_xmin2.Text
    If Len(TextBox_xmin3.Text) And ConvertStringToNumber(TextBox_xmin3.Text) < ConvertStringToNumber(xmin) Then xmin = TextBox_xmin3.Text
    If Len(TextBox_xmin4.Text) And ConvertStringToNumber(TextBox_xmin4.Text) < ConvertStringToNumber(xmin) Then xmin = TextBox_xmin4.Text
    If Len(TextBox_xmin5.Text) And ConvertStringToNumber(TextBox_xmin5.Text) < ConvertStringToNumber(xmin) Then xmin = TextBox_xmin5.Text
    If Len(TextBox_xmin6.Text) And ConvertStringToNumber(TextBox_xmin6.Text) < ConvertStringToNumber(xmin) Then xmin = TextBox_xmin6.Text
    xmax = TextBox_xmax1.Text
    If Len(TextBox_xmax2.Text) And ConvertStringToNumber(TextBox_xmax2.Text) > ConvertStringToNumber(xmax) Then xmax = TextBox_xmax2.Text
    If Len(TextBox_xmax3.Text) And ConvertStringToNumber(TextBox_xmax3.Text) > ConvertStringToNumber(xmax) Then xmax = TextBox_xmax3.Text
    If Len(TextBox_xmax4.Text) And ConvertStringToNumber(TextBox_xmax4.Text) > ConvertStringToNumber(xmax) Then xmax = TextBox_xmax4.Text
    If Len(TextBox_xmax5.Text) And ConvertStringToNumber(TextBox_xmax5.Text) > ConvertStringToNumber(xmax) Then xmax = TextBox_xmax5.Text
    If Len(TextBox_xmax6.Text) And ConvertStringToNumber(TextBox_xmax6.Text) > ConvertStringToNumber(xmax) Then xmax = TextBox_xmax6.Text
    
    If Len(xmin) > 0 Then TextBox_xmin.Text = xmin
    If Len(xmax) > 0 Then TextBox_xmax.Text = xmax
    If Len(TextBox_xmin.Text) > 0 And Len(TextBox_xmax.Text) > 0 Then
        If ConvertStringToNumber(TextBox_xmin.Text) > ConvertStringToNumber(TextBox_xmax.Text) Then
            TextBox_xmax.Text = ConvertNumberToString(ConvertStringToNumber(TextBox_xmin.Text) + 10)
        End If
    End If
'    TextBox_xmin.text = "-5"
'    TextBox_xmax.text = "5"
    
    OpdaterGraf
    
    TextBox_ligning1.SetFocus
End Sub

Private Sub UserForm_Initialize()
#If Mac Then
    Me.Width = 300
'    Me.Left = 50
'    Me.Top = 50
#End If
    FillLineStyleCombos
    Label_symbol.Caption = VBA.ChrW(937)

End Sub

Private Sub CommandButton_insertmarkerpunkt_Click()
nytmarkerpunkt = True
End Sub

Private Sub CommandButton_insertplan_Click()
Dim linje As String
'    plan = "a*(x-x0)+b*(y-y0)=0"
    linje = "1*(x-0)+1*(y-0)=0"
    If TextBox_lig1.Text = "" Then
        TextBox_lig1.Text = linje
    ElseIf TextBox_lig2.Text = "" Then
        TextBox_lig2.Text = linje
    ElseIf TextBox_Lig3.Text = "" Then
        TextBox_Lig3.Text = linje
    End If

End Sub

Private Sub CommandButton_kugle_Click()
Dim cirkel As String
    cirkel = "(x-0)^2+(y-0)^2=1^2"
    If TextBox_lig1.Text = "" Then
        TextBox_lig1.Text = cirkel
    ElseIf TextBox_lig2.Text = "" Then
        TextBox_lig2.Text = cirkel
    ElseIf TextBox_Lig3.Text = "" Then
        TextBox_Lig3.Text = cirkel
    End If

End Sub

Private Sub CommandButton_nulstillign1_Click()
TextBox_lig1.Text = ""
End Sub

Private Sub CommandButton_nulstillign3_Click()
TextBox_Lig3.Text = ""
End Sub

Private Sub CommandButton_nulstilligning2_Click()
TextBox_lig2.Text = ""
End Sub

Private Sub CommandButton_nulstilpar1_Click()
TextBox_parametric1x.Text = ""
TextBox_parametric1y.Text = ""
TextBox_tmin1.Text = ""
TextBox_tmax1.Text = ""

End Sub

Private Sub CommandButton_nulstilpar2_Click()
TextBox_parametric2x.Text = ""
TextBox_parametric2y.Text = ""
TextBox_tmin2.Text = ""
TextBox_tmax2.Text = ""

End Sub

Private Sub CommandButton_nulstilpar3_Click()
TextBox_parametric3x.Text = ""
TextBox_parametric3y.Text = ""
TextBox_tmin3.Text = ""
TextBox_tmax3.Text = ""

End Sub

Private Sub CommandButton_nulstilvektorer_Click()
    TextBox_vektorer.Text = ""
End Sub

Private Sub CommandButton_nyetiket_Click()

    etikettext = InputBox(Sprog.A(299), Sprog.A(298), "")
    
End Sub

Private Sub CommandButton_nytpunkt_Click()
    nytpunkt = True
End Sub

Private Sub CommandButton_nyvektor_Click()
    If TextBox_vektorer.Text <> "" Then
        TextBox_vektorer.Text = TextBox_vektorer.Text & VbCrLfMac
    End If
    TextBox_vektorer.Text = TextBox_vektorer.Text & "(0;0)-(1;1)"

End Sub

Private Sub CommandButton_parlinje_Click()
Dim px As String
Dim py As String
px = "0+1*t"
py = "0+1*t"

If TextBox_parametric1x.Text = "" Then
    TextBox_parametric1x.Text = px
    TextBox_parametric1y.Text = py
    TextBox_tmin1.Text = "0"
    TextBox_tmax1.Text = "1"
ElseIf TextBox_parametric2x.Text = "" Then
    TextBox_parametric2x.Text = px
    TextBox_parametric2y.Text = py
    TextBox_tmin2.Text = "0"
    TextBox_tmax2.Text = "1"
ElseIf TextBox_parametric3x.Text = "" Then
    TextBox_parametric3x.Text = px
    TextBox_parametric3y.Text = py
    TextBox_tmin3.Text = "0"
    TextBox_tmax3.Text = "1"
End If


End Sub

Private Sub TextBox_ligning1_AfterUpdate()
    OpdaterDefinitioner
End Sub
Private Sub TextBox_ligning2_AfterUpdate()
    OpdaterDefinitioner
End Sub
Private Sub TextBox_ligning3_AfterUpdate()
    OpdaterDefinitioner
End Sub
Private Sub TextBox_ligning4_AfterUpdate()
    OpdaterDefinitioner
End Sub
Private Sub TextBox_ligning5_AfterUpdate()
    OpdaterDefinitioner
End Sub
Private Sub TextBox_ligning6_AfterUpdate()
    OpdaterDefinitioner
End Sub
Private Sub TextBox_lig1_AfterUpdate()
    OpdaterDefinitioner
End Sub
Private Sub TextBox_lig2_AfterUpdate()
    OpdaterDefinitioner
End Sub
Private Sub TextBox_lig3_AfterUpdate()
    OpdaterDefinitioner
End Sub

Private Sub CommandButton_excelindlejret_Click()
    embed = True
    ExcelPlot
    Unload Me

End Sub

Private Sub CommandButton_excelopen_Click()
    embed = False
    ExcelPlot
    Unload Me
End Sub

Private Sub CommandButton_opdater_Click()
OpdaterGraf
#If Mac Then
    ShowPreviewMac
#Else
    Me.Repaint
#End If
End Sub

Private Sub Label_punkterhelp_Click()
    MsgBox Sprog.A(197), vbOKOnly, Sprog.Help
End Sub

Private Sub CommandButton_nulstil1_Click()
    TextBox_ligning1.Text = ""
    TextBox_var1.Text = "x"
    TextBox_xmin1.Text = ""
    TextBox_xmax1.Text = ""
    ComboBox_ligning1.Text = ""
End Sub
Private Sub CommandButton_nulstil2_Click()
    TextBox_ligning2.Text = ""
    TextBox_var2.Text = "x"
    TextBox_xmin2.Text = ""
    TextBox_xmax2.Text = ""
    ComboBox_ligning2.Text = ""
End Sub
Private Sub CommandButton_nulstil3_Click()
    TextBox_ligning3.Text = ""
    TextBox_var3.Text = "x"
    TextBox_xmin3.Text = ""
    TextBox_xmax3.Text = ""
    ComboBox_ligning3.Text = ""
End Sub
Private Sub CommandButton_nulstil4_Click()
    TextBox_ligning4.Text = ""
    TextBox_var4.Text = "x"
    TextBox_xmin4.Text = ""
    TextBox_xmax4.Text = ""
    ComboBox_ligning4.Text = ""
End Sub
Private Sub CommandButton_nulstil5_Click()
    TextBox_ligning5.Text = ""
    TextBox_var5.Text = "x"
    TextBox_xmin5.Text = ""
    TextBox_xmax5.Text = ""
    ComboBox_ligning5.Text = ""
End Sub
Private Sub CommandButton_nulstil6_Click()
    TextBox_ligning6.Text = ""
    TextBox_var6.Text = "x"
    TextBox_xmin6.Text = ""
    TextBox_xmax6.Text = ""
    ComboBox_ligning6.Text = ""
End Sub


Sub FillLineStyleCombos()
ComboBox_ligning1.Clear
ComboBox_ligning1.AddItem ("---")
ComboBox_ligning1.AddItem ("...")
'ComboBox_ligning1.AddItem ("- - -")
'ComboBox_ligning1.AddItem ("-.-.-.")
'ComboBox_ligning1.AddItem ("- . . - . .")

ComboBox_ligning2.Clear
ComboBox_ligning2.AddItem ("---")
ComboBox_ligning2.AddItem ("...")
'ComboBox_ligning2.AddItem ("- - -")
'ComboBox_ligning2.AddItem ("-.-.-.")
'ComboBox_ligning2.AddItem ("- . . - . .")

ComboBox_ligning3.Clear
ComboBox_ligning3.AddItem ("---")
ComboBox_ligning3.AddItem ("...")
'ComboBox_ligning3.AddItem ("- - -")
'ComboBox_ligning3.AddItem ("-.-.-.")
'ComboBox_ligning3.AddItem ("- . . - . .")

ComboBox_ligning4.Clear
ComboBox_ligning4.AddItem ("---")
ComboBox_ligning4.AddItem ("...")
'ComboBox_ligning4.AddItem ("- - -")
'ComboBox_ligning4.AddItem ("-.-.-.")
'ComboBox_ligning4.AddItem ("- . . - . .")

ComboBox_ligning5.Clear
ComboBox_ligning5.AddItem ("---")
ComboBox_ligning5.AddItem ("...")
'ComboBox_ligning5.AddItem ("- - -")
'ComboBox_ligning5.AddItem ("-.-.-.")
'ComboBox_ligning5.AddItem ("- . . - . .")

ComboBox_ligning6.Clear
ComboBox_ligning6.AddItem ("---")
ComboBox_ligning6.AddItem ("...")
'ComboBox_ligning6.AddItem ("- - -")
'ComboBox_ligning6.AddItem ("-.-.-.")
'ComboBox_ligning6.AddItem ("- . . - . .")

End Sub
Private Sub CommandButton_ok_Click()
        GnuPlot
End Sub

Sub ExcelPlot()
'Dim xcl As New CExcel
'Dim ws As Worksheet
'Dim ws As excel.Worksheet
'Dim ws As Worksheet
'Dim wb As Workbook
Dim ws As Object
Dim WB As Object

Dim path As String
Dim ils As InlineShape
Dim tb As TextBox
Dim forskrift As String
Dim xmin As Double, xmax As Double
Dim plinjer As Variant
Dim linje As Variant
Dim i As Integer

If cxl Is Nothing Then Set cxl = New CExcel
Application.ScreenUpdating = False
Me.Hide
    Dim UfWait2 As New UserFormWaitForMaxima
    UfWait2.Show vbModeless
    DoEvents
    UfWait2.Label_progress = "***"


If Not embed Then
cxl.LoadFile ("Graphs.xltm")
Set WB = cxl.xlwb
Set ws = cxl.xlwb.Sheets("Tabel")
    UfWait2.Label_progress = UfWait2.Label_progress & "***"

'Dim excl As Object

'Set excl = CreateObject("Excel.Application")
'Set ws = excl.ActiveWorkbook.Sheets("Tabel")


Else

path = """" & GetProgramFilesDir & "\WordMat\ExcelFiles\Graphs.xltm"""
PrepareMaximaNoSplash
omax.GoToEndOfSelectedMaths
'Selection.Collapse wdCollapseEnd
Selection.TypeParagraph
    UfWait2.Label_progress = UfWait2.Label_progress & "**"

EnableExcelMacros

Set ils = ActiveDocument.InlineShapes.AddOLEObject(FileName:=path, LinkToFile:=False, _
DisplayAsIcon:=False, Range:=Selection.Range)

'Ils.Height = 300
'Ils.Width = 500
    UfWait2.Label_progress = UfWait2.Label_progress & "***********"


'Ils.OLEFormat.DoVerb (wdOLEVerbOpen)
ils.OLEFormat.DoVerb (wdOLEVerbShow)
Set WB = ils.OLEFormat.Object
Set ws = WB.Sheets("Tabel")
'ws.Activate
DisableExcelMacros
End If

XLapp.Application.EnableEvents = False
XLapp.Application.ScreenUpdating = False
'excel.Application.EnableEvents = False
'excel.Application.ScreenUpdating = False

    UfWait2.Label_progress = UfWait2.Label_progress & "*****"
xmin = val(TextBox_xmin.Text)
xmax = val(TextBox_xmax.Text)
If xmin < xmax Then
    ws.Range("n3").Value = Me.TextBox_xmin.Text
    ws.Range("o3").Value = Me.TextBox_xmax.Text
Else
    ws.Range("n3").Value = -5
    ws.Range("o3").Value = 5
End If

    
ws.Range("b4").Value = Me.TextBox_ligning1.Text
ws.Range("c4").Value = Me.TextBox_ligning2.Text
ws.Range("d4").Value = Me.TextBox_ligning3.Text
ws.Range("e4").Value = Me.TextBox_ligning4.Text
ws.Range("f4").Value = Me.TextBox_ligning5.Text
ws.Range("g4").Value = Me.TextBox_ligning6.Text
'xmin og xmax kopieres over
ws.Range("B2").Value = Me.TextBox_xmin1.Text
ws.Range("B3").Value = Me.TextBox_xmax1.Text
ws.Range("C2").Value = Me.TextBox_xmin2.Text
ws.Range("C3").Value = Me.TextBox_xmax2.Text
ws.Range("D2").Value = Me.TextBox_xmin3.Text
ws.Range("D3").Value = Me.TextBox_xmax3.Text
ws.Range("E2").Value = Me.TextBox_xmin4.Text
ws.Range("E3").Value = Me.TextBox_xmax4.Text
ws.Range("F2").Value = Me.TextBox_xmin5.Text
ws.Range("F3").Value = Me.TextBox_xmax5.Text
ws.Range("G2").Value = Me.TextBox_xmin6.Text
ws.Range("G3").Value = Me.TextBox_xmax6.Text
'variabelnavn kopieres over
ws.Range("B1").Value = Me.TextBox_var1.Text
ws.Range("C1").Value = Me.TextBox_var2.Text
ws.Range("D1").Value = Me.TextBox_var3.Text
ws.Range("E1").Value = Me.TextBox_var4.Text
ws.Range("F1").Value = Me.TextBox_var5.Text
ws.Range("G1").Value = Me.TextBox_var6.Text
' indstillinger
If Radians Then
    ws.Range("A4").Value = "rad"
Else
    ws.Range("A4").Value = "grad"
End If

'If TextBox_ligning1.text <> "" Then Call InsertFormula(ws, wb, TextBox_ligning1, 0)
'If TextBox_ligning2.text <> "" Then Call InsertFormula(ws, wb, TextBox_ligning2, 1)
'If TextBox_ligning3.text <> "" Then Call InsertFormula(ws, wb, TextBox_ligning3, 2)
'If TextBox_ligning4.text <> "" Then Call InsertFormula(ws, wb, TextBox_ligning4, 3)
'If TextBox_ligning5.text <> "" Then Call InsertFormula(ws, wb, TextBox_ligning5, 4)
'If TextBox_ligning6.text <> "" Then Call InsertFormula(ws, wb, TextBox_ligning6, 5)

On Error GoTo Slut
'If TextBox_ligning1.text <> "" Then Call SetLineStyle(ComboBox_ligning1, 1)
'If TextBox_ligning2.text <> "" Then Call SetLineStyle(ComboBox_ligning2, 2)
'If TextBox_ligning3.text <> "" Then Call SetLineStyle(ComboBox_ligning3, 3)
'If TextBox_ligning4.text <> "" Then Call SetLineStyle(ComboBox_ligning4, 4)
'If TextBox_ligning5.text <> "" Then Call SetLineStyle(ComboBox_ligning5, 5)
'If TextBox_ligning6.text <> "" Then Call SetLineStyle(ComboBox_ligning6, 6)

'ws.ChartObjects(1).SeriesCollection("f1").ChartType = xlXYScatterSmoothNoMarkers

'ActiveChart.SeriesCollection("f1").Border.LineStyle = xlDashDotDot 'xlContinuous '

'tb = ws.ChartObjects(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 30.75, 11.25, 57.75, 21#)
'tb.text = "hej"

    'datapunkter
    If TextBox_punkter.Text <> "" Then
        Dim punkttekst As String, Sep As String
        punkttekst = Me.TextBox_punkter.Text
        plinjer = Split(punkttekst, VbCrLfMac)
        For i = 0 To UBound(plinjer)
            If InStr(plinjer(i), ";") > 0 Then
                Sep = ";"
            Else
                Sep = ","
            End If
            linje = Split(plinjer(i), Sep)
            If UBound(linje) > 0 Then
                ws.Range("H7").Offset(i, 0).Value = Replace(linje(0), ",", ".")
                ws.Range("I7").Offset(i, 0).Value = Replace(linje(1), ",", ".")
            End If
        Next
    End If


Slut:
On Error GoTo slut2
    ws.Range("p3").Value = Me.TextBox_ymin.Text
    ws.Range("q3").Value = Me.TextBox_ymax.Text
'wb.Charts(1).Activate
If TextBox_xaksetitel.Text <> "" Then
    WB.Charts(1).Axes(xlCategory, xlPrimary).AxisTitle.Text = Me.TextBox_xaksetitel.Text
    ws.ChartObjects(1).Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = Me.TextBox_xaksetitel.Text
End If
If TextBox_yaksetitel.Text <> "" Then
    WB.Charts(1).Axes(xlValue, xlPrimary).AxisTitle.Text = Me.TextBox_yaksetitel.Text
    ws.ChartObjects(1).Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = Me.TextBox_yaksetitel.Text
End If
    UfWait2.Label_progress = UfWait2.Label_progress & "**"

'excel.Run ("UpDateAll")
XLapp.Run ("UpDateAll")
    
    UfWait2.Label_progress = UfWait2.Label_progress & "***"
WB.Charts(1).Activate
'excel.Application.EnableEvents = True
'excel.Application.ScreenUpdating = True
XLapp.Application.EnableEvents = True
XLapp.Application.ScreenUpdating = True
slut2:
    Unload UfWait2

'Excel.Application.ScreenUpdating = True

'excel.ActiveSheet.ChartObjects(1).Copy
'Selection.Collapse Direction:=wdCollapseStart
'Selection.Paste
'Selection.PasteSpecial DataType:=wdPasteOLEObject
'Selection.PasteSpecial DataType:=wdPasteShape
End Sub
Sub SetLineStyle(cb As ComboBox, n As Integer)
' sætter linestyle efter hvad comboxen er sat til

If cb.ListIndex = 0 Then
XLapp.ActiveChart.SeriesCollection(n).Border.LineStyle = xlContinuous '
ElseIf cb.ListIndex = 1 Then
XLapp.ActiveChart.SeriesCollection(n).Border.LineStyle = xlDot 'xlContinuous '
ElseIf cb.ListIndex = 2 Then
XLapp.ActiveChart.SeriesCollection(n).Border.LineStyle = xlDash 'xlContinuous '
ElseIf cb.ListIndex = 3 Then
XLapp.ActiveChart.SeriesCollection(n).Border.LineStyle = xlDashDot 'xlContinuous '
ElseIf cb.ListIndex = 4 Then
XLapp.ActiveChart.SeriesCollection(n).Border.LineStyle = xlDashDotDot 'xlContinuous '
Else
XLapp.ActiveChart.SeriesCollection(n).Border.LineStyle = xlContinuous '
End If

End Sub
'Sub InsertFormula(ws As Worksheet, wb As Workbook, tb As TextBox, col As Integer)
Sub InsertFormula(ws As Variant, WB As Variant, tb As TextBox, col As Integer)
' indsætter formel fra textbox i kolonne col
Dim ea As New ExpressionAnalyser
    Dim varnavn As String
    Dim i As Integer
    Dim j As Integer
    Dim forskrift As String

If tb.Text <> "" Then
    forskrift = tb.Text
    forskrift = ConvertToExcelFormula(forskrift)
    ea.Text = forskrift
    forskrift = ea.Text
    
    ws.Range("b4").Offset(0, col).Value = forskrift
    
    
    'find variable
    ea.Text = forskrift
    ea.Pos = 1
    varnavn = ea.GetNextVar
    i = 0
    ' find ledig variabel plads
    Do While ws.Range("N6").Offset(i, 0).Value <> ""
      i = i + 1
    Loop
    On Error Resume Next
    Do While varnavn <> ""
'        If VBA.LCase(varnavn) = "c" Then
'            varnavn = InputBox("Du har angivet et variabelnavn der ikke kan bruges i Excel. Angiv et andet.", "Fejl", varnavn)
'            forskrift = Left(forskrift, ea.pos - 1) & varnavn & Right(forskrift, Len(forskrift) - ea.pos)
'            ws.Range("b4").Offset(0, col).value = forskrift
'        End If

        If varnavn <> "x" And Left(varnavn, 4) <> "matm" And Not (ea.IsFunction(varnavn)) Then
        Call ea.ReplaceVar(varnavn, "matm" & varnavn)
        ea.Pos = ea.Pos + Len(varnavn) + 4
        j = 0
        Do While ws.Range("N6").Offset(j, 0).Value <> varnavn & "=" And j < i ' check om findes
          j = j + 1
        Loop
        If j = i Then
        ws.Range("N6").Offset(i, 0).Value = varnavn & "="
        ws.Range("N6").Offset(i, 1).Value = 1
        WB.Names.Add Name:="matm" & varnavn, RefersToR1C1:="=Tabel!R" & i + 6 & "C15"
        i = i + 1
        End If
        Else
            ea.Pos = ea.Pos + Len(varnavn)
        End If
        varnavn = ea.GetNextVar
    Loop

    On Error GoTo fejlindtast
    ea.Pos = 1
    Call ea.ReplaceVar("x", "A7")
    forskrift = ea.Text
     
     ' indsæt forskrift i regneark
    ws.Range("b7").Offset(0, col).Formula = "=" & forskrift
'    If TypeName(ws.Range("b7").Offset(0, col).value) = "Error" Then GoTo fejlindtast

    ' kopier formel ned
    ws.Activate
    ws.Range("B7").Offset(0, col).Select
'    ws.Application.Selection.AutoFill Destination:=ws.Range("b7:b207").Offset(0, col), Type:=xlFillDefault
    ws.Application.Selection.AutoFill Destination:=ws.Range("b7:b207").Offset(0, col), Type:=0   'xlFillDefault=0
    
    ' fejl nogen steder?
    For i = 0 To 200
        If TypeName(ws.Range("b7").Offset(i, col).Value) = "Error" Then ws.Range("b7").Offset(i, col).Value = ""
    Next
        
End If
GoTo Slut:
fejlindtast:
    MsgBox Sprog.A(300) & " " & col + 1
Slut:

End Sub

Function ConvertToExcelFormula(ByVal forskrift As String)
Dim ea As New ExpressionAnalyser
Dim Pos As Integer
Dim posb As Integer
Dim rod As Integer
Dim pos2 As Integer
Dim pos4 As Integer
Dim pos5 As Integer
Dim posog As Integer
Dim Arr As Variant

    ea.SetNormalBrackets

'      ws.Range("b7").Replace What:="x", Replacement:="A7", SearchOrder:=xlByColumns, MatchCase:=True
    
    forskrift = Replace(forskrift, "pi", "PI()")
    forskrift = Replace(forskrift, VBA.ChrW(960), "PI()") ' pi symbol
    forskrift = Replace(forskrift, "e", "2.718281828")
    forskrift = Replace(forskrift, VBA.ChrW(12310), "") ' specielle usynlige paranteser fjernes
    forskrift = Replace(forskrift, VBA.ChrW(12311), "") ' specielle usynlige paranteser fjernes
'    forskrift = Replace(forskrift, VBA.ChrW(11), "") ' en af de nedenstående?
    forskrift = Replace(forskrift, vbLf, "") ' shift-enter og enter
    forskrift = Replace(forskrift, vbCrLf, "")
    forskrift = Replace(forskrift, vbCr, "")
    forskrift = Replace(forskrift, """", "") ' apostrof fjernes
    forskrift = Replace(forskrift, VBA.ChrW(8289), "") ' symbol der definerer funktion fjernes fra word syntaks
    forskrift = Replace(forskrift, VBA.ChrW(8212), "+") 'dobbbelt minustegn giver plus
    forskrift = Replace(forskrift, VBA.ChrW(183), "*") ' prik erstattes med gange
    forskrift = Replace(forskrift, VBA.ChrW(215), "*") ' kryds erstattes med gange
    forskrift = Replace(forskrift, VBA.ChrW(8901), "*") ' prik \cdot erstattes med gange
    forskrift = Replace(forskrift, VBA.ChrW(8226), "*") ' tyk prik erstattes med gange
    forskrift = Replace(forskrift, "%", "/100") ' procenttegn
    forskrift = Replace(forskrift, ",", ".")
    forskrift = Replace(forskrift, VBA.ChrW(178), "^2")
    forskrift = Replace(forskrift, VBA.ChrW(179), "^3")
    
    forskrift = Replace(forskrift, "cos^(-1)", "ARCCOS")
    forskrift = Replace(forskrift, "sin^(-1)", "ARCSIN")
    forskrift = Replace(forskrift, "tan^(-1)", "ARCTAN")
      
    Do
    Pos = InStr(forskrift, VBA.ChrW(124))
    If Pos > 0 Then
        posb = InStr(Pos + 1, forskrift, VBA.ChrW(124))
        forskrift = Left(forskrift, Pos - 1) & "abs(" & Mid(forskrift, Pos + 1, posb - Pos - 1) & ")" & right(forskrift, Len(forskrift) - posb)
    End If
    Loop While Pos > 0
    
    ' 3 og 4 rod
    For rod = 3 To 4
    Do
    Pos = InStr(forskrift, VBA.ChrW(8728 + rod))
    If Pos > 0 Or pos4 > 0 Or pos5 > 0 Then
        ea.Text = forskrift
        ea.Pos = Pos + 1
        If Mid(forskrift, Pos + 1, 1) <> "(" Then
            ea.InsertUnderstoodBracketPair
        End If
        ea.Pos = Pos
        Call ea.GetNextBracketContent ' bare for at finde slut parantes
        Call ea.InsertBeforePos("^(1/" & rod & ")")
        ea.Text = Replace(ea.Text, VBA.ChrW(8728 + rod), "", 1, 1)
        forskrift = ea.Text
    End If
    Loop While Pos > 0
    Next
    
    'kvadratrod
    Do
    Pos = InStr(forskrift, VBA.ChrW(8730))
    If Pos > 0 Then
        If Mid(forskrift, Pos + 1, 1) <> "(" Then
            forskrift = Replace(forskrift, VBA.ChrW(8730), "sqrt", 1, 1)
            Pos = Pos + 4
            ea.Text = forskrift
            ea.Pos = Pos
            ea.InsertUnderstoodBracketPair
            forskrift = ea.Text
        Else
            ea.Text = forskrift
            ea.Pos = Pos
            Arr = Split(ea.GetNextBracketContent, "&")
            pos2 = ea.Pos
            If UBound(Arr) = 0 Then
                forskrift = Replace(forskrift, VBA.ChrW(8730), "sqrt", 1, 1)
            ElseIf UBound(Arr) = 1 Then
                rod = Arr(0)
                Call ea.InsertBeforePos("^(1/(" & rod & "))")
                ea.Text = Replace(ea.Text, VBA.ChrW(8730), "", 1, 1)
                posog = ea.FindChr("&", 1)
                forskrift = Left(ea.Text, Pos) & right(ea.Text, Len(ea.Text) - posog)
               
            End If
        End If
    End If
    Loop While Pos > 0
    
    
    'trigfunktioner hvis 360 grader
    If Not (Radians) Then
        forskrift = ConvertDegreeToRad(forskrift, "sin")
        forskrift = ConvertDegreeToRad(forskrift, "cos")
        forskrift = ConvertDegreeToRad(forskrift, "tan")
        forskrift = ConvertDegreeToRad(forskrift, "sec")
        forskrift = ConvertDegreeToRad(forskrift, "cot")
        forskrift = ConvertDegreeToRad(forskrift, "csc")
    End If
    
    ' find underforståede paranteser efter ^ og / ' (skal være efter diff og andre funktioner med komma)
    ea.Text = forskrift
    ea.InsertBracketAfter ("^")
    ea.InsertBracketAfter ("/")
    forskrift = ea.Text
    
    ' mellemrum fjernes
    forskrift = Replace(forskrift, " ", "")

    ' indsæt underforståede gangetegn ' skal være efter fjern mellem
    ea.Text = forskrift
    ea.Pos = 1
    ea.InsertMultSigns
    forskrift = ea.Text
    
    ConvertToExcelFormula = forskrift

End Function
Function GetDraw2Dtext(Optional highres As Double = 1) As String
On Error GoTo Fejl
Dim grafobj As String
Dim xmin As String
Dim xmax As String
Dim ymin As String
Dim ymax As String
Dim ymin2 As String
Dim ymax2 As String
Dim xming As String
Dim xmaxg As String
Dim yming As String
Dim ymaxg As String
Dim labeltext As String
Dim lign As String
Dim punkttekst As String
Dim Arr As Variant
Dim arr2 As Variant
Dim i As Integer
Dim vekt As String
Dim parx As String
Dim pary As String
Dim tmin As String
Dim tmax As String
Dim x As String
Dim Y As String

colindex = 0
xming = ConvertNumberToMaxima(TextBox_xmin.Text)
xmaxg = ConvertNumberToMaxima(TextBox_xmax.Text)
yming = ConvertNumberToMaxima(TextBox_ymin.Text)
ymaxg = ConvertNumberToMaxima(TextBox_ymax.Text)

'forskrifter
If TextBox_ligning1.Text <> "" Then
    lign = Replace(TextBox_ligning1.Text, "'", "‰")
    If CheckBox_visforklaring.Value Then
        grafobj = grafobj & "key=""" & omax.ConvertToAscii(lign) & ""","
    Else
        grafobj = grafobj & "key="""","
    End If
    lign = omax.CodeForMaxima(lign)
'    End If
    If ComboBox_ligning1.ListIndex > 0 Then
        grafobj = grafobj & "line_type=dots,"
    Else
        grafobj = grafobj & "line_type=solid,"
    End If
    If Len(TextBox_xmin1.Text) = 0 Then
        xmin = xming
    Else
        xmin = ConvertNumberToMaxima(TextBox_xmin1.Text)
    End If
    If Len(TextBox_xmax1.Text) = 0 Then
        xmax = xmaxg
    Else
        xmax = ConvertNumberToMaxima(TextBox_xmax1.Text)
    End If
'    If Not MaximaComplex Then lign = "'CheckDef(" & lign & ",""" & TextBox_var1.text & """)"
'    If Not MaximaComplex Then lign = "'RealOnly(" & lign & ")"
    grafobj = grafobj & "color=" & GetNextColor & ",explicit(" & lign & "," & TextBox_var1.Text & "," & xmin & "," & xmax & "),"
End If
If TextBox_ligning2.Text <> "" Then
    lign = Replace(TextBox_ligning2.Text, "'", "‰")
    If CheckBox_visforklaring.Value Then
        grafobj = grafobj & "key=""" & omax.ConvertToAscii(lign) & ""","
    Else
        grafobj = grafobj & "key="""","
    End If
    lign = omax.CodeForMaxima(lign)
'    End If
    If ComboBox_ligning2.ListIndex > 0 Then
        grafobj = grafobj & "line_type=dots,"
    Else
        grafobj = grafobj & "line_type=solid,"
    End If
    If Len(TextBox_xmin2.Text) = 0 Then
        xmin = xming
    Else
        xmin = ConvertNumberToMaxima(TextBox_xmin2.Text)
    End If
    If Len(TextBox_xmax2.Text) = 0 Then
        xmax = xmaxg
    Else
        xmax = ConvertNumberToMaxima(TextBox_xmax2.Text)
    End If
'    If Not MaximaComplex Then lign = "'CheckDef(" & lign & ",""" & TextBox_var2.text & """)"
'    If Not MaximaComplex Then lign = "'RealOnly(" & lign & ")"
    grafobj = grafobj & "color=" & GetNextColor & ",explicit(" & lign & "," & TextBox_var2.Text & "," & xmin & "," & xmax & "),"
End If
If TextBox_ligning3.Text <> "" Then
    lign = Replace(TextBox_ligning3.Text, "'", "‰")
    If CheckBox_visforklaring.Value Then
        grafobj = grafobj & "key=""" & omax.ConvertToAscii(lign) & ""","
    Else
        grafobj = grafobj & "key="""","
    End If
    lign = omax.CodeForMaxima(lign)
'    End If
    If ComboBox_ligning3.ListIndex > 0 Then
        grafobj = grafobj & "line_type=dots,"
    Else
        grafobj = grafobj & "line_type=solid,"
    End If
    If Len(TextBox_xmin3.Text) = 0 Then
        xmin = xming
    Else
        xmin = ConvertNumberToMaxima(TextBox_xmin3.Text)
    End If
    If Len(TextBox_xmax3.Text) = 0 Then
        xmax = xmaxg
    Else
        xmax = ConvertNumberToMaxima(TextBox_xmax3.Text)
    End If
'    If Not MaximaComplex Then lign = "'RealOnly(" & lign & ")"
    grafobj = grafobj & "color=" & GetNextColor & ",explicit(" & lign & "," & TextBox_var3.Text & "," & xmin & "," & xmax & "),"
End If
If TextBox_ligning4.Text <> "" Then
    lign = Replace(TextBox_ligning4.Text, "'", "‰")
    If CheckBox_visforklaring.Value Then
        grafobj = grafobj & "key=""" & omax.ConvertToAscii(lign) & ""","
    Else
        grafobj = grafobj & "key="""","
    End If
    lign = omax.CodeForMaxima(lign)
'    End If
    If ComboBox_ligning4.ListIndex > 0 Then
        grafobj = grafobj & "line_type=dots,"
    Else
        grafobj = grafobj & "line_type=solid,"
    End If
    If Len(TextBox_xmin4.Text) = 0 Then
        xmin = xming
    Else
        xmin = ConvertNumberToMaxima(TextBox_xmin4.Text)
    End If
    If Len(TextBox_xmax4.Text) = 0 Then
        xmax = xmaxg
    Else
        xmax = ConvertNumberToMaxima(TextBox_xmax4.Text)
    End If
'    If Not MaximaComplex Then lign = "'RealOnly(" & lign & ")"
    grafobj = grafobj & "color=" & GetNextColor & ",explicit(" & lign & "," & TextBox_var4.Text & "," & xmin & "," & xmax & "),"
End If
If TextBox_ligning5.Text <> "" Then
    lign = Replace(TextBox_ligning5.Text, "'", "‰")
    If CheckBox_visforklaring.Value Then
        grafobj = grafobj & "key=""" & omax.ConvertToAscii(lign) & ""","
    Else
        grafobj = grafobj & "key="""","
    End If
    lign = omax.CodeForMaxima(lign)
'    End If
    If ComboBox_ligning5.ListIndex > 0 Then
        grafobj = grafobj & "line_type=dots,"
    Else
        grafobj = grafobj & "line_type=solid,"
    End If
    If Len(TextBox_xmin5.Text) = 0 Then
        xmin = xming
    Else
        xmin = ConvertNumberToMaxima(TextBox_xmin5.Text)
    End If
    If Len(TextBox_xmax5.Text) = 0 Then
        xmax = xmaxg
    Else
        xmax = ConvertNumberToMaxima(TextBox_xmax5.Text)
    End If
'    If Not MaximaComplex Then lign = "'RealOnly(" & lign & ")"
    grafobj = grafobj & "color=" & GetNextColor & ",explicit(" & lign & "," & TextBox_var5.Text & "," & xmin & "," & xmax & "),"
End If
If TextBox_ligning6.Text <> "" Then
    lign = Replace(TextBox_ligning6.Text, "'", "‰")
    If CheckBox_visforklaring.Value Then
        grafobj = grafobj & "key=""" & omax.ConvertToAscii(lign) & ""","
    Else
        grafobj = grafobj & "key="""","
    End If
    lign = omax.CodeForMaxima(lign)
'    End If
    If ComboBox_ligning6.ListIndex > 0 Then
        grafobj = grafobj & "line_type=dots,"
    Else
        grafobj = grafobj & "line_type=solid,"
    End If
    If Len(TextBox_xmin6.Text) = 0 Then
        xmin = xming
    Else
        xmin = ConvertNumberToMaxima(TextBox_xmin6.Text)
    End If
    If Len(TextBox_xmax6.Text) = 0 Then
        xmax = xmaxg
    Else
        xmax = ConvertNumberToMaxima(TextBox_xmax6.Text)
    End If
'    If Not MaximaComplex Then lign = "'RealOnly(" & lign & ")"
    grafobj = grafobj & "color=" & GetNextColor & ",explicit(" & lign & "," & TextBox_var6.Text & "," & xmin & "," & xmax & "),"
End If

'ligninger
If TextBox_lig1.Text <> "" Then
    lign = Replace(TextBox_lig1.Text, "'", "‰")
    If CheckBox_visforklaring.Value Then
        grafobj = grafobj & "key=""" & omax.ConvertToAscii(lign) & ""","
    End If
    lign = omax.CodeForMaxima(lign)
    If yming = "" Then
        ymin2 = xming
    Else
        ymin2 = yming
    End If
    If ymaxg = "" Then
        ymax2 = xmaxg
    Else
        ymax2 = ymaxg
    End If
    grafobj = grafobj & "color=" & GetNextColor & ",implicit(" & lign & ",x," & xming & "," & xmaxg & ",y," & ymin2 & "," & ymax2 & "),"
End If
If TextBox_lig2.Text <> "" Then
    lign = Replace(TextBox_lig2.Text, "'", "‰")
    If CheckBox_visforklaring.Value Then
        grafobj = grafobj & "key=""" & omax.ConvertToAscii(lign) & ""","
    End If
    lign = omax.CodeForMaxima(lign)
    If yming = "" Then
        ymin2 = xming
    Else
        ymin2 = yming
    End If
    If ymaxg = "" Then
        ymax2 = xmaxg
    Else
        ymax2 = ymaxg
    End If
    grafobj = grafobj & "color=" & GetNextColor & ",implicit(" & lign & ",x," & xming & "," & xmaxg & ",y," & ymin2 & "," & ymax2 & "),"
End If
If TextBox_Lig3.Text <> "" Then
    lign = Replace(TextBox_Lig3.Text, "'", "‰")
    If CheckBox_visforklaring.Value Then
        grafobj = grafobj & "key=""" & omax.ConvertToAscii(lign) & ""","
    End If
    lign = omax.CodeForMaxima(lign)
    If yming = "" Then
        ymin2 = xming
    Else
        ymin2 = yming
    End If
    If ymaxg = "" Then
        ymax2 = xmaxg
    Else
        ymax2 = ymaxg
    End If
    grafobj = grafobj & "color=" & GetNextColor & ",implicit(" & lign & ",x," & xming & "," & xmaxg & ",y," & ymin2 & "," & ymax2 & "),"
End If

'parameterfremstillinger
If TextBox_parametric1x.Text <> "" Then
    parx = omax.CodeForMaxima(TextBox_parametric1x.Text)
    pary = omax.CodeForMaxima(TextBox_parametric1y.Text)
    tmin = ConvertNumberToMaxima(TextBox_tmin1.Text)
    tmax = ConvertNumberToMaxima(TextBox_tmax1.Text)
    If CheckBox_visforklaring.Value Then
        grafobj = grafobj & "key=""(" & omax.ConvertToAscii(TextBox_parametric1x.Text) & "," & omax.ConvertToAscii(TextBox_parametric1y.Text) & ")"","
    Else
        grafobj = grafobj & "key="""","
    End If
    grafobj = grafobj & "line_type=solid,color=" & GetNextColor & ","
    grafobj = grafobj & "parametric(" & parx & "," & pary & ",t," & tmin & "," & tmax & "),"
End If
If TextBox_parametric2x.Text <> "" Then
    parx = omax.CodeForMaxima(TextBox_parametric2x.Text)
    pary = omax.CodeForMaxima(TextBox_parametric2y.Text)
    tmin = ConvertNumberToMaxima(TextBox_tmin2.Text)
    tmax = ConvertNumberToMaxima(TextBox_tmax2.Text)
    If CheckBox_visforklaring.Value Then
        grafobj = grafobj & "key=""(" & omax.ConvertToAscii(TextBox_parametric2x.Text) & "," & omax.ConvertToAscii(TextBox_parametric2y.Text) & ")"","
    Else
        grafobj = grafobj & "key="""","
    End If
    grafobj = grafobj & "line_type=solid,color=" & GetNextColor & ","
    grafobj = grafobj & "parametric(" & parx & "," & pary & ",t," & tmin & "," & tmax & "),"
End If
If TextBox_parametric3x.Text <> "" Then
    parx = omax.CodeForMaxima(TextBox_parametric3x.Text)
    pary = omax.CodeForMaxima(TextBox_parametric3y.Text)
    tmin = ConvertNumberToMaxima(TextBox_tmin3.Text)
    tmax = ConvertNumberToMaxima(TextBox_tmax3.Text)
    If CheckBox_visforklaring.Value Then
        grafobj = grafobj & "key=""(" & omax.ConvertToAscii(TextBox_parametric3x.Text) & "," & omax.ConvertToAscii(TextBox_parametric3y.Text) & ")"","
    Else
        grafobj = grafobj & "key="""","
    End If
    grafobj = grafobj & "line_type=solid,color=" & GetNextColor & ","
    grafobj = grafobj & "parametric(" & parx & "," & pary & ",t," & tmin & "," & tmax & "),"
End If


'punkter
If TextBox_punkter.Text <> "" Then
    grafobj = grafobj & "key="""",color=black,"
    Arr = Split(TextBox_punkter.Text, VbCrLfMac)
    For i = 0 To UBound(Arr)
    If InStr(Arr(i), ";") > 0 Or InStr(Arr(i), vbTab) > 0 Then
        Arr(i) = Replace(Arr(i), ",", ".")
        Arr(i) = Replace(Arr(i), ";", ",")
    End If
        Arr(i) = Replace(Arr(i), vbTab, ",") ' hvis tab. f.eks. hvis kopieret fra excel
        Arr(i) = Replace(Arr(i), " ", "")
        If Len(Arr(i)) > 0 Then
        If Left(Arr(i), 1) <> "(" Then
            Arr(i) = "(" & Arr(i)
        End If
        If right(Arr(i), 1) <> ")" Then
            Arr(i) = Arr(i) & ")"
        End If
        Arr(i) = Replace(Arr(i), "),(", "],[")
        Arr(i) = Replace(Arr(i), ");(", "],[")
        Arr(i) = Replace(Arr(i), "(", "[")
        Arr(i) = Replace(Arr(i), ")", "]")
        punkttekst = punkttekst & Arr(i) & ","
        End If
    Next
    If right(punkttekst, 1) = "," Then punkttekst = Left(punkttekst, Len(punkttekst) - 1)
    
    grafobj = grafobj & "point_type=filled_circle,point_size=" & Replace(highres * ConvertStringToNumber(TextBox_pointsize.Text), ",", ".") & ",points_joined=" & VBA.LCase(CheckBox_pointsjoined.Value) & ",points([" & punkttekst & "]),"
End If

'punkter 2
If TextBox_punkter2.Text <> "" Then
    punkttekst = ""
    grafobj = grafobj & "key="""",color=blue,"
    Arr = Split(TextBox_punkter2.Text, VbCrLfMac)
    For i = 0 To UBound(Arr)
    If InStr(Arr(i), ";") > 0 Or InStr(Arr(i), vbTab) > 0 Then
        Arr(i) = Replace(Arr(i), ",", ".")
        Arr(i) = Replace(Arr(i), ";", ",")
    End If
        Arr(i) = Replace(Arr(i), vbTab, ",") ' hvis tab. f.eks. hvis kopieret fra excel
        Arr(i) = Replace(Arr(i), " ", "")
        If Len(Arr(i)) > 0 Then
        If Left(Arr(i), 1) <> "(" Then
            Arr(i) = "(" & Arr(i)
        End If
        If right(Arr(i), 1) <> ")" Then
            Arr(i) = Arr(i) & ")"
        End If
        Arr(i) = Replace(Arr(i), "),(", "],[")
        Arr(i) = Replace(Arr(i), ");(", "],[")
        Arr(i) = Replace(Arr(i), "(", "[")
        Arr(i) = Replace(Arr(i), ")", "]")
        punkttekst = punkttekst & Arr(i) & ","
        End If
    Next
    If right(punkttekst, 1) = "," Then punkttekst = Left(punkttekst, Len(punkttekst) - 1)
    
    grafobj = grafobj & "point_type=filled_circle,point_size=" & Replace(TextBox_pointsize2.Text, ",", ".") & ",points_joined=" & VBA.LCase(CheckBox_pointsjoined2.Value) & ",points([" & punkttekst & "]),"
End If

'markerede punkter
If TextBox_markerpunkter.Text <> "" Then
    punkttekst = ""
    grafobj = grafobj & "key="""",color=red,"
    Arr = Split(TextBox_markerpunkter.Text, VbCrLfMac)
    For i = 0 To UBound(Arr)
    If InStr(Arr(i), ";") > 0 Or InStr(Arr(i), vbTab) > 0 Then
        Arr(i) = Replace(Arr(i), ",", ".")
        Arr(i) = Replace(Arr(i), ";", ",")
    End If
        Arr(i) = Replace(Arr(i), vbTab, ",") ' hvis tab. f.eks. hvis kopieret fra excel
        Arr(i) = Replace(Arr(i), " ", "")
        If Len(Arr(i)) > 0 Then
'        If Left(arr(i), 1) = "(" Then
'            arr(i) = Right(arr(i), Len(arr(i)) - 1)
'        End If
'        If Right(arr(i), 1) = ")" Then
'            arr(i) = Left(arr(i), Len(arr(i)) - 1)
'        End If
        arr2 = Split(Arr(i), ",")
        If UBound(arr2) = 1 Then
            x = arr2(0)
            Y = arr2(1)
            punkttekst = punkttekst & "points([[" & x & ",0],[" & x & "," & Y & "],[0," & Y & "]]),"
        End If
        End If
    Next
'    If Right(punkttekst, 1) = "," Then punkttekst = Left(punkttekst, Len(punkttekst) - 1)
    
    grafobj = grafobj & "line_type=dots,line_width=" & Replace(highres, ",", ".") & ",point_size=0.1,points_joined=true," & punkttekst
End If


'vektorer
If TextBox_vektorer.Text <> "" Then
    vekt = TextBox_vektorer.Text
    Arr = Split(vekt, VbCrLfMac)
    For i = 0 To UBound(Arr)
        If Arr(i) <> "" Then
    If InStr(Arr(i), ";") > 0 Then
        Arr(i) = Replace(Arr(i), ",", ".")
        Arr(i) = Replace(Arr(i), ";", ",")
    End If
    If InStr(Arr(i), ")-(") > 0 Then
        Arr(i) = Replace(Arr(i), ")-(", "],[")
    ElseIf InStr(Arr(i), ")(") > 0 Then
        Arr(i) = Replace(Arr(i), ")(", "],[")
    Else
        Arr(i) = "[0,0]," & Arr(i)
    End If
    Arr(i) = Replace(Arr(i), "(", "[")
    Arr(i) = Replace(Arr(i), ")", "]")
            If CheckBox_visforklaring.Value Then
                grafobj = grafobj & "key=""Vektor: " & Arr(i) & ""","
            Else
                grafobj = grafobj & "key="""","
            End If
            grafobj = grafobj & "color=" & GetNextColor & ","
            grafobj = grafobj & "line_type=solid,line_width=" & Replace(highres, ",", ".") & ",head_angle=25,head_length=" & Replace((ConvertStringToNumber(TextBox_xmax.Text) - ConvertStringToNumber(TextBox_xmin.Text)) / 40, ",", ".") & ","
            grafobj = grafobj & "vector(" & Arr(i) & "),"
        End If
    Next
End If

    'labels
    If Len(TextBox_labels.Text) > 0 Then
        Arr = Split(TextBox_labels.Text, VbCrLfMac)
        For i = 0 To UBound(Arr)
            If InStr(Arr(i), ";") > 0 Then
                arr2 = Split(Arr(i), ";")
            Else
                arr2 = Split(Arr(i), ",")
            End If
            If UBound(arr2) >= 2 Then
            labeltext = labeltext & "["
            labeltext = labeltext & """" & arr2(0) & """"
            labeltext = labeltext & "," & Replace(arr2(1), ",", ".")
            labeltext = labeltext & "," & Replace(arr2(2), ",", ".")
            labeltext = labeltext & "],"
            End If
        Next
        If Len(labeltext) > 0 Then
            labeltext = Left(labeltext, Len(labeltext) - 1)
            grafobj = "color=black,label(" & ConvertDrawLabel(labeltext) & ")," & grafobj
        End If
    End If
    
    If Len(grafobj) = 0 Then GoTo Slut ' ellers fejler når print starter op, men med denne fejler retningsfelt hvis alene

' diverse
    If Len(TextBox_xmin.Text) > 0 And Len(TextBox_xmax.Text) > 0 Then
        grafobj = "xrange=[" & ConvertNumberToMaxima(TextBox_xmin.Text) & "," & ConvertNumberToMaxima(TextBox_xmax.Text) & "]," & grafobj
    End If
    If Len(TextBox_ymin.Text) > 0 And Len(TextBox_ymax.Text) > 0 And Len(TextBox_dfligning.Text) = 0 Then
        grafobj = "yrange=[" & ConvertNumberToMaxima(TextBox_ymin.Text) & "," & ConvertNumberToMaxima(TextBox_ymax.Text) & "]," & grafobj
    End If
'    If ToggleButton_propto.value Then
'        grafobj = "proportional_axes = xy," & grafobj
'    End If
    
    grafobj = "font=""Arial"",font_size=8," & grafobj
    grafobj = "nticks=100," & grafobj
    grafobj = "ip_grid=[70,70]," & grafobj
'    grafobj = "xu_grid=50,yv_grid=50," & grafobj ' ser ikke ud til at være nødv
    grafobj = "xtics_axis = true," & grafobj
    grafobj = "ytics_axis = true," & grafobj
    grafobj = "line_width=" & Replace(highres, ",", ".") & "," & grafobj
    If Not MaximaComplex Then grafobj = "draw_realpart = false," & grafobj
    
    If CheckBox_logx.Value Then
        If ConvertStringToNumber(TextBox_xmin.Text) > 0 Then
            grafobj = "logx=true," & grafobj
        Else
            MsgBox "xmin skal være >0 for at bruge logaritmisk x-akse."
        End If
    End If
    If CheckBox_logy.Value Then
        If ConvertStringToNumber(TextBox_ymin.Text) > 0 Then
            grafobj = "logy=true," & grafobj
        Else
            MsgBox "ymin skal være >0 for at bruge logaritmisk y-akse."
        End If
    End If
    
    If Len(TextBox_titel.Text) > 0 Then
        grafobj = "title=""" & omax.ConvertToAscii(TextBox_titel.Text) & """," & grafobj
    End If
    
'        grafobj = "user_preamble = ""set grid lc 2""," & grafobj ' lw 2 er linewidth 2, lt er linetype, lc er linecolor
        'grid lt 0 lw 1 lc

'    grafobj = "user_preamble = ""set xyplane at 0""," & grafobj 'palette=gray,
    If Len(grafobj) > 0 Then
        grafobj = Left(grafobj, Len(grafobj) - 1)
    End If

    GetDraw2Dtext = grafobj
GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:
End Function
Private Sub OpdaterGraf(Optional highres As Double = 1)
Dim Text As String
Dim df As String
Dim dfsol As String
Dim pm As String
On Error GoTo Fejl
    Label_wait.Caption = Sprog.Wait & "!"
    Label_wait.Font.Size = 36
    Label_wait.visible = True
    omax.PrepareNewCommand finddef:=False  ' uden at søge efter definitioner i dokument
    InsertDefinitioner
    Text = GetDraw2Dtext(highres)
    If Len(TextBox_dfligning.Text) > 0 Then
        df = omax.CodeForMaxima(TextBox_dfligning.Text)
        df = df & ",[" & TextBox_dfx.Text & "," & TextBox_dfy.Text & "]"
        If Len(TextBox_xmin.Text) > 0 And Len(TextBox_ymin.Text) > 0 And Len(TextBox_xmax.Text) > 0 And Len(TextBox_ymax.Text) > 0 Then
            df = df & ",[" & TextBox_dfx.Text & "," & ConvertNumberToMaxima(TextBox_xmin.Text) & "," & ConvertNumberToMaxima(TextBox_xmax.Text) & "],[" & TextBox_dfy.Text & "," & ConvertNumberToMaxima(TextBox_ymin.Text) & "," & ConvertNumberToMaxima(TextBox_ymax.Text) & "]"
        ElseIf Len(TextBox_xmin.Text) > 0 And Len(TextBox_xmax.Text) > 0 Then
            df = df & ",[" & TextBox_dfx.Text & "," & ConvertNumberToMaxima(TextBox_xmin.Text) & "," & ConvertNumberToMaxima(TextBox_xmax.Text) & "]"
        Else
            df = df & ",[" & TextBox_dfx.Text & "," & TextBox_dfy.Text & "]"
        End If
        df = df & ",field_arrows=false"
        If Len(TextBox_dfsol1x.Text) > 0 And Len(TextBox_dfsol1x.Text) > 0 Then
            dfsol = dfsol & "[" & ConvertNumberToMaxima(TextBox_dfsol1x.Text) & "," & ConvertNumberToMaxima(TextBox_dfsol1y.Text) & "]"
        End If
        If Len(TextBox_dfsol2x.Text) > 0 And Len(TextBox_dfsol2x.Text) > 0 Then
            If Len(dfsol) > 0 Then dfsol = dfsol & ","
            dfsol = dfsol & "[" & ConvertNumberToMaxima(TextBox_dfsol2x.Text) & "," & ConvertNumberToMaxima(TextBox_dfsol2y.Text) & "]"
        End If
        If Len(TextBox_dfsol3x.Text) > 0 And Len(TextBox_dfsol3x.Text) > 0 Then
            If Len(dfsol) > 0 Then dfsol = dfsol & ","
            dfsol = dfsol & "[" & ConvertNumberToMaxima(TextBox_dfsol3x.Text) & "," & ConvertNumberToMaxima(TextBox_dfsol3y.Text) & "]"
        End If
        If Len(TextBox_dfsol4x.Text) > 0 And Len(TextBox_dfsol4x.Text) > 0 Then
            If Len(dfsol) > 0 Then dfsol = dfsol & ","
            dfsol = dfsol & "[" & ConvertNumberToMaxima(TextBox_dfsol4x.Text) & "," & ConvertNumberToMaxima(TextBox_dfsol4y.Text) & "]"
        End If
        If Len(TextBox_dfsol5x.Text) > 0 And Len(TextBox_dfsol5x.Text) > 0 Then
            If Len(dfsol) > 0 Then dfsol = dfsol & ","
            dfsol = dfsol & "[" & ConvertNumberToMaxima(TextBox_dfsol5x.Text) & "," & ConvertNumberToMaxima(TextBox_dfsol5y.Text) & "]"
        End If
        If Len(dfsol) > 0 Then
            df = df & ",duration=100,solns_at(" & dfsol & ")" ' duration defaulat er 10. ved at øge plottes længere og tættere på asymptoter
        End If
        If CheckBox_onlykurver.Value Then
            df = df & ",show_field=false"
        End If
        If Text = "" Then ' must be range
            If Len(TextBox_xmin.Text) > 0 And Len(TextBox_xmax.Text) > 0 Then
                Text = "xrange=[" & ConvertNumberToMaxima(TextBox_xmin.Text) & "," & ConvertNumberToMaxima(TextBox_xmax.Text) & "]"
            End If
            If Len(TextBox_ymin.Text) > 0 And Len(TextBox_ymax.Text) > 0 And Len(TextBox_dfligning.Text) = 0 Then
                Text = Text & ",yrange=[" & ConvertNumberToMaxima(TextBox_ymin.Text) & "," & ConvertNumberToMaxima(TextBox_ymax.Text) & "]"
            End If
        End If
    End If
    If Len(Text) > 0 Then
        Call omax.Draw2D(Text, df, ConvertDrawLabel(TextBox_xaksetitel.Text), ConvertDrawLabel(TextBox_yaksetitel.Text), CheckBox_gitter.Value, True, highres)
        If omax.MaximaOutput = "" Then
            Label_wait.Caption = "Fejl!"
            Label_wait.visible = True
            GoTo Slut
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
GoTo Slut
Fejl:
    On Error Resume Next
    Label_wait.Caption = Sprog.A(94)
    Label_wait.Font.Size = 12
    Label_wait.Width = 150
    Label_wait.visible = True
    Image1.Picture = Nothing
Slut:

End Sub
Private Sub GnuPlot()
Dim Text As String
    omax.PrepareNewCommand finddef:=False  ' uden at søge efter definitioner i dokument
    InsertDefinitioner

    Text = GetDraw2Dtext()
    
    If Len(Text) > 0 Then
    Call omax.Draw2D(Text, "", omax.ConvertToAscii(TextBox_xaksetitel.Text), omax.ConvertToAscii(TextBox_yaksetitel.Text), CheckBox_gitter.Value, False, 1)
    DoEvents
    End If

GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:
End Sub

Private Sub CommandButton_linregr_Click()
    Dim Cregr As New CRegression
    On Error GoTo Slut
    Cregr.Datatext = TextBox_punkter.Text
    Cregr.ComputeLinRegr
'    Selection.Collapse
'    Selection.TypeParagraph
'    Cregr.InsertEquation

    If TextBox_ligning1.Text = "" Then
        TextBox_ligning1.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning2.Text = "" Then
        TextBox_ligning2.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning3.Text = "" Then
        TextBox_ligning3.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning4.Text = "" Then
        TextBox_ligning4.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning5.Text = "" Then
        TextBox_ligning5.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning6.Text = "" Then
        TextBox_ligning6.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    End If

    OpdaterGraf
    Me.Repaint
Slut:
End Sub

Private Sub CommandButton_polregr_Click()
    Dim Cregr As New CRegression
    On Error GoTo Slut
    
    Cregr.Datatext = TextBox_punkter.Text
    Cregr.ComputePolRegr
'    Selection.Collapse
'    Selection.TypeParagraph
'    Cregr.InsertEquation

    If TextBox_ligning1.Text = "" Then
        TextBox_ligning1.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning2.Text = "" Then
        TextBox_ligning2.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning3.Text = "" Then
        TextBox_ligning3.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning4.Text = "" Then
        TextBox_ligning4.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning5.Text = "" Then
        TextBox_ligning5.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning6.Text = "" Then
        TextBox_ligning6.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    End If
    
    OpdaterGraf
    Me.Repaint
Slut:
End Sub
Private Sub CommandButton_ekspregr_Click()
   Dim Cregr As New CRegression
    On Error GoTo Slut
    
    Cregr.Datatext = TextBox_punkter.Text
    Cregr.ComputeExpRegr
'    Selection.Collapse
'    Selection.TypeParagraph
'    Cregr.InsertEquation

    If TextBox_ligning1.Text = "" Then
        TextBox_ligning1.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning2.Text = "" Then
        TextBox_ligning2.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning3.Text = "" Then
        TextBox_ligning3.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning4.Text = "" Then
        TextBox_ligning4.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning5.Text = "" Then
        TextBox_ligning5.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning6.Text = "" Then
        TextBox_ligning6.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    End If
    
    OpdaterGraf
    Me.Repaint
Slut:
End Sub

Private Sub CommandButton_potregr_Click()
       Dim Cregr As New CRegression
    On Error GoTo Slut
    
    Cregr.Datatext = TextBox_punkter.Text
    Cregr.ComputePowRegr
'    Selection.Collapse
'    Selection.TypeParagraph
'    Cregr.InsertEquation

    If TextBox_ligning1.Text = "" Then
        TextBox_ligning1.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning2.Text = "" Then
        TextBox_ligning2.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning3.Text = "" Then
        TextBox_ligning3.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning4.Text = "" Then
        TextBox_ligning4.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning5.Text = "" Then
        TextBox_ligning5.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning6.Text = "" Then
        TextBox_ligning6.Text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    End If

    OpdaterGraf
    Me.Repaint
Slut:
End Sub

Private Sub CommandButton_nulstil_Click()
    TextBox_ligning1.Text = ""
    TextBox_ligning2.Text = ""
    TextBox_ligning3.Text = ""
    TextBox_ligning4.Text = ""
    TextBox_ligning5.Text = ""
    TextBox_ligning6.Text = ""
    TextBox_lig1.Text = ""
    TextBox_lig2.Text = ""
    TextBox_Lig3.Text = ""
    TextBox_xmin1.Text = ""
    TextBox_xmin2.Text = ""
    TextBox_xmin3.Text = ""
    TextBox_xmin4.Text = ""
    TextBox_xmin5.Text = ""
    TextBox_xmin6.Text = ""
    TextBox_xmax1.Text = ""
    TextBox_xmax2.Text = ""
    TextBox_xmax3.Text = ""
    TextBox_xmax4.Text = ""
    TextBox_xmax5.Text = ""
    TextBox_xmax6.Text = ""
    TextBox_xmin.Text = "-5"
    TextBox_xmax.Text = "5"
    TextBox_ymin.Text = ""
    TextBox_ymax.Text = ""
    TextBox_xaksetitel.Text = ""
    TextBox_yaksetitel.Text = ""
    TextBox_titel.Text = ""
    TextBox_punkter.Text = ""
    TextBox_punkter2.Text = ""
    TextBox_labels.Text = ""
    TextBox_vektorer.Text = ""
    Call FillLineStyleCombos
End Sub
Private Sub MMathPlot()
    Dim MathSyntax As String
    Dim objEq As OMath
    Dim objRange As Range
    Dim plinjer As Variant
    Dim xmin As Double
    Dim mmax As Double
    Dim i As Integer
    Dim linje As Variant
    
    Set objRange = Selection.Range
    MathSyntax = "show("
    
    If TextBox_ligning1.Text <> "" Then
        If Len(MathSyntax) > 5 Then MathSyntax = MathSyntax + ","
        If InStr(TextBox_ligning1.Text, "=") Then
            MathSyntax = MathSyntax & "ploteq("
        Else
            MathSyntax = MathSyntax & "plot("
        End If
        MathSyntax = MathSyntax & Replace(TextBox_ligning1.Text, ",", ".")
        If TextBox_xmin1.Text <> "" And TextBox_xmax1.Text <> "" Then
            MathSyntax = MathSyntax + ",{x," & Me.TextBox_xmin1.Text & "," & Me.TextBox_xmax1.Text & "}"
        End If
'        If Me.TextBox_xmin1.text <> "" And Me.TextBox_ymax1.text <> "" Then
'            MathSyntax = MathSyntax + ",{y," & Me.TextBox_ymin1.text & "," & Me.TextBox_ymax1.text & "}"
'        End If
        If ComboBox_ligning1.Text = "---" Then
        ElseIf ComboBox_ligning1.Text = "..." Then
            MathSyntax = MathSyntax + ",{LineStyle,Dot}"
        ElseIf ComboBox_ligning1.Text = "- - -" Then
            MathSyntax = MathSyntax + ",{LineStyle,Dash}"
        ElseIf ComboBox_ligning1.Text = "-.-.-." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDot}"
        ElseIf ComboBox_ligning1.Text = "- . . - . ." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDotDot}"
        End If
'        If TextBox_samplesize.text <> "" Then
'            MathSyntax = MathSyntax + ",{SampleSize," & TextBox_samplesize.text & "}"
'        End If
        MathSyntax = MathSyntax + ")"
    End If
    If TextBox_ligning2.Text <> "" Then
        If Len(MathSyntax) > 5 Then MathSyntax = MathSyntax + ","
        If InStr(TextBox_ligning2.Text, "=") Then
            MathSyntax = MathSyntax & "ploteq("
        Else
            MathSyntax = MathSyntax & "plot("
        End If
        MathSyntax = MathSyntax & Replace(TextBox_ligning2.Text, ",", ".")
        If TextBox_xmin2.Text <> "" And TextBox_xmax2.Text <> "" Then
            MathSyntax = MathSyntax + ",{x," & TextBox_xmin2.Text & "," & TextBox_xmax2.Text & "}"
        End If
'        If TextBox_ymin2.text <> "" And TextBox_ymax2.text <> "" Then
'            MathSyntax = MathSyntax + ",{y," & TextBox_ymin2.text & "," & TextBox_ymax2.text & "}"
'        End If
        If ComboBox_ligning2.Text = "---" Then
        ElseIf ComboBox_ligning2.Text = "..." Then
            MathSyntax = MathSyntax + ",{LineStyle,Dot}"
        ElseIf ComboBox_ligning2.Text = "- - -" Then
            MathSyntax = MathSyntax + ",{LineStyle,Dash}"
        ElseIf ComboBox_ligning2.Text = "-.-.-." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDot}"
        ElseIf ComboBox_ligning2.Text = "- . . - . ." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDotDot}"
        End If
'        If TextBox_samplesize.text <> "" Then
'            MathSyntax = MathSyntax + ",{SampleSize," & TextBox_samplesize.text & "}"
'        End If
        MathSyntax = MathSyntax + ")"
    End If
    If TextBox_ligning3.Text <> "" Then
        If Len(MathSyntax) > 5 Then MathSyntax = MathSyntax + ","
        If InStr(TextBox_ligning3.Text, "=") Then
            MathSyntax = MathSyntax & "ploteq("
        Else
            MathSyntax = MathSyntax & "plot("
        End If
        MathSyntax = MathSyntax & Replace(TextBox_ligning3.Text, ",", ".")
        If TextBox_xmin3.Text <> "" And TextBox_xmax3.Text <> "" Then
            MathSyntax = MathSyntax + ",{x," & TextBox_xmin3.Text & "," & TextBox_xmax3.Text & "}"
        End If
'        If TextBox_ymin3.text <> "" And TextBox_ymax3.text <> "" Then
'            MathSyntax = MathSyntax + ",{y," & TextBox_ymin3.text & "," & TextBox_ymax3.text & "}"
'        End If
        If ComboBox_ligning3.Text = "---" Then
        ElseIf ComboBox_ligning3.Text = "..." Then
            MathSyntax = MathSyntax + ",{LineStyle,Dot}"
        ElseIf ComboBox_ligning3.Text = "- - -" Then
            MathSyntax = MathSyntax + ",{LineStyle,Dash}"
        ElseIf ComboBox_ligning3.Text = "-.-.-." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDot}"
        ElseIf ComboBox_ligning3.Text = "- . . - . ." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDotDot}"
        End If
'        If TextBox_samplesize.text <> "" Then
'            MathSyntax = MathSyntax + ",{SampleSize," & TextBox_samplesize.text & "}"
'        End If
        MathSyntax = MathSyntax + ")"
    End If
    If TextBox_ligning4.Text <> "" Then
        If Len(MathSyntax) > 5 Then MathSyntax = MathSyntax + ","
        If InStr(TextBox_ligning4.Text, "=") Then
            MathSyntax = MathSyntax & "ploteq("
        Else
            MathSyntax = MathSyntax & "plot("
        End If
        MathSyntax = MathSyntax & Replace(TextBox_ligning4.Text, ",", ".")
        If TextBox_xmin4.Text <> "" And TextBox_xmax4.Text <> "" Then
            MathSyntax = MathSyntax + ",{x," & TextBox_xmin4.Text & "," & TextBox_xmax4.Text & "}"
        End If
'        If TextBox_ymin4.text <> "" And TextBox_ymax4.text <> "" Then
'            MathSyntax = MathSyntax + ",{y," & TextBox_ymin4.text & "," & TextBox_ymax4.text & "}"
'        End If
        If ComboBox_ligning4.Text = "---" Then
        ElseIf ComboBox_ligning4.Text = "..." Then
            MathSyntax = MathSyntax + ",{LineStyle,Dot}"
        ElseIf ComboBox_ligning4.Text = "- - -" Then
            MathSyntax = MathSyntax + ",{LineStyle,Dash}"
        ElseIf ComboBox_ligning4.Text = "-.-.-." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDot}"
        ElseIf ComboBox_ligning4.Text = "- . . - . ." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDotDot}"
        End If
'        If TextBox_samplesize.text <> "" Then
'            MathSyntax = MathSyntax + ",{SampleSize," & TextBox_samplesize.text & "}"
'        End If
        MathSyntax = MathSyntax + ")"
    End If
    If TextBox_ligning5.Text <> "" Then
        If Len(MathSyntax) > 5 Then MathSyntax = MathSyntax + ","
        If InStr(TextBox_ligning5.Text, "=") Then
            MathSyntax = MathSyntax & "ploteq("
        Else
            MathSyntax = MathSyntax & "plot("
        End If
        MathSyntax = MathSyntax & Replace(TextBox_ligning5.Text, ",", ".")
        If TextBox_xmin5.Text <> "" And TextBox_xmax5.Text <> "" Then
            MathSyntax = MathSyntax + ",{x," & TextBox_xmin5.Text & "," & TextBox_xmax5.Text & "}"
        End If
'        If TextBox_ymin5.text <> "" And TextBox_ymax5.text <> "" Then
'            MathSyntax = MathSyntax + ",{y," & TextBox_ymin5.text & "," & TextBox_ymax5.text & "}"
'        End If
        If ComboBox_ligning5.Text = "---" Then
        ElseIf ComboBox_ligning5.Text = "..." Then
            MathSyntax = MathSyntax + ",{LineStyle,Dot}"
        ElseIf ComboBox_ligning5.Text = "- - -" Then
            MathSyntax = MathSyntax + ",{LineStyle,Dash}"
        ElseIf ComboBox_ligning5.Text = "-.-.-." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDot}"
        ElseIf ComboBox_ligning5.Text = "- . . - . ." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDotDot}"
        End If
        MathSyntax = MathSyntax + ")"
    End If
    If TextBox_ligning6.Text <> "" Then
        If Len(MathSyntax) > 5 Then MathSyntax = MathSyntax + ","
        If InStr(TextBox_ligning6.Text, "=") Then
            MathSyntax = MathSyntax & "ploteq("
        Else
            MathSyntax = MathSyntax & "plot("
        End If
        MathSyntax = MathSyntax & Replace(TextBox_ligning6.Text, ",", ".")
        If TextBox_xmin6.Text <> "" And TextBox_xmax6.Text <> "" Then
            MathSyntax = MathSyntax + ",{x," & TextBox_xmin6.Text & "," & TextBox_xmax6.Text & "}"
        End If
        If ComboBox_ligning6.Text = "---" Then
        ElseIf ComboBox_ligning6.Text = "..." Then
            MathSyntax = MathSyntax + ",{LineStyle,Dot}"
        ElseIf ComboBox_ligning6.Text = "- - -" Then
            MathSyntax = MathSyntax + ",{LineStyle,Dash}"
        ElseIf ComboBox_ligning6.Text = "-.-.-." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDot}"
        ElseIf ComboBox_ligning6.Text = "- . . - . ." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDotDot}"
        End If
        MathSyntax = MathSyntax + ")"
    End If
    
    'datapunkter
    If TextBox_punkter.Text <> "" Then
        If Len(MathSyntax) > 5 Then MathSyntax = MathSyntax + ","
        MathSyntax = MathSyntax + "plotdataset({"
        Dim punkttekst As String
        punkttekst = TextBox_punkter.Text
'        Dim plinjer(20) As String
        plinjer = Split(punkttekst, VbCrLfMac)
        For i = 0 To UBound(plinjer)
            linje = Split(plinjer(i), ",")
            If UBound(linje) > 0 Then
                MathSyntax = MathSyntax + "{" & linje(0) & "," & linje(1) & "},"
            End If
        Next
        MathSyntax = Left(MathSyntax, Len(MathSyntax) - 1)
        
        MathSyntax = MathSyntax + "})"
    End If
    
    ' xmin og xmax
    If TextBox_xmin.Text <> "" And TextBox_xmax.Text <> "" Then
        MathSyntax = MathSyntax + ",{x," & TextBox_xmin.Text & "," & TextBox_xmax.Text & "}"
    End If
    ' ymin og ymax
    If TextBox_ymin.Text <> "" And TextBox_ymax.Text <> "" Then
        MathSyntax = MathSyntax + ",{y," & TextBox_ymin.Text & "," & TextBox_ymax.Text & "}"
    End If
    
    If TextBox_xaksetitel.Text <> "" Then
            MathSyntax = MathSyntax + ",{aliasX," & TextBox_xaksetitel.Text & "}"
    End If
    If TextBox_yaksetitel.Text <> "" Then
            MathSyntax = MathSyntax + ",{aliasY," & TextBox_yaksetitel.Text & "}"
    End If
    
    
    MathSyntax = MathSyntax + ",{ShowBox,false}"
'    MathSyntax = MathSyntax + ",{color,""00ff00""}"
'    MathSyntax = MathSyntax + ",{color,""ff00ff00""}" ' første er opaque 00 gennemsigtig ff uigennemsigtig
'    MathSyntax = MathSyntax + ",{ShowGrid,true}"
'    MathSyntax = MathSyntax + ",{ShowAxis,true}"
'    MathSyntax = MathSyntax + ",{ShowWireframe,false}"
'    MathSyntax = MathSyntax + ",{Proportional,false}"
    MathSyntax = MathSyntax + ")"
    If Selection.OMaths.Count > 0 Then
        Selection.OMaths(Selection.OMaths.Count).Range.Select
        Selection.Range.Collapse (wdCollapseEnd)
        Selection.MoveRight Unit:=wdWord, Count:=1
    End If
    
'    Selection.Range.Collapse (wdCollapseEnd)
'    Selection.MoveRight Unit:=wdWord, Count:=1
    Selection.TypeParagraph
    
   Selection.OMaths.Add Range:=Selection.Range
     Selection.TypeText Text:=MathSyntax
    Selection.OMaths.BuildUp
'    Set objEq = objRange.OMaths(1)
'    objEq.BuildUp
    Selection.TypeParagraph

    Hide
    
Slut:
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' skjuler istedet for at lukke, så funktioner gemmes.
'If CloseMode = vbFormControlMenu Then
'    Cancel = 1
'    Hide
'End If
'Unload Me
PicOpen = False
End Sub

Function ConvertDegreeToRad(Text As String, trigfunc As String) As String
    Dim Pos, spos As Integer
    Dim ea As New ExpressionAnalyser
    ea.StartBracket = "("
    ea.EndBracket = ")"
    ea.Text = Text
    spos = 1
    
    Do
    Pos = ea.FindChr("arc" & trigfunc, spos)
    If Pos > 0 Then
        ea.GetNextBracketContent
        ea.InsertBeforePos (")")
        ea.Pos = Pos
        ea.InsertBeforePos ("180/PI()*(")
        spos = Pos + 13
    End If
    Loop While Pos > 0
    
    spos = 1
    Do
    Pos = ea.FindChr(trigfunc, spos)
    If Pos > 0 Then
        If Not (ea.ChrByIndex(Pos - 1) = "a") Then
        ea.GetNextBracketContent
        ea.InsertBeforePos (")")
        ea.Pos = Pos + Len(trigfunc)
        ea.InsertAfterPos ("PI()/180*(")
        spos = Pos + 13
        Else
            spos = Pos + 3
        End If
    End If
    Loop While Pos > 0
    
    ConvertDegreeToRad = ea.Text

End Function

Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    gemx = x
    gemy = Y
    If Len(etikettext) > 0 Then
        If Len(TextBox_labels.Text) > 0 Then
        TextBox_labels.Text = TextBox_labels.Text & VbCrLfMac
        End If
        If Len(TextBox_ymin.Text) = 0 Or Len(TextBox_ymax.Text) = 0 Then
            MsgBox Sprog.A(301), vbOKOnly, Sprog.Error
        Else
        TextBox_labels.Text = TextBox_labels.Text & etikettext & ";" & ConvertPixelToCoordX(x) & ";" & ConvertPixelToCoordY(Y)
        etikettext = ""
        OpdaterGraf
        Me.Repaint
        End If
    ElseIf nytpunkt Then
        nytpunkt = False
        If Len(TextBox_punkter2.Text) > 0 Then
        TextBox_punkter2.Text = TextBox_punkter2.Text & VbCrLfMac
        End If
        If Len(TextBox_ymin.Text) = 0 Or Len(TextBox_ymax.Text) = 0 Then
            MsgBox Sprog.A(301), vbOKOnly, Sprog.Error
        Else
        TextBox_punkter2.Text = TextBox_punkter2.Text & ConvertPixelToCoordX(x) & ";" & ConvertPixelToCoordY(Y)
        OpdaterGraf
        Me.Repaint
        End If
    ElseIf nytmarkerpunkt Then
        If Len(TextBox_markerpunkter.Text) > 0 Then
        TextBox_markerpunkter.Text = TextBox_markerpunkter.Text & VbCrLfMac
        End If
        If Len(TextBox_ymin.Text) = 0 Or Len(TextBox_ymax.Text) = 0 Then
            MsgBox Sprog.A(301), vbOKOnly, Sprog.Error
        Else
        TextBox_markerpunkter.Text = TextBox_markerpunkter.Text & ConvertPixelToCoordX(x) & ";" & ConvertPixelToCoordY(Y)
        nytmarkerpunkt = False
        OpdaterGraf
        Me.Repaint
        End If
    Else
    Label_zoom.Left = gemx + Image1.Left
    Label_zoom.top = gemy + Image1.top
    Label_zoom.Width = 1
    Label_zoom.Height = 1
    Label_zoom.visible = True
    End If
End Sub

Private Sub Image1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
On Error Resume Next
'image1.Picture.Render(image1.Picture.Handle,0,0,600,600,image1.Picture.Width,image1.Picture.Height
'hDC, 0, 0, ScaleWidth, ScaleHeight, 0, p.Height, p.Width, -p.Height, ByVal 0
    If Label_zoom.visible Then
    Label_zoom.Left = gemx + Image1.Left
    Label_zoom.top = gemy + Image1.top
    Label_zoom.Width = x - gemx
    Label_zoom.Height = Y - gemy
    End If
End Sub

Private Sub Image1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
Dim xmin As Single
Dim xmax As Single
Dim ymin As Single
Dim ymax As Single
Dim Temp As Single
Dim s As String

Label_zoom.visible = False
If Abs(x - gemx) < 5 Then GoTo Slut

xmin = ConvertStringToNumber(TextBox_xmin.Text)
xmax = ConvertStringToNumber(TextBox_xmax.Text)
ymin = ConvertStringToNumber(TextBox_ymin.Text)
ymax = ConvertStringToNumber(TextBox_ymax.Text)

s = ConvertPixelToCoordX(gemx)
TextBox_xmax.Text = ConvertPixelToCoordX(x)
TextBox_xmin.Text = s
If TextBox_ymin.Text <> "" And TextBox_ymax.Text <> "" Then
    s = ConvertPixelToCoordY(Y)
    TextBox_ymax.Text = ConvertPixelToCoordY(gemy)
    TextBox_ymin.Text = s
End If
OpdaterGraf

    Me.Repaint
Slut:
End Sub
Private Sub Image1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim xmin As Single
Dim xmax As Single
Dim ymin As Single
Dim ymax As Single
Dim cfakt As Single
Dim dx As Single, dy As Single
Dim midt As Single
Dim nyy As Single
Dim x As Single
Dim s As String
Label_zoom.visible = False

xmin = ConvertStringToNumber(TextBox_xmin.Text)
xmax = ConvertStringToNumber(TextBox_xmax.Text)
ymin = ConvertStringToNumber(TextBox_ymin.Text)
ymax = ConvertStringToNumber(TextBox_ymax.Text)
dx = (xmax - xmin) * 0.3
dy = (ymax - ymin) * 0.3
nyy = ConvertPixelToCoordY(gemy)
'MsgBox dy, vbOKOnly, ""
'cfakt = (xmax - xmin) / (Image1.Width * 0.85)
'x = gemx - Image1.Width * 0.1
'gemx = gemx - Image1.Width * 0.1
'TextBox_xmin.text = betcif(xmin + cfakt * x - dx, 2, False)
'TextBox_xmax.text = betcif(xmin + cfakt * x + dx, 2, False)
s = ConvertPixelToCoordX(gemx) - dx
TextBox_xmax.Text = ConvertPixelToCoordX(gemx) + dx
TextBox_xmin.Text = s
If TextBox_ymin.Text <> "" And TextBox_ymax.Text <> "" Then
TextBox_ymin.Text = nyy - dy
TextBox_ymax.Text = nyy + dy
End If
OpdaterGraf

    Me.Repaint
GoTo Slut
Fejl:
    MsgBox Sprog.A(95), vbOKOnly, Sprog.Error
Slut:
    
End Sub
Function ConvertPixelToCoordX(x As Single) As Single
Dim xmin As Single, xmax As Single, cfakt As Single
xmin = ConvertStringToNumber(TextBox_xmin.Text)
xmax = ConvertStringToNumber(TextBox_xmax.Text)
cfakt = (xmax - xmin) / (Image1.Width * 0.9)
x = x - Image1.Width * 0.06
ConvertPixelToCoordX = xmin + cfakt * x
'MsgBox ConvertPixelToCoordX
End Function
Function ConvertPixelToCoordY(Y As Single) As Single
Dim ymin As Single, ymax As Single, cfakt As Single
ymin = ConvertStringToNumber(TextBox_ymin.Text)
ymax = ConvertStringToNumber(TextBox_ymax.Text)
cfakt = (ymax - ymin) / (Image1.Height * 0.9)
Y = Image1.Height - Y
Y = Y - Image1.Height * 0.08
ConvertPixelToCoordY = ymin + cfakt * Y
End Function
Private Sub CommandButton_zoom_Click()
Dim dx As Single, dy As Single
Dim midtx As Single, midty As Single
Dim xmin As Single
Dim xmax As Single
Dim ymin As Single
Dim ymax As Single
On Error GoTo Fejl
xmin = ConvertStringToNumber(TextBox_xmin.Text)
xmax = ConvertStringToNumber(TextBox_xmax.Text)
ymin = ConvertStringToNumber(TextBox_ymin.Text)
ymax = ConvertStringToNumber(TextBox_ymax.Text)

midtx = (xmax + xmin) / 2
midty = (ymax + ymin) / 2
dx = (xmax - xmin) * 0.3
dy = (ymax - ymin) * 0.3

TextBox_xmin.Text = betcif(midtx - dx, 5, False)
TextBox_xmax.Text = betcif(midtx + dx, 5, False)
If TextBox_ymin.Text <> "" And TextBox_ymax.Text <> "" Then
TextBox_ymin.Text = betcif(midty - dy, 5, False)
TextBox_ymax.Text = betcif(midty + dy, 5, False)
End If
OpdaterGraf

Me.Repaint
GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:


End Sub

Private Sub CommandButton_zoomud_Click()
Dim dx As Single, dy As Single
Dim midtx As Single, midty As Single
Dim xmin As Single
Dim xmax As Single
Dim ymin As Single
Dim ymax As Single
On Error GoTo Fejl
xmin = ConvertStringToNumber(TextBox_xmin.Text)
xmax = ConvertStringToNumber(TextBox_xmax.Text)
ymin = ConvertStringToNumber(TextBox_ymin.Text)
ymax = ConvertStringToNumber(TextBox_ymax.Text)

midtx = (xmax + xmin) / 2
midty = (ymax + ymin) / 2
dx = (xmax - xmin) * 1
dy = (ymax - ymin) * 1

TextBox_xmin.Text = betcif(midtx - dx, 5, False)
TextBox_xmax.Text = betcif(midtx + dx, 5, False)
If TextBox_ymin.Text <> "" And TextBox_ymax.Text <> "" Then
TextBox_ymin.Text = betcif(midty - dy, 5, False)
TextBox_ymax.Text = betcif(midty + dy, 5, False)
End If
OpdaterGraf
 
Me.Repaint
GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:

End Sub
Private Sub CommandButton_insertpic_Click()
Dim ils As InlineShape
Dim s As String
Dim Sep As String
On Error GoTo Fejl
#If Mac Then
#Else
    OpdaterGraf 3
#End If
If Not PicOpen Then
    If Selection.OMaths.Count > 0 Then
        omax.GoToEndOfSelectedMaths
    End If
    If Selection.Tables.Count > 0 Then
        Selection.Tables(Selection.Tables.Count).Select
        Selection.Collapse wdCollapseEnd
    End If
    Selection.MoveRight wdCharacter, 1
    Selection.TypeParagraph
End If
PicOpen = False
#If Mac Then
    Set ils = Selection.InlineShapes.AddPicture(GetTempDir() & "WordMatGraf.pdf", False, True)
#Else
    Set ils = Selection.InlineShapes.AddPicture(GetTempDir() & "WordMatGraf.gif", False, True)
#End If
Sep = "|"
s = "WordMat" & Sep & AppVersion & Sep & TextBox_definitioner.Text & Sep & TextBox_titel.Text & Sep & TextBox_xaksetitel.Text & Sep & TextBox_yaksetitel.Text & Sep
s = s & TextBox_xmin.Text & Sep & TextBox_xmax.Text & Sep & TextBox_ymin.Text & Sep & TextBox_ymax.Text & Sep
s = s & TextBox_ligning1.Text & Sep & TextBox_var1.Text & Sep & TextBox_xmin1.Text & Sep & TextBox_xmax1.Text & Sep & ComboBox_ligning1.ListIndex & Sep
s = s & TextBox_ligning2.Text & Sep & TextBox_var2.Text & Sep & TextBox_xmin2.Text & Sep & TextBox_xmax2.Text & Sep & ComboBox_ligning2.ListIndex & Sep
s = s & TextBox_ligning3.Text & Sep & TextBox_var3.Text & Sep & TextBox_xmin3.Text & Sep & TextBox_xmax3.Text & Sep & ComboBox_ligning3.ListIndex & Sep
s = s & TextBox_ligning4.Text & Sep & TextBox_var4.Text & Sep & TextBox_xmin4.Text & Sep & TextBox_xmax4.Text & Sep & ComboBox_ligning4.ListIndex & Sep
s = s & TextBox_ligning5.Text & Sep & TextBox_var5.Text & Sep & TextBox_xmin5.Text & Sep & TextBox_xmax5.Text & Sep & ComboBox_ligning5.ListIndex & Sep
s = s & TextBox_ligning6.Text & Sep & TextBox_var6.Text & Sep & TextBox_xmin6.Text & Sep & TextBox_xmax6.Text & Sep & ComboBox_ligning6.ListIndex & Sep
s = s & TextBox_lig1.Text & Sep & TextBox_lig2.Text & Sep & TextBox_Lig3.Text & Sep
s = s & TextBox_parametric1x.Text & Sep & TextBox_parametric1y.Text & Sep & TextBox_tmin1.Text & Sep & TextBox_tmax1.Text & Sep
s = s & TextBox_parametric2x.Text & Sep & TextBox_parametric2y.Text & Sep & TextBox_tmin2.Text & Sep & TextBox_tmax2.Text & Sep
s = s & TextBox_parametric3x.Text & Sep & TextBox_parametric3y.Text & Sep & TextBox_tmin3.Text & Sep & TextBox_tmax3.Text & Sep
s = s & TextBox_punkter.Text & Sep & TextBox_punkter2.Text & Sep & TextBox_markerpunkter.Text & Sep & CheckBox_pointsjoined.Value & Sep & CheckBox_pointsjoined2.Value & Sep & TextBox_pointsize.Text & Sep & TextBox_pointsize2.Text & Sep
s = s & TextBox_vektorer.Text & Sep
s = s & TextBox_labels.Text & Sep
s = s & CheckBox_gitter.Value & Sep & CheckBox_logx.Value & Sep & CheckBox_logy.Value & Sep & CheckBox_visforklaring.Value & Sep


ils.AlternativeText = s
PicOpen = False
Unload Me
GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:
Application.ScreenUpdating = True
End Sub

Sub InsertDefinitioner()
' indsætter definitioner fra textboxen i maximainputstring
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
DefString = TextBox_definitioner.Text
If Len(DefString) > 0 Then
DefString = Replace(DefString, VbCrLfMac, ListSeparator)
    DefString = TrimB(DefString, ListSeparator)
Do While InStr(DefString, ListSeparator & ListSeparator) > 0
    DefString = Replace(DefString, ListSeparator & ListSeparator, ListSeparator) ' dobbelt ;; fjernes
Loop
DefString = omax.AddDefinition("definer:" & DefString)
GetDefString = DefString
End If
End Function
Sub OpdaterDefinitioner()
' ser efter variable i textboxene og indsætter under definitioner
Dim vars As String
Dim var As String, var2 As String
Dim ea As New ExpressionAnalyser
Dim ea2 As New ExpressionAnalyser
Dim Arr As Variant
Dim arr2 As Variant
Dim i As Integer
    
    
    vars = vars & GetTextboxVars(TextBox_ligning1, TextBox_var1)
    vars = vars & GetTextboxVars(TextBox_ligning2, TextBox_var2)
    vars = vars & GetTextboxVars(TextBox_ligning3, TextBox_var3)
    vars = vars & GetTextboxVars(TextBox_ligning4, TextBox_var4)
    vars = vars & GetTextboxVars(TextBox_ligning5, TextBox_var5)
    vars = vars & GetTextboxVars(TextBox_ligning6, TextBox_var6)
    vars = vars & GetTextboxLignVars(TextBox_lig1)
    vars = vars & GetTextboxLignVars(TextBox_lig2)
    vars = vars & GetTextboxLignVars(TextBox_Lig3)
    
    
    omax.FindVariable vars, False ' fjerner dobbelte
    vars = omax.vars
    If Left(vars, 1) = ";" Then vars = right(vars, Len(vars) - 1)
    
    ea.Text = vars
    Do While right(TextBox_definitioner.Text, 2) = VbCrLfMac
        TextBox_definitioner.Text = Left(TextBox_definitioner.Text, Len(TextBox_definitioner.Text) - 2)
    Loop
    Arr = Split(TextBox_definitioner.Text, VbCrLfMac)
    
    Do
    var = ea.GetNextListItem
    var = Replace(var, VbCrLfMac, "")
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
        If Len(TextBox_definitioner.Text) > 0 Then
            TextBox_definitioner.Text = TextBox_definitioner.Text & VbCrLfMac
        End If
        TextBox_definitioner.Text = TextBox_definitioner.Text & var & "=1"
    End If
    Loop While ea.Pos <= Len(ea.Text)

    
End Sub
Function GetTextboxVars(tb As TextBox, tbvar As TextBox) As String
Dim ea As New ExpressionAnalyser
Dim var As String
    If Len(tb.Text) > 0 Then
        omax.vars = ""
        omax.FindVariable (tb.Text)
        If InStr(omax.vars, "x") > 0 Then
            var = "x"
        ElseIf InStr(omax.vars, "t") > 0 Then
            var = "t"
        Else
            ea.Text = omax.vars
            var = ea.GetNextVar(1)
        End If
        If Len(var) > 0 Then
            tbvar.Text = var
        End If
        omax.vars = RemoveVar(omax.vars, tbvar.Text)
        If Len(omax.vars) > 0 Then
            GetTextboxVars = ";" & omax.vars
        End If
    End If
End Function
Function GetTextboxLignVars(tb As TextBox) As String
Dim ea As New ExpressionAnalyser
Dim var As String
    If Len(tb.Text) > 0 Then
        omax.vars = ""
        omax.FindVariable (tb.Text)
        omax.vars = RemoveVar(omax.vars, "x")
        omax.vars = RemoveVar(omax.vars, "y")
        If Len(omax.vars) > 0 Then
            GetTextboxLignVars = ";" & omax.vars
        End If
    End If
End Function
Function RemoveVar(Text As String, var As String)
' fjerner var fra string
Dim ea As New ExpressionAnalyser

ea.Text = Text
Call ea.ReplaceVar(var, "")
Text = Replace(ea.Text, ";;", ";")
If Left(Text, 1) = ";" Then Text = right(Text, Len(Text) - 1)
If right(Text, 1) = ";" Then Text = Left(Text, Len(Text) - 1)

RemoveVar = Text
End Function

Sub opdaterLabels()
    Label_diffy.Caption = TextBox_dfy.Text & "'(" & TextBox_dfx.Text & ")="
End Sub
Private Sub CommandButton_plotdf_Click()
Dim Text As String
Dim pm As String
Dim sl As String
    Label_vent.visible = True
    omax.PrepareNewCommand finddef:=False  ' uden at søge efter definitioner i dokument
    InsertDefinitioner
    If Len(TextBox_skyd1k.Text) > 0 And Len(TextBox_skyd1f.Text) > 0 And Len(TextBox_skyd1t.Text) > 0 Then
        If Len(pm) > 0 Then pm = pm & ","
        If Len(sl) > 0 Then sl = sl & ","
        pm = pm & TextBox_skyd1k.Text & "=" & ConvertNumberToMaxima(TextBox_skyd1t.Text)
        sl = sl & TextBox_skyd1k.Text & "=" & ConvertNumberToMaxima(TextBox_skyd1f.Text) & ":" & ConvertNumberToMaxima(TextBox_skyd1t.Text)
    End If
    If Len(TextBox_skyd2k.Text) > 0 And Len(TextBox_skyd2f.Text) > 0 And Len(TextBox_skyd2t.Text) > 0 Then
        If Len(pm) > 0 Then pm = pm & ","
        If Len(sl) > 0 Then sl = sl & ","
        pm = pm & TextBox_skyd2k.Text & "=" & ConvertNumberToMaxima(TextBox_skyd2t.Text)
        sl = sl & TextBox_skyd2k.Text & "=" & ConvertNumberToMaxima(TextBox_skyd2f.Text) & ":" & ConvertNumberToMaxima(TextBox_skyd2t.Text)
    End If
    If Len(TextBox_skyd3k.Text) > 0 And Len(TextBox_skyd3f.Text) > 0 And Len(TextBox_skyd3t.Text) > 0 Then
        If Len(pm) > 0 Then pm = pm & ","
        If Len(sl) > 0 Then sl = sl & ","
        pm = pm & TextBox_skyd3k.Text & "=" & ConvertNumberToMaxima(TextBox_skyd3t.Text)
        sl = sl & TextBox_skyd3k.Text & "=" & ConvertNumberToMaxima(TextBox_skyd3f.Text) & ":" & ConvertNumberToMaxima(TextBox_skyd3t.Text)
    End If
    If Len(pm) > 0 Then
        pm = "[parameters,""" & pm & """],"
        pm = pm & "[sliders,""" & sl & """]"
    End If
    
    Call omax.PlotDF(omax.CodeForMaxima(TextBox_dfligning.Text), TextBox_dfx.Text, TextBox_dfy.Text, ConvertNumberToMaxima(TextBox_xmin.Text), ConvertNumberToMaxima(TextBox_xmax.Text), ConvertNumberToMaxima(TextBox_ymin.Text), ConvertNumberToMaxima(TextBox_ymax.Text), ConvertNumberToMaxima(TextBox_dfsol1x.Text), ConvertNumberToMaxima(TextBox_dfsol1y.Text), pm)
    Label_vent.visible = False

End Sub
Private Sub TextBox_dfx_Change()
    opdaterLabels
End Sub

Private Sub TextBox_dfy_Change()
    opdaterLabels
End Sub

Sub CheckForAssume()
' checker om der er nogle antagelser i def-textboxen og bruger dem til at lave begrænsninger på xmin og xmax
Dim DefS As String
Dim Pos As Integer
Dim ea As New ExpressionAnalyser
Dim ea2 As New ExpressionAnalyser
Dim s As String, l As String
Dim assumelist As String
ea.SetNormalBrackets
ea2.SetNormalBrackets
    DefS = GetDefString()
    TextBox_xmin1.Text = ""
    TextBox_xmin2.Text = ""
    TextBox_xmin3.Text = ""
    TextBox_xmin4.Text = ""
    TextBox_xmin5.Text = ""
    TextBox_xmin6.Text = ""
    TextBox_xmax1.Text = ""
    TextBox_xmax2.Text = ""
    TextBox_xmax3.Text = ""
    TextBox_xmax4.Text = ""
    TextBox_xmax5.Text = ""
    TextBox_xmax6.Text = ""
    
    ea.Text = DefS
    Pos = InStr(ea.Text, "assume(")
    Do While Pos > 0
        s = ea.GetNextBracketContent(Pos)
        ea2.Text = s
        l = ea2.GetNextListItem(1, ",")
        Do While Len(l) > 0
            InsertBoundary TextBox_var1.Text, l, TextBox_xmin1, TextBox_xmax1
            InsertBoundary TextBox_var2.Text, l, TextBox_xmin2, TextBox_xmax2
            InsertBoundary TextBox_var3.Text, l, TextBox_xmin3, TextBox_xmax3
            InsertBoundary TextBox_var4.Text, l, TextBox_xmin4, TextBox_xmax4
            InsertBoundary TextBox_var5.Text, l, TextBox_xmin5, TextBox_xmax5
            InsertBoundary TextBox_var6.Text, l, TextBox_xmin6, TextBox_xmax6
            l = ea2.GetNextListItem(ea2.Pos, ",")
        Loop
        Pos = InStr(Pos + 8, ea.Text, "assume(")
    Loop
    
End Sub

Sub InsertBoundary(var As String, assumetext As String, tbmin As TextBox, tbmax As TextBox)
Dim dlhs As String, drhs As String
Dim Arr As Variant
    Arr = Split(assumetext, "<")
    If UBound(Arr) > 0 Then
        dlhs = Replace(Arr(0), "=", "")
        drhs = Replace(Arr(1), "=", "")
        If dlhs = var Then
            tbmax.Text = drhs
        ElseIf drhs = var Then
            tbmin.Text = dlhs
        End If
    End If
    Arr = Split(assumetext, ">")
    If UBound(Arr) > 0 Then
        dlhs = Replace(Arr(0), "=", "")
        drhs = Replace(Arr(1), "=", "")
        If dlhs = var Then
            tbmin.Text = drhs
        ElseIf drhs = var Then
            tbmax.Text = dlhs
        End If
    End If
        
End Sub


Sub ShowPreviewMac()
#If Mac Then
    RunScript "OpenPreview", GetTempDir() & "WordMatGraf.pdf"
#End If
End Sub

Private Sub SetCaptions()
    Me.Caption = Sprog.PlotCaption
    Label_wait.Caption = Sprog.Wait
    CommandButton_opdater.Caption = Sprog.Update
    MultiPage1.Pages("Page1").Caption = Sprog.Functions
    MultiPage1.Pages("Page2").Caption = Sprog.Equations
    MultiPage1.Pages("Page3").Caption = Sprog.Points
    MultiPage1.Pages("Page5").Caption = Sprog.Vectors
    MultiPage1.Pages("Page7").Caption = Sprog.RibDirectionField
    MultiPage1.Pages("Page4").Caption = Sprog.RibSettingsShort
    Label29.Caption = Sprog.Definitions
    Label45.Caption = Sprog.Title
    Label_Ligninger.Caption = Sprog.Functions & "  f(x)=..."
    CommandButton_nulstil1.Caption = Sprog.Reset
    CommandButton_nulstil2.Caption = Sprog.Reset
    CommandButton_nulstil3.Caption = Sprog.Reset
    CommandButton_nulstil4.Caption = Sprog.Reset
    CommandButton_nulstil5.Caption = Sprog.Reset
    CommandButton_nulstil6.Caption = Sprog.Reset
    CommandButton_nulstil.Caption = Sprog.ResetAll
    CommandButton_insertpic.Caption = Sprog.A(93) 'Sprog.OK
    CommandButton_cancel.Caption = Sprog.Cancel
    Label2.Caption = "x-" & Sprog.AxisTitle
    Label3.Caption = "y-" & Sprog.AxisTitle
    Label4.Caption = Sprog.LineType
    CommandButton_nulstillign1.Caption = Sprog.Reset
    CommandButton_nulstilligning2.Caption = Sprog.Reset
    CommandButton_nulstillign3.Caption = Sprog.Reset
    Label50.Caption = Sprog.Equation & " 1"
    Label51.Caption = Sprog.Equation & " 2"
    Label52.Caption = Sprog.Equation & " 3"
    
    Label1.Caption = Sprog.A(198)
    CommandButton_kugle.Caption = Sprog.A(199) 'insert circle
    CommandButton_insertplan.Caption = Sprog.A(200) ' iinsert line
    CommandButton_parlinje.Caption = Sprog.A(200)
    CommandButton_nulstilpar1.Caption = Sprog.Reset
    Label10.Caption = Sprog.Points
    Label48.Caption = Sprog.Points & " 2"
    Label53.Caption = Sprog.A(201) ' markerede punkter
    Label28.Caption = Sprog.A(202)
    
End Sub







