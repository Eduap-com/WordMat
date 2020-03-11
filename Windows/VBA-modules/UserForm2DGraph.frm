VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2DGraph 
   Caption         =   "Plot af grafer og punkter i planen"
   ClientHeight    =   7160
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

Private Sub CommandButton_cancel_Click()
    On Error Resume Next
    PicOpen = False
    If MaxProc.Finished = 0 Then
        MaxProc.CloseProcess
        MaxProc.StartMaximaProcess
    End If
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
ctrl.text = ctrl.text & UserFormSymbol.tegn
Fejl:
End Sub


Private Sub TextBox_definitioner_AfterUpdate()
    CheckForAssume

End Sub

Private Sub TextBox_xmin_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If ToggleButton_propto.Value Then
        TextBox_ymin.text = TextBox_xmin.text
    End If
End Sub
Private Sub TextBox_xmax_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If ToggleButton_propto.Value Then
        TextBox_ymax.text = TextBox_xmax.text
    End If
End Sub
Private Sub ToggleButton_auto_Click()
If DisableEvents Then Exit Sub
DisableEvents = True
ToggleButton_propto.Value = False
ToggleButton_manuel.Value = False
TextBox_ymin.text = ""
TextBox_ymax.text = ""
DisableEvents = False
OpdaterGraf
Me.Repaint
End Sub

Private Sub ToggleButton_manuel_Click()
If DisableEvents Then Exit Sub
DisableEvents = True
ToggleButton_auto.Value = False
ToggleButton_propto.Value = False
If TextBox_ymin.text = "" Then
    TextBox_ymin.text = TextBox_xmin.text
End If
If TextBox_ymax.text = "" Then
    TextBox_ymax.text = TextBox_xmax.text
End If
DisableEvents = False
End Sub

Private Sub ToggleButton_propto_Click()
If DisableEvents Then Exit Sub
DisableEvents = True
ToggleButton_auto.Value = False
ToggleButton_manuel.Value = False
If TextBox_ymin.text = "" Then
    TextBox_ymin.text = TextBox_xmin.text
Else
End If
TextBox_ymax.text = TextBox_ymin.text + (TextBox_xmax.text - TextBox_xmin.text) * 23 / 33

DisableEvents = False
OpdaterGraf
Me.Repaint
End Sub

Private Sub UserForm_Activate()
Dim d As String, nd As String
Dim ea As New ExpressionAnalyser
Dim s As String
Dim xmin As String, xmax As String
Dim arr As Variant, i As Integer
On Error Resume Next
    SetCaptions
#If Mac Then
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
    If Len(omax.defstring) > 1 Then
    d = omax.defstringtext

    d = Replace(d, "assume", "")
    d = Replace(d, ":=", "=")
    d = Replace(d, ":", "=")
'    d = omax.ConvertToAscii(omax.ConvertToWordSymbols(d)) ' fjernet efter defstringtext anvendes
    d = Trim(d)
    
    ' reverse definition order
    arr = Split(d, "$")
    nd = arr(0)
    For i = 1 To UBound(arr)
        If Len(arr(i)) > 0 Then nd = Trim(arr(i)) & VbCrLfMac & nd
    Next
    
'    d = Replace(d, "$", vbCrLf)
    
    TextBox_definitioner.text = nd
    
    End If
End If
    OpdaterDefinitioner
    CheckForAssume
    ' indsæt xmin og xmax hvis der var definerede
    xmin = TextBox_xmin1.text
    If Len(TextBox_xmin2.text) And ConvertStringToNumber(TextBox_xmin2.text) < ConvertStringToNumber(xmin) Then xmin = TextBox_xmin2.text
    If Len(TextBox_xmin3.text) And ConvertStringToNumber(TextBox_xmin3.text) < ConvertStringToNumber(xmin) Then xmin = TextBox_xmin3.text
    If Len(TextBox_xmin4.text) And ConvertStringToNumber(TextBox_xmin4.text) < ConvertStringToNumber(xmin) Then xmin = TextBox_xmin4.text
    If Len(TextBox_xmin5.text) And ConvertStringToNumber(TextBox_xmin5.text) < ConvertStringToNumber(xmin) Then xmin = TextBox_xmin5.text
    If Len(TextBox_xmin6.text) And ConvertStringToNumber(TextBox_xmin6.text) < ConvertStringToNumber(xmin) Then xmin = TextBox_xmin6.text
    xmax = TextBox_xmax1.text
    If Len(TextBox_xmax2.text) And ConvertStringToNumber(TextBox_xmax2.text) > ConvertStringToNumber(xmax) Then xmax = TextBox_xmax2.text
    If Len(TextBox_xmax3.text) And ConvertStringToNumber(TextBox_xmax3.text) > ConvertStringToNumber(xmax) Then xmax = TextBox_xmax3.text
    If Len(TextBox_xmax4.text) And ConvertStringToNumber(TextBox_xmax4.text) > ConvertStringToNumber(xmax) Then xmax = TextBox_xmax4.text
    If Len(TextBox_xmax5.text) And ConvertStringToNumber(TextBox_xmax5.text) > ConvertStringToNumber(xmax) Then xmax = TextBox_xmax5.text
    If Len(TextBox_xmax6.text) And ConvertStringToNumber(TextBox_xmax6.text) > ConvertStringToNumber(xmax) Then xmax = TextBox_xmax6.text
    
    If Len(xmin) > 0 Then TextBox_xmin.text = xmin
    If Len(xmax) > 0 Then TextBox_xmax.text = xmax
    If Len(TextBox_xmin.text) > 0 And Len(TextBox_xmax.text) > 0 Then
        If ConvertStringToNumber(TextBox_xmin.text) > ConvertStringToNumber(TextBox_xmax.text) Then
            TextBox_xmax.text = ConvertNumberToString(ConvertStringToNumber(TextBox_xmin.text) + 10)
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
    If TextBox_lig1.text = "" Then
        TextBox_lig1.text = linje
    ElseIf TextBox_lig2.text = "" Then
        TextBox_lig2.text = linje
    ElseIf TextBox_Lig3.text = "" Then
        TextBox_Lig3.text = linje
    End If

End Sub

Private Sub CommandButton_kugle_Click()
Dim cirkel As String
    cirkel = "(x-0)^2+(y-0)^2=1^2"
    If TextBox_lig1.text = "" Then
        TextBox_lig1.text = cirkel
    ElseIf TextBox_lig2.text = "" Then
        TextBox_lig2.text = cirkel
    ElseIf TextBox_Lig3.text = "" Then
        TextBox_Lig3.text = cirkel
    End If

End Sub

Private Sub CommandButton_nulstillign1_Click()
TextBox_lig1.text = ""
End Sub

Private Sub CommandButton_nulstillign3_Click()
TextBox_Lig3.text = ""
End Sub

Private Sub CommandButton_nulstilligning2_Click()
TextBox_lig2.text = ""
End Sub

Private Sub CommandButton_nulstilpar1_Click()
TextBox_parametric1x.text = ""
TextBox_parametric1y.text = ""
TextBox_tmin1.text = ""
TextBox_tmax1.text = ""

End Sub

Private Sub CommandButton_nulstilpar2_Click()
TextBox_parametric2x.text = ""
TextBox_parametric2y.text = ""
TextBox_tmin2.text = ""
TextBox_tmax2.text = ""

End Sub

Private Sub CommandButton_nulstilpar3_Click()
TextBox_parametric3x.text = ""
TextBox_parametric3y.text = ""
TextBox_tmin3.text = ""
TextBox_tmax3.text = ""

End Sub

Private Sub CommandButton_nulstilvektorer_Click()
    TextBox_vektorer.text = ""
End Sub

Private Sub CommandButton_nyetiket_Click()

    etikettext = InputBox(Sprog.A(299), Sprog.A(298), "")
    
End Sub

Private Sub CommandButton_nytpunkt_Click()
    nytpunkt = True
End Sub

Private Sub CommandButton_nyvektor_Click()
    If TextBox_vektorer.text <> "" Then
        TextBox_vektorer.text = TextBox_vektorer.text & VbCrLfMac
    End If
    TextBox_vektorer.text = TextBox_vektorer.text & "(0;0)-(1;1)"

End Sub

Private Sub CommandButton_parlinje_Click()
Dim px As String
Dim py As String
px = "0+1*t"
py = "0+1*t"

If TextBox_parametric1x.text = "" Then
    TextBox_parametric1x.text = px
    TextBox_parametric1y.text = py
    TextBox_tmin1.text = "0"
    TextBox_tmax1.text = "1"
ElseIf TextBox_parametric2x.text = "" Then
    TextBox_parametric2x.text = px
    TextBox_parametric2y.text = py
    TextBox_tmin2.text = "0"
    TextBox_tmax2.text = "1"
ElseIf TextBox_parametric3x.text = "" Then
    TextBox_parametric3x.text = px
    TextBox_parametric3y.text = py
    TextBox_tmin3.text = "0"
    TextBox_tmax3.text = "1"
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
    TextBox_ligning1.text = ""
    TextBox_var1.text = "x"
    TextBox_xmin1.text = ""
    TextBox_xmax1.text = ""
    ComboBox_ligning1.text = ""
End Sub
Private Sub CommandButton_nulstil2_Click()
    TextBox_ligning2.text = ""
    TextBox_var2.text = "x"
    TextBox_xmin2.text = ""
    TextBox_xmax2.text = ""
    ComboBox_ligning2.text = ""
End Sub
Private Sub CommandButton_nulstil3_Click()
    TextBox_ligning3.text = ""
    TextBox_var3.text = "x"
    TextBox_xmin3.text = ""
    TextBox_xmax3.text = ""
    ComboBox_ligning3.text = ""
End Sub
Private Sub CommandButton_nulstil4_Click()
    TextBox_ligning4.text = ""
    TextBox_var4.text = "x"
    TextBox_xmin4.text = ""
    TextBox_xmax4.text = ""
    ComboBox_ligning4.text = ""
End Sub
Private Sub CommandButton_nulstil5_Click()
    TextBox_ligning5.text = ""
    TextBox_var5.text = "x"
    TextBox_xmin5.text = ""
    TextBox_xmax5.text = ""
    ComboBox_ligning5.text = ""
End Sub
Private Sub CommandButton_nulstil6_Click()
    TextBox_ligning6.text = ""
    TextBox_var6.text = "x"
    TextBox_xmin6.text = ""
    TextBox_xmax6.text = ""
    ComboBox_ligning6.text = ""
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
Dim wb As Object

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
Me.hide
    Dim ufwait2 As New UserFormWaitForMaxima
    ufwait2.Show vbModeless
    DoEvents
    ufwait2.Label_progress = "***"


If Not embed Then
cxl.LoadFile ("Graphs.xltm")
Set wb = cxl.xlwb
Set ws = cxl.xlwb.Sheets("Tabel")
    ufwait2.Label_progress = ufwait2.Label_progress & "***"

'Dim excl As Object

'Set excl = CreateObject("Excel.Application")
'Set ws = excl.ActiveWorkbook.Sheets("Tabel")


Else

path = """" & GetProgramFilesDir & "\WordMat\ExcelFiles\Graphs.xltm"""
PrepareMaximaNoSplash
omax.GoToEndOfSelectedMaths
'Selection.Collapse wdCollapseEnd
Selection.TypeParagraph
    ufwait2.Label_progress = ufwait2.Label_progress & "**"

EnableExcelMacros

Set ils = ActiveDocument.InlineShapes.AddOLEObject(FileName:=path, LinkToFile:=False, _
DisplayAsIcon:=False, Range:=Selection.Range)

'Ils.Height = 300
'Ils.Width = 500
    ufwait2.Label_progress = ufwait2.Label_progress & "***********"


'Ils.OLEFormat.DoVerb (wdOLEVerbOpen)
ils.OLEFormat.DoVerb (wdOLEVerbShow)
Set wb = ils.OLEFormat.Object
Set ws = wb.Sheets("Tabel")
'ws.Activate
DisableExcelMacros
End If

XLapp.Application.EnableEvents = False
XLapp.Application.ScreenUpdating = False
'excel.Application.EnableEvents = False
'excel.Application.ScreenUpdating = False

    ufwait2.Label_progress = ufwait2.Label_progress & "*****"
xmin = val(TextBox_xmin.text)
xmax = val(TextBox_xmax.text)
If xmin < xmax Then
    ws.Range("n3").Value = Me.TextBox_xmin.text
    ws.Range("o3").Value = Me.TextBox_xmax.text
Else
    ws.Range("n3").Value = -5
    ws.Range("o3").Value = 5
End If

    
ws.Range("b4").Value = Me.TextBox_ligning1.text
ws.Range("c4").Value = Me.TextBox_ligning2.text
ws.Range("d4").Value = Me.TextBox_ligning3.text
ws.Range("e4").Value = Me.TextBox_ligning4.text
ws.Range("f4").Value = Me.TextBox_ligning5.text
ws.Range("g4").Value = Me.TextBox_ligning6.text
'xmin og xmax kopieres over
ws.Range("B2").Value = Me.TextBox_xmin1.text
ws.Range("B3").Value = Me.TextBox_xmax1.text
ws.Range("C2").Value = Me.TextBox_xmin2.text
ws.Range("C3").Value = Me.TextBox_xmax2.text
ws.Range("D2").Value = Me.TextBox_xmin3.text
ws.Range("D3").Value = Me.TextBox_xmax3.text
ws.Range("E2").Value = Me.TextBox_xmin4.text
ws.Range("E3").Value = Me.TextBox_xmax4.text
ws.Range("F2").Value = Me.TextBox_xmin5.text
ws.Range("F3").Value = Me.TextBox_xmax5.text
ws.Range("G2").Value = Me.TextBox_xmin6.text
ws.Range("G3").Value = Me.TextBox_xmax6.text
'variabelnavn kopieres over
ws.Range("B1").Value = Me.TextBox_var1.text
ws.Range("C1").Value = Me.TextBox_var2.text
ws.Range("D1").Value = Me.TextBox_var3.text
ws.Range("E1").Value = Me.TextBox_var4.text
ws.Range("F1").Value = Me.TextBox_var5.text
ws.Range("G1").Value = Me.TextBox_var6.text
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
    If TextBox_punkter.text <> "" Then
        Dim punkttekst As String, sep As String
        punkttekst = Me.TextBox_punkter.text
        plinjer = Split(punkttekst, VbCrLfMac)
        For i = 0 To UBound(plinjer)
            If InStr(plinjer(i), ";") > 0 Then
                sep = ";"
            Else
                sep = ","
            End If
            linje = Split(plinjer(i), sep)
            If UBound(linje) > 0 Then
                ws.Range("H7").Offset(i, 0).Value = Replace(linje(0), ",", ".")
                ws.Range("I7").Offset(i, 0).Value = Replace(linje(1), ",", ".")
            End If
        Next
    End If


Slut:
On Error GoTo slut2
    ws.Range("p3").Value = Me.TextBox_ymin.text
    ws.Range("q3").Value = Me.TextBox_ymax.text
'wb.Charts(1).Activate
If TextBox_xaksetitel.text <> "" Then
    wb.Charts(1).Axes(xlCategory, xlPrimary).AxisTitle.text = Me.TextBox_xaksetitel.text
    ws.ChartObjects(1).Chart.Axes(xlCategory, xlPrimary).AxisTitle.text = Me.TextBox_xaksetitel.text
End If
If TextBox_yaksetitel.text <> "" Then
    wb.Charts(1).Axes(xlValue, xlPrimary).AxisTitle.text = Me.TextBox_yaksetitel.text
    ws.ChartObjects(1).Chart.Axes(xlValue, xlPrimary).AxisTitle.text = Me.TextBox_yaksetitel.text
End If
    ufwait2.Label_progress = ufwait2.Label_progress & "**"

'excel.Run ("UpDateAll")
XLapp.Run ("UpDateAll")
    
    ufwait2.Label_progress = ufwait2.Label_progress & "***"
wb.Charts(1).Activate
'excel.Application.EnableEvents = True
'excel.Application.ScreenUpdating = True
XLapp.Application.EnableEvents = True
XLapp.Application.ScreenUpdating = True
slut2:
    Unload ufwait2

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
Sub InsertFormula(ws As Variant, wb As Variant, tb As TextBox, col As Integer)
' indsætter formel fra textbox i kolonne col
Dim ea As New ExpressionAnalyser
    Dim varnavn As String
    Dim i As Integer
    Dim j As Integer
    Dim forskrift As String

If tb.text <> "" Then
    forskrift = tb.text
    forskrift = ConvertToExcelFormula(forskrift)
    ea.text = forskrift
    forskrift = ea.text
    
    ws.Range("b4").Offset(0, col).Value = forskrift
    
    
    'find variable
    ea.text = forskrift
    ea.pos = 1
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
        ea.pos = ea.pos + Len(varnavn) + 4
        j = 0
        Do While ws.Range("N6").Offset(j, 0).Value <> varnavn & "=" And j < i ' check om findes
          j = j + 1
        Loop
        If j = i Then
        ws.Range("N6").Offset(i, 0).Value = varnavn & "="
        ws.Range("N6").Offset(i, 1).Value = 1
        wb.Names.Add Name:="matm" & varnavn, RefersToR1C1:="=Tabel!R" & i + 6 & "C15"
        i = i + 1
        End If
        Else
            ea.pos = ea.pos + Len(varnavn)
        End If
        varnavn = ea.GetNextVar
    Loop

    On Error GoTo fejlindtast
    ea.pos = 1
    Call ea.ReplaceVar("x", "A7")
    forskrift = ea.text
     
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
Dim pos As Integer
Dim posb As Integer
Dim rod As Integer
Dim pos2 As Integer
Dim pos4 As Integer
Dim pos5 As Integer
Dim posog As Integer
Dim arr As Variant

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
    pos = InStr(forskrift, VBA.ChrW(124))
    If pos > 0 Then
        posb = InStr(pos + 1, forskrift, VBA.ChrW(124))
        forskrift = Left(forskrift, pos - 1) & "abs(" & Mid(forskrift, pos + 1, posb - pos - 1) & ")" & right(forskrift, Len(forskrift) - posb)
    End If
    Loop While pos > 0
    
    ' 3 og 4 rod
    For rod = 3 To 4
    Do
    pos = InStr(forskrift, VBA.ChrW(8728 + rod))
    If pos > 0 Or pos4 > 0 Or pos5 > 0 Then
        ea.text = forskrift
        ea.pos = pos + 1
        If Mid(forskrift, pos + 1, 1) <> "(" Then
            ea.InsertUnderstoodBracketPair
        End If
        ea.pos = pos
        Call ea.GetNextBracketContent ' bare for at finde slut parantes
        Call ea.InsertBeforePos("^(1/" & rod & ")")
        ea.text = Replace(ea.text, VBA.ChrW(8728 + rod), "", 1, 1)
        forskrift = ea.text
    End If
    Loop While pos > 0
    Next
    
    'kvadratrod
    Do
    pos = InStr(forskrift, VBA.ChrW(8730))
    If pos > 0 Then
        If Mid(forskrift, pos + 1, 1) <> "(" Then
            forskrift = Replace(forskrift, VBA.ChrW(8730), "sqrt", 1, 1)
            pos = pos + 4
            ea.text = forskrift
            ea.pos = pos
            ea.InsertUnderstoodBracketPair
            forskrift = ea.text
        Else
            ea.text = forskrift
            ea.pos = pos
            arr = Split(ea.GetNextBracketContent, "&")
            pos2 = ea.pos
            If UBound(arr) = 0 Then
                forskrift = Replace(forskrift, VBA.ChrW(8730), "sqrt", 1, 1)
            ElseIf UBound(arr) = 1 Then
                rod = arr(0)
                Call ea.InsertBeforePos("^(1/(" & rod & "))")
                ea.text = Replace(ea.text, VBA.ChrW(8730), "", 1, 1)
                posog = ea.FindChr("&", 1)
                forskrift = Left(ea.text, pos) & right(ea.text, Len(ea.text) - posog)
               
            End If
        End If
    End If
    Loop While pos > 0
    
    
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
    ea.text = forskrift
    ea.InsertBracketAfter ("^")
    ea.InsertBracketAfter ("/")
    forskrift = ea.text
    
    ' mellemrum fjernes
    forskrift = Replace(forskrift, " ", "")

    ' indsæt underforståede gangetegn ' skal være efter fjern mellem
    ea.text = forskrift
    ea.pos = 1
    ea.InsertMultSigns
    forskrift = ea.text
    
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
Dim arr As Variant
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
xming = ConvertNumberToMaxima(TextBox_xmin.text)
xmaxg = ConvertNumberToMaxima(TextBox_xmax.text)
yming = ConvertNumberToMaxima(TextBox_ymin.text)
ymaxg = ConvertNumberToMaxima(TextBox_ymax.text)

'forskrifter
If TextBox_ligning1.text <> "" Then
    lign = Replace(TextBox_ligning1.text, "'", "´")
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
    If Len(TextBox_xmin1.text) = 0 Then
        xmin = xming
    Else
        xmin = ConvertNumberToMaxima(TextBox_xmin1.text)
    End If
    If Len(TextBox_xmax1.text) = 0 Then
        xmax = xmaxg
    Else
        xmax = ConvertNumberToMaxima(TextBox_xmax1.text)
    End If
'    If Not MaximaComplex Then lign = "'CheckDef(" & lign & ",""" & TextBox_var1.text & """)"
'    If Not MaximaComplex Then lign = "'RealOnly(" & lign & ")"
    grafobj = grafobj & "color=" & GetNextColor & ",explicit(" & lign & "," & TextBox_var1.text & "," & xmin & "," & xmax & "),"
End If
If TextBox_ligning2.text <> "" Then
    lign = Replace(TextBox_ligning2.text, "'", "´")
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
    If Len(TextBox_xmin2.text) = 0 Then
        xmin = xming
    Else
        xmin = ConvertNumberToMaxima(TextBox_xmin2.text)
    End If
    If Len(TextBox_xmax2.text) = 0 Then
        xmax = xmaxg
    Else
        xmax = ConvertNumberToMaxima(TextBox_xmax2.text)
    End If
'    If Not MaximaComplex Then lign = "'CheckDef(" & lign & ",""" & TextBox_var2.text & """)"
'    If Not MaximaComplex Then lign = "'RealOnly(" & lign & ")"
    grafobj = grafobj & "color=" & GetNextColor & ",explicit(" & lign & "," & TextBox_var2.text & "," & xmin & "," & xmax & "),"
End If
If TextBox_ligning3.text <> "" Then
    lign = Replace(TextBox_ligning3.text, "'", "´")
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
    If Len(TextBox_xmin3.text) = 0 Then
        xmin = xming
    Else
        xmin = ConvertNumberToMaxima(TextBox_xmin3.text)
    End If
    If Len(TextBox_xmax3.text) = 0 Then
        xmax = xmaxg
    Else
        xmax = ConvertNumberToMaxima(TextBox_xmax3.text)
    End If
'    If Not MaximaComplex Then lign = "'RealOnly(" & lign & ")"
    grafobj = grafobj & "color=" & GetNextColor & ",explicit(" & lign & "," & TextBox_var3.text & "," & xmin & "," & xmax & "),"
End If
If TextBox_ligning4.text <> "" Then
    lign = Replace(TextBox_ligning4.text, "'", "´")
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
    If Len(TextBox_xmin4.text) = 0 Then
        xmin = xming
    Else
        xmin = ConvertNumberToMaxima(TextBox_xmin4.text)
    End If
    If Len(TextBox_xmax4.text) = 0 Then
        xmax = xmaxg
    Else
        xmax = ConvertNumberToMaxima(TextBox_xmax4.text)
    End If
'    If Not MaximaComplex Then lign = "'RealOnly(" & lign & ")"
    grafobj = grafobj & "color=" & GetNextColor & ",explicit(" & lign & "," & TextBox_var4.text & "," & xmin & "," & xmax & "),"
End If
If TextBox_ligning5.text <> "" Then
    lign = Replace(TextBox_ligning5.text, "'", "´")
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
    If Len(TextBox_xmin5.text) = 0 Then
        xmin = xming
    Else
        xmin = ConvertNumberToMaxima(TextBox_xmin5.text)
    End If
    If Len(TextBox_xmax5.text) = 0 Then
        xmax = xmaxg
    Else
        xmax = ConvertNumberToMaxima(TextBox_xmax5.text)
    End If
'    If Not MaximaComplex Then lign = "'RealOnly(" & lign & ")"
    grafobj = grafobj & "color=" & GetNextColor & ",explicit(" & lign & "," & TextBox_var5.text & "," & xmin & "," & xmax & "),"
End If
If TextBox_ligning6.text <> "" Then
    lign = Replace(TextBox_ligning6.text, "'", "´")
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
    If Len(TextBox_xmin6.text) = 0 Then
        xmin = xming
    Else
        xmin = ConvertNumberToMaxima(TextBox_xmin6.text)
    End If
    If Len(TextBox_xmax6.text) = 0 Then
        xmax = xmaxg
    Else
        xmax = ConvertNumberToMaxima(TextBox_xmax6.text)
    End If
'    If Not MaximaComplex Then lign = "'RealOnly(" & lign & ")"
    grafobj = grafobj & "color=" & GetNextColor & ",explicit(" & lign & "," & TextBox_var6.text & "," & xmin & "," & xmax & "),"
End If

'ligninger
If TextBox_lig1.text <> "" Then
    lign = Replace(TextBox_lig1.text, "'", "´")
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
If TextBox_lig2.text <> "" Then
    lign = Replace(TextBox_lig2.text, "'", "´")
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
If TextBox_Lig3.text <> "" Then
    lign = Replace(TextBox_Lig3.text, "'", "´")
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
If TextBox_parametric1x.text <> "" Then
    parx = omax.CodeForMaxima(TextBox_parametric1x.text)
    pary = omax.CodeForMaxima(TextBox_parametric1y.text)
    tmin = ConvertNumberToMaxima(TextBox_tmin1.text)
    tmax = ConvertNumberToMaxima(TextBox_tmax1.text)
    If CheckBox_visforklaring.Value Then
        grafobj = grafobj & "key=""(" & omax.ConvertToAscii(TextBox_parametric1x.text) & "," & omax.ConvertToAscii(TextBox_parametric1y.text) & ")"","
    Else
        grafobj = grafobj & "key="""","
    End If
    grafobj = grafobj & "line_type=solid,color=" & GetNextColor & ","
    grafobj = grafobj & "parametric(" & parx & "," & pary & ",t," & tmin & "," & tmax & "),"
End If
If TextBox_parametric2x.text <> "" Then
    parx = omax.CodeForMaxima(TextBox_parametric2x.text)
    pary = omax.CodeForMaxima(TextBox_parametric2y.text)
    tmin = ConvertNumberToMaxima(TextBox_tmin2.text)
    tmax = ConvertNumberToMaxima(TextBox_tmax2.text)
    If CheckBox_visforklaring.Value Then
        grafobj = grafobj & "key=""(" & omax.ConvertToAscii(TextBox_parametric2x.text) & "," & omax.ConvertToAscii(TextBox_parametric2y.text) & ")"","
    Else
        grafobj = grafobj & "key="""","
    End If
    grafobj = grafobj & "line_type=solid,color=" & GetNextColor & ","
    grafobj = grafobj & "parametric(" & parx & "," & pary & ",t," & tmin & "," & tmax & "),"
End If
If TextBox_parametric3x.text <> "" Then
    parx = omax.CodeForMaxima(TextBox_parametric3x.text)
    pary = omax.CodeForMaxima(TextBox_parametric3y.text)
    tmin = ConvertNumberToMaxima(TextBox_tmin3.text)
    tmax = ConvertNumberToMaxima(TextBox_tmax3.text)
    If CheckBox_visforklaring.Value Then
        grafobj = grafobj & "key=""(" & omax.ConvertToAscii(TextBox_parametric3x.text) & "," & omax.ConvertToAscii(TextBox_parametric3y.text) & ")"","
    Else
        grafobj = grafobj & "key="""","
    End If
    grafobj = grafobj & "line_type=solid,color=" & GetNextColor & ","
    grafobj = grafobj & "parametric(" & parx & "," & pary & ",t," & tmin & "," & tmax & "),"
End If


'punkter
If TextBox_punkter.text <> "" Then
    grafobj = grafobj & "key="""",color=black,"
    arr = Split(TextBox_punkter.text, VbCrLfMac)
    For i = 0 To UBound(arr)
    If InStr(arr(i), ";") > 0 Or InStr(arr(i), vbTab) > 0 Then
        arr(i) = Replace(arr(i), ",", ".")
        arr(i) = Replace(arr(i), ";", ",")
    End If
        arr(i) = Replace(arr(i), vbTab, ",") ' hvis tab. f.eks. hvis kopieret fra excel
        arr(i) = Replace(arr(i), " ", "")
        If Len(arr(i)) > 0 Then
        If Left(arr(i), 1) <> "(" Then
            arr(i) = "(" & arr(i)
        End If
        If right(arr(i), 1) <> ")" Then
            arr(i) = arr(i) & ")"
        End If
        arr(i) = Replace(arr(i), "),(", "],[")
        arr(i) = Replace(arr(i), ");(", "],[")
        arr(i) = Replace(arr(i), "(", "[")
        arr(i) = Replace(arr(i), ")", "]")
        punkttekst = punkttekst & arr(i) & ","
        End If
    Next
    If right(punkttekst, 1) = "," Then punkttekst = Left(punkttekst, Len(punkttekst) - 1)
    
    grafobj = grafobj & "point_type=filled_circle,point_size=" & Replace(highres * ConvertStringToNumber(TextBox_pointsize.text), ",", ".") & ",points_joined=" & VBA.LCase(CheckBox_pointsjoined.Value) & ",points([" & punkttekst & "]),"
End If

'punkter 2
If TextBox_punkter2.text <> "" Then
    punkttekst = ""
    grafobj = grafobj & "key="""",color=blue,"
    arr = Split(TextBox_punkter2.text, VbCrLfMac)
    For i = 0 To UBound(arr)
    If InStr(arr(i), ";") > 0 Or InStr(arr(i), vbTab) > 0 Then
        arr(i) = Replace(arr(i), ",", ".")
        arr(i) = Replace(arr(i), ";", ",")
    End If
        arr(i) = Replace(arr(i), vbTab, ",") ' hvis tab. f.eks. hvis kopieret fra excel
        arr(i) = Replace(arr(i), " ", "")
        If Len(arr(i)) > 0 Then
        If Left(arr(i), 1) <> "(" Then
            arr(i) = "(" & arr(i)
        End If
        If right(arr(i), 1) <> ")" Then
            arr(i) = arr(i) & ")"
        End If
        arr(i) = Replace(arr(i), "),(", "],[")
        arr(i) = Replace(arr(i), ");(", "],[")
        arr(i) = Replace(arr(i), "(", "[")
        arr(i) = Replace(arr(i), ")", "]")
        punkttekst = punkttekst & arr(i) & ","
        End If
    Next
    If right(punkttekst, 1) = "," Then punkttekst = Left(punkttekst, Len(punkttekst) - 1)
    
    grafobj = grafobj & "point_type=filled_circle,point_size=" & Replace(TextBox_pointsize2.text, ",", ".") & ",points_joined=" & VBA.LCase(CheckBox_pointsjoined2.Value) & ",points([" & punkttekst & "]),"
End If

'markerede punkter
If TextBox_markerpunkter.text <> "" Then
    punkttekst = ""
    grafobj = grafobj & "key="""",color=red,"
    arr = Split(TextBox_markerpunkter.text, VbCrLfMac)
    For i = 0 To UBound(arr)
    If InStr(arr(i), ";") > 0 Or InStr(arr(i), vbTab) > 0 Then
        arr(i) = Replace(arr(i), ",", ".")
        arr(i) = Replace(arr(i), ";", ",")
    End If
        arr(i) = Replace(arr(i), vbTab, ",") ' hvis tab. f.eks. hvis kopieret fra excel
        arr(i) = Replace(arr(i), " ", "")
        If Len(arr(i)) > 0 Then
'        If Left(arr(i), 1) = "(" Then
'            arr(i) = Right(arr(i), Len(arr(i)) - 1)
'        End If
'        If Right(arr(i), 1) = ")" Then
'            arr(i) = Left(arr(i), Len(arr(i)) - 1)
'        End If
        arr2 = Split(arr(i), ",")
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
If TextBox_vektorer.text <> "" Then
    vekt = TextBox_vektorer.text
    arr = Split(vekt, VbCrLfMac)
    For i = 0 To UBound(arr)
        If arr(i) <> "" Then
    If InStr(arr(i), ";") > 0 Then
        arr(i) = Replace(arr(i), ",", ".")
        arr(i) = Replace(arr(i), ";", ",")
    End If
    If InStr(arr(i), ")-(") > 0 Then
        arr(i) = Replace(arr(i), ")-(", "],[")
    ElseIf InStr(arr(i), ")(") > 0 Then
        arr(i) = Replace(arr(i), ")(", "],[")
    Else
        arr(i) = "[0,0]," & arr(i)
    End If
    arr(i) = Replace(arr(i), "(", "[")
    arr(i) = Replace(arr(i), ")", "]")
            If CheckBox_visforklaring.Value Then
                grafobj = grafobj & "key=""Vektor: " & arr(i) & ""","
            Else
                grafobj = grafobj & "key="""","
            End If
            grafobj = grafobj & "color=" & GetNextColor & ","
            grafobj = grafobj & "line_type=solid,line_width=" & Replace(highres, ",", ".") & ",head_angle=25,head_length=" & Replace((ConvertStringToNumber(TextBox_xmax.text) - ConvertStringToNumber(TextBox_xmin.text)) / 40, ",", ".") & ","
            grafobj = grafobj & "vector(" & arr(i) & "),"
        End If
    Next
End If

    'labels
    If Len(TextBox_labels.text) > 0 Then
        arr = Split(TextBox_labels.text, VbCrLfMac)
        For i = 0 To UBound(arr)
            If InStr(arr(i), ";") > 0 Then
                arr2 = Split(arr(i), ";")
            Else
                arr2 = Split(arr(i), ",")
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
    If Len(TextBox_xmin.text) > 0 And Len(TextBox_xmax.text) > 0 Then
        grafobj = "xrange=[" & ConvertNumberToMaxima(TextBox_xmin.text) & "," & ConvertNumberToMaxima(TextBox_xmax.text) & "]," & grafobj
    End If
    If Len(TextBox_ymin.text) > 0 And Len(TextBox_ymax.text) > 0 And Len(TextBox_dfligning.text) = 0 Then
        grafobj = "yrange=[" & ConvertNumberToMaxima(TextBox_ymin.text) & "," & ConvertNumberToMaxima(TextBox_ymax.text) & "]," & grafobj
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
        If ConvertStringToNumber(TextBox_xmin.text) > 0 Then
            grafobj = "logx=true," & grafobj
        Else
            MsgBox "xmin skal være >0 for at bruge logaritmisk x-akse."
        End If
    End If
    If CheckBox_logy.Value Then
        If ConvertStringToNumber(TextBox_ymin.text) > 0 Then
            grafobj = "logy=true," & grafobj
        Else
            MsgBox "ymin skal være >0 for at bruge logaritmisk y-akse."
        End If
    End If
    
    If Len(TextBox_titel.text) > 0 Then
        grafobj = "title=""" & omax.ConvertToAscii(TextBox_titel.text) & """," & grafobj
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
Dim text As String
Dim df As String
Dim dfsol As String
Dim pm As String
On Error GoTo Fejl
    Label_wait.Caption = Sprog.Wait & "!"
    Label_wait.Font.Size = 36
    Label_wait.visible = True
    omax.PrepareNewCommand finddef:=False  ' uden at søge efter definitioner i dokument
    InsertDefinitioner
    text = GetDraw2Dtext(highres)
    If Len(TextBox_dfligning.text) > 0 Then
        df = omax.CodeForMaxima(TextBox_dfligning.text)
        df = df & ",[" & TextBox_dfx.text & "," & TextBox_dfy.text & "]"
        If Len(TextBox_xmin.text) > 0 And Len(TextBox_ymin.text) > 0 And Len(TextBox_xmax.text) > 0 And Len(TextBox_ymax.text) > 0 Then
            df = df & ",[" & TextBox_dfx.text & "," & ConvertNumberToMaxima(TextBox_xmin.text) & "," & ConvertNumberToMaxima(TextBox_xmax.text) & "],[" & TextBox_dfy.text & "," & ConvertNumberToMaxima(TextBox_ymin.text) & "," & ConvertNumberToMaxima(TextBox_ymax.text) & "]"
        ElseIf Len(TextBox_xmin.text) > 0 And Len(TextBox_xmax.text) > 0 Then
            df = df & ",[" & TextBox_dfx.text & "," & ConvertNumberToMaxima(TextBox_xmin.text) & "," & ConvertNumberToMaxima(TextBox_xmax.text) & "]"
        Else
            df = df & ",[" & TextBox_dfx.text & "," & TextBox_dfy.text & "]"
        End If
        df = df & ",field_arrows=false"
        If Len(TextBox_dfsol1x.text) > 0 And Len(TextBox_dfsol1x.text) > 0 Then
            dfsol = dfsol & "[" & ConvertNumberToMaxima(TextBox_dfsol1x.text) & "," & ConvertNumberToMaxima(TextBox_dfsol1y.text) & "]"
        End If
        If Len(TextBox_dfsol2x.text) > 0 And Len(TextBox_dfsol2x.text) > 0 Then
            If Len(dfsol) > 0 Then dfsol = dfsol & ","
            dfsol = dfsol & "[" & ConvertNumberToMaxima(TextBox_dfsol2x.text) & "," & ConvertNumberToMaxima(TextBox_dfsol2y.text) & "]"
        End If
        If Len(TextBox_dfsol3x.text) > 0 And Len(TextBox_dfsol3x.text) > 0 Then
            If Len(dfsol) > 0 Then dfsol = dfsol & ","
            dfsol = dfsol & "[" & ConvertNumberToMaxima(TextBox_dfsol3x.text) & "," & ConvertNumberToMaxima(TextBox_dfsol3y.text) & "]"
        End If
        If Len(TextBox_dfsol4x.text) > 0 And Len(TextBox_dfsol4x.text) > 0 Then
            If Len(dfsol) > 0 Then dfsol = dfsol & ","
            dfsol = dfsol & "[" & ConvertNumberToMaxima(TextBox_dfsol4x.text) & "," & ConvertNumberToMaxima(TextBox_dfsol4y.text) & "]"
        End If
        If Len(TextBox_dfsol5x.text) > 0 And Len(TextBox_dfsol5x.text) > 0 Then
            If Len(dfsol) > 0 Then dfsol = dfsol & ","
            dfsol = dfsol & "[" & ConvertNumberToMaxima(TextBox_dfsol5x.text) & "," & ConvertNumberToMaxima(TextBox_dfsol5y.text) & "]"
        End If
        If Len(dfsol) > 0 Then
            df = df & ",duration=100,solns_at(" & dfsol & ")" ' duration defaulat er 10. ved at øge plottes længere og tættere på asymptoter
        End If
        If CheckBox_onlykurver.Value Then
            df = df & ",show_field=false"
        End If
        If text = "" Then ' must be range
            If Len(TextBox_xmin.text) > 0 And Len(TextBox_xmax.text) > 0 Then
                text = "xrange=[" & ConvertNumberToMaxima(TextBox_xmin.text) & "," & ConvertNumberToMaxima(TextBox_xmax.text) & "]"
            End If
            If Len(TextBox_ymin.text) > 0 And Len(TextBox_ymax.text) > 0 And Len(TextBox_dfligning.text) = 0 Then
                text = text & ",yrange=[" & ConvertNumberToMaxima(TextBox_ymin.text) & "," & ConvertNumberToMaxima(TextBox_ymax.text) & "]"
            End If
        End If
    End If
    If Len(text) > 0 Then
        Call omax.Draw2D(text, df, ConvertDrawLabel(TextBox_xaksetitel.text), ConvertDrawLabel(TextBox_yaksetitel.text), CheckBox_gitter.Value, True, highres)
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
Dim text As String
    omax.PrepareNewCommand finddef:=False  ' uden at søge efter definitioner i dokument
    InsertDefinitioner

    text = GetDraw2Dtext()
    
    If Len(text) > 0 Then
    Call omax.Draw2D(text, "", omax.ConvertToAscii(TextBox_xaksetitel.text), omax.ConvertToAscii(TextBox_yaksetitel.text), CheckBox_gitter.Value, False, 1)
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
    Cregr.Datatext = TextBox_punkter.text
    Cregr.ComputeLinRegr
'    Selection.Collapse
'    Selection.TypeParagraph
'    Cregr.InsertEquation

    If TextBox_ligning1.text = "" Then
        TextBox_ligning1.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning2.text = "" Then
        TextBox_ligning2.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning3.text = "" Then
        TextBox_ligning3.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning4.text = "" Then
        TextBox_ligning4.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning5.text = "" Then
        TextBox_ligning5.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning6.text = "" Then
        TextBox_ligning6.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    End If

    OpdaterGraf
    Me.Repaint
Slut:
End Sub

Private Sub CommandButton_polregr_Click()
    Dim Cregr As New CRegression
    On Error GoTo Slut
    
    Cregr.Datatext = TextBox_punkter.text
    Cregr.ComputePolRegr
'    Selection.Collapse
'    Selection.TypeParagraph
'    Cregr.InsertEquation

    If TextBox_ligning1.text = "" Then
        TextBox_ligning1.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning2.text = "" Then
        TextBox_ligning2.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning3.text = "" Then
        TextBox_ligning3.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning4.text = "" Then
        TextBox_ligning4.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning5.text = "" Then
        TextBox_ligning5.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning6.text = "" Then
        TextBox_ligning6.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    End If
    
    OpdaterGraf
    Me.Repaint
Slut:
End Sub
Private Sub CommandButton_ekspregr_Click()
   Dim Cregr As New CRegression
    On Error GoTo Slut
    
    Cregr.Datatext = TextBox_punkter.text
    Cregr.ComputeExpRegr
'    Selection.Collapse
'    Selection.TypeParagraph
'    Cregr.InsertEquation

    If TextBox_ligning1.text = "" Then
        TextBox_ligning1.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning2.text = "" Then
        TextBox_ligning2.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning3.text = "" Then
        TextBox_ligning3.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning4.text = "" Then
        TextBox_ligning4.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning5.text = "" Then
        TextBox_ligning5.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning6.text = "" Then
        TextBox_ligning6.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    End If
    
    OpdaterGraf
    Me.Repaint
Slut:
End Sub

Private Sub CommandButton_potregr_Click()
       Dim Cregr As New CRegression
    On Error GoTo Slut
    
    Cregr.Datatext = TextBox_punkter.text
    Cregr.ComputePowRegr
'    Selection.Collapse
'    Selection.TypeParagraph
'    Cregr.InsertEquation

    If TextBox_ligning1.text = "" Then
        TextBox_ligning1.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning2.text = "" Then
        TextBox_ligning2.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning3.text = "" Then
        TextBox_ligning3.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning4.text = "" Then
        TextBox_ligning4.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning5.text = "" Then
        TextBox_ligning5.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning6.text = "" Then
        TextBox_ligning6.text = right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    End If

    OpdaterGraf
    Me.Repaint
Slut:
End Sub

Private Sub CommandButton_nulstil_Click()
    TextBox_ligning1.text = ""
    TextBox_ligning2.text = ""
    TextBox_ligning3.text = ""
    TextBox_ligning4.text = ""
    TextBox_ligning5.text = ""
    TextBox_ligning6.text = ""
    TextBox_lig1.text = ""
    TextBox_lig2.text = ""
    TextBox_Lig3.text = ""
    TextBox_xmin1.text = ""
    TextBox_xmin2.text = ""
    TextBox_xmin3.text = ""
    TextBox_xmin4.text = ""
    TextBox_xmin5.text = ""
    TextBox_xmin6.text = ""
    TextBox_xmax1.text = ""
    TextBox_xmax2.text = ""
    TextBox_xmax3.text = ""
    TextBox_xmax4.text = ""
    TextBox_xmax5.text = ""
    TextBox_xmax6.text = ""
    TextBox_xmin.text = "-5"
    TextBox_xmax.text = "5"
    TextBox_ymin.text = ""
    TextBox_ymax.text = ""
    TextBox_xaksetitel.text = ""
    TextBox_yaksetitel.text = ""
    TextBox_titel.text = ""
    TextBox_punkter.text = ""
    TextBox_punkter2.text = ""
    TextBox_labels.text = ""
    TextBox_vektorer.text = ""
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
    
    If TextBox_ligning1.text <> "" Then
        If Len(MathSyntax) > 5 Then MathSyntax = MathSyntax + ","
        If InStr(TextBox_ligning1.text, "=") Then
            MathSyntax = MathSyntax & "ploteq("
        Else
            MathSyntax = MathSyntax & "plot("
        End If
        MathSyntax = MathSyntax & Replace(TextBox_ligning1.text, ",", ".")
        If TextBox_xmin1.text <> "" And TextBox_xmax1.text <> "" Then
            MathSyntax = MathSyntax + ",{x," & Me.TextBox_xmin1.text & "," & Me.TextBox_xmax1.text & "}"
        End If
'        If Me.TextBox_xmin1.text <> "" And Me.TextBox_ymax1.text <> "" Then
'            MathSyntax = MathSyntax + ",{y," & Me.TextBox_ymin1.text & "," & Me.TextBox_ymax1.text & "}"
'        End If
        If ComboBox_ligning1.text = "---" Then
        ElseIf ComboBox_ligning1.text = "..." Then
            MathSyntax = MathSyntax + ",{LineStyle,Dot}"
        ElseIf ComboBox_ligning1.text = "- - -" Then
            MathSyntax = MathSyntax + ",{LineStyle,Dash}"
        ElseIf ComboBox_ligning1.text = "-.-.-." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDot}"
        ElseIf ComboBox_ligning1.text = "- . . - . ." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDotDot}"
        End If
'        If TextBox_samplesize.text <> "" Then
'            MathSyntax = MathSyntax + ",{SampleSize," & TextBox_samplesize.text & "}"
'        End If
        MathSyntax = MathSyntax + ")"
    End If
    If TextBox_ligning2.text <> "" Then
        If Len(MathSyntax) > 5 Then MathSyntax = MathSyntax + ","
        If InStr(TextBox_ligning2.text, "=") Then
            MathSyntax = MathSyntax & "ploteq("
        Else
            MathSyntax = MathSyntax & "plot("
        End If
        MathSyntax = MathSyntax & Replace(TextBox_ligning2.text, ",", ".")
        If TextBox_xmin2.text <> "" And TextBox_xmax2.text <> "" Then
            MathSyntax = MathSyntax + ",{x," & TextBox_xmin2.text & "," & TextBox_xmax2.text & "}"
        End If
'        If TextBox_ymin2.text <> "" And TextBox_ymax2.text <> "" Then
'            MathSyntax = MathSyntax + ",{y," & TextBox_ymin2.text & "," & TextBox_ymax2.text & "}"
'        End If
        If ComboBox_ligning2.text = "---" Then
        ElseIf ComboBox_ligning2.text = "..." Then
            MathSyntax = MathSyntax + ",{LineStyle,Dot}"
        ElseIf ComboBox_ligning2.text = "- - -" Then
            MathSyntax = MathSyntax + ",{LineStyle,Dash}"
        ElseIf ComboBox_ligning2.text = "-.-.-." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDot}"
        ElseIf ComboBox_ligning2.text = "- . . - . ." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDotDot}"
        End If
'        If TextBox_samplesize.text <> "" Then
'            MathSyntax = MathSyntax + ",{SampleSize," & TextBox_samplesize.text & "}"
'        End If
        MathSyntax = MathSyntax + ")"
    End If
    If TextBox_ligning3.text <> "" Then
        If Len(MathSyntax) > 5 Then MathSyntax = MathSyntax + ","
        If InStr(TextBox_ligning3.text, "=") Then
            MathSyntax = MathSyntax & "ploteq("
        Else
            MathSyntax = MathSyntax & "plot("
        End If
        MathSyntax = MathSyntax & Replace(TextBox_ligning3.text, ",", ".")
        If TextBox_xmin3.text <> "" And TextBox_xmax3.text <> "" Then
            MathSyntax = MathSyntax + ",{x," & TextBox_xmin3.text & "," & TextBox_xmax3.text & "}"
        End If
'        If TextBox_ymin3.text <> "" And TextBox_ymax3.text <> "" Then
'            MathSyntax = MathSyntax + ",{y," & TextBox_ymin3.text & "," & TextBox_ymax3.text & "}"
'        End If
        If ComboBox_ligning3.text = "---" Then
        ElseIf ComboBox_ligning3.text = "..." Then
            MathSyntax = MathSyntax + ",{LineStyle,Dot}"
        ElseIf ComboBox_ligning3.text = "- - -" Then
            MathSyntax = MathSyntax + ",{LineStyle,Dash}"
        ElseIf ComboBox_ligning3.text = "-.-.-." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDot}"
        ElseIf ComboBox_ligning3.text = "- . . - . ." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDotDot}"
        End If
'        If TextBox_samplesize.text <> "" Then
'            MathSyntax = MathSyntax + ",{SampleSize," & TextBox_samplesize.text & "}"
'        End If
        MathSyntax = MathSyntax + ")"
    End If
    If TextBox_ligning4.text <> "" Then
        If Len(MathSyntax) > 5 Then MathSyntax = MathSyntax + ","
        If InStr(TextBox_ligning4.text, "=") Then
            MathSyntax = MathSyntax & "ploteq("
        Else
            MathSyntax = MathSyntax & "plot("
        End If
        MathSyntax = MathSyntax & Replace(TextBox_ligning4.text, ",", ".")
        If TextBox_xmin4.text <> "" And TextBox_xmax4.text <> "" Then
            MathSyntax = MathSyntax + ",{x," & TextBox_xmin4.text & "," & TextBox_xmax4.text & "}"
        End If
'        If TextBox_ymin4.text <> "" And TextBox_ymax4.text <> "" Then
'            MathSyntax = MathSyntax + ",{y," & TextBox_ymin4.text & "," & TextBox_ymax4.text & "}"
'        End If
        If ComboBox_ligning4.text = "---" Then
        ElseIf ComboBox_ligning4.text = "..." Then
            MathSyntax = MathSyntax + ",{LineStyle,Dot}"
        ElseIf ComboBox_ligning4.text = "- - -" Then
            MathSyntax = MathSyntax + ",{LineStyle,Dash}"
        ElseIf ComboBox_ligning4.text = "-.-.-." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDot}"
        ElseIf ComboBox_ligning4.text = "- . . - . ." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDotDot}"
        End If
'        If TextBox_samplesize.text <> "" Then
'            MathSyntax = MathSyntax + ",{SampleSize," & TextBox_samplesize.text & "}"
'        End If
        MathSyntax = MathSyntax + ")"
    End If
    If TextBox_ligning5.text <> "" Then
        If Len(MathSyntax) > 5 Then MathSyntax = MathSyntax + ","
        If InStr(TextBox_ligning5.text, "=") Then
            MathSyntax = MathSyntax & "ploteq("
        Else
            MathSyntax = MathSyntax & "plot("
        End If
        MathSyntax = MathSyntax & Replace(TextBox_ligning5.text, ",", ".")
        If TextBox_xmin5.text <> "" And TextBox_xmax5.text <> "" Then
            MathSyntax = MathSyntax + ",{x," & TextBox_xmin5.text & "," & TextBox_xmax5.text & "}"
        End If
'        If TextBox_ymin5.text <> "" And TextBox_ymax5.text <> "" Then
'            MathSyntax = MathSyntax + ",{y," & TextBox_ymin5.text & "," & TextBox_ymax5.text & "}"
'        End If
        If ComboBox_ligning5.text = "---" Then
        ElseIf ComboBox_ligning5.text = "..." Then
            MathSyntax = MathSyntax + ",{LineStyle,Dot}"
        ElseIf ComboBox_ligning5.text = "- - -" Then
            MathSyntax = MathSyntax + ",{LineStyle,Dash}"
        ElseIf ComboBox_ligning5.text = "-.-.-." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDot}"
        ElseIf ComboBox_ligning5.text = "- . . - . ." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDotDot}"
        End If
        MathSyntax = MathSyntax + ")"
    End If
    If TextBox_ligning6.text <> "" Then
        If Len(MathSyntax) > 5 Then MathSyntax = MathSyntax + ","
        If InStr(TextBox_ligning6.text, "=") Then
            MathSyntax = MathSyntax & "ploteq("
        Else
            MathSyntax = MathSyntax & "plot("
        End If
        MathSyntax = MathSyntax & Replace(TextBox_ligning6.text, ",", ".")
        If TextBox_xmin6.text <> "" And TextBox_xmax6.text <> "" Then
            MathSyntax = MathSyntax + ",{x," & TextBox_xmin6.text & "," & TextBox_xmax6.text & "}"
        End If
        If ComboBox_ligning6.text = "---" Then
        ElseIf ComboBox_ligning6.text = "..." Then
            MathSyntax = MathSyntax + ",{LineStyle,Dot}"
        ElseIf ComboBox_ligning6.text = "- - -" Then
            MathSyntax = MathSyntax + ",{LineStyle,Dash}"
        ElseIf ComboBox_ligning6.text = "-.-.-." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDot}"
        ElseIf ComboBox_ligning6.text = "- . . - . ." Then
            MathSyntax = MathSyntax + ",{LineStyle,DashDotDot}"
        End If
        MathSyntax = MathSyntax + ")"
    End If
    
    'datapunkter
    If TextBox_punkter.text <> "" Then
        If Len(MathSyntax) > 5 Then MathSyntax = MathSyntax + ","
        MathSyntax = MathSyntax + "plotdataset({"
        Dim punkttekst As String
        punkttekst = TextBox_punkter.text
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
    If TextBox_xmin.text <> "" And TextBox_xmax.text <> "" Then
        MathSyntax = MathSyntax + ",{x," & TextBox_xmin.text & "," & TextBox_xmax.text & "}"
    End If
    ' ymin og ymax
    If TextBox_ymin.text <> "" And TextBox_ymax.text <> "" Then
        MathSyntax = MathSyntax + ",{y," & TextBox_ymin.text & "," & TextBox_ymax.text & "}"
    End If
    
    If TextBox_xaksetitel.text <> "" Then
            MathSyntax = MathSyntax + ",{aliasX," & TextBox_xaksetitel.text & "}"
    End If
    If TextBox_yaksetitel.text <> "" Then
            MathSyntax = MathSyntax + ",{aliasY," & TextBox_yaksetitel.text & "}"
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
     Selection.TypeText text:=MathSyntax
    Selection.OMaths.BuildUp
'    Set objEq = objRange.OMaths(1)
'    objEq.BuildUp
    Selection.TypeParagraph

    hide
    
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

Function ConvertDegreeToRad(text As String, trigfunc As String) As String
    Dim pos, spos As Integer
    Dim ea As New ExpressionAnalyser
    ea.StartBracket = "("
    ea.EndBracket = ")"
    ea.text = text
    spos = 1
    
    Do
    pos = ea.FindChr("arc" & trigfunc, spos)
    If pos > 0 Then
        ea.GetNextBracketContent
        ea.InsertBeforePos (")")
        ea.pos = pos
        ea.InsertBeforePos ("180/PI()*(")
        spos = pos + 13
    End If
    Loop While pos > 0
    
    spos = 1
    Do
    pos = ea.FindChr(trigfunc, spos)
    If pos > 0 Then
        If Not (ea.ChrByIndex(pos - 1) = "a") Then
        ea.GetNextBracketContent
        ea.InsertBeforePos (")")
        ea.pos = pos + Len(trigfunc)
        ea.InsertAfterPos ("PI()/180*(")
        spos = pos + 13
        Else
            spos = pos + 3
        End If
    End If
    Loop While pos > 0
    
    ConvertDegreeToRad = ea.text

End Function

Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    gemx = x
    gemy = Y
    If Len(etikettext) > 0 Then
        If Len(TextBox_labels.text) > 0 Then
        TextBox_labels.text = TextBox_labels.text & VbCrLfMac
        End If
        If Len(TextBox_ymin.text) = 0 Or Len(TextBox_ymax.text) = 0 Then
            MsgBox Sprog.A(301), vbOKOnly, Sprog.Error
        Else
        TextBox_labels.text = TextBox_labels.text & etikettext & ";" & ConvertPixelToCoordX(x) & ";" & ConvertPixelToCoordY(Y)
        etikettext = ""
        OpdaterGraf
        Me.Repaint
        End If
    ElseIf nytpunkt Then
        nytpunkt = False
        If Len(TextBox_punkter2.text) > 0 Then
        TextBox_punkter2.text = TextBox_punkter2.text & VbCrLfMac
        End If
        If Len(TextBox_ymin.text) = 0 Or Len(TextBox_ymax.text) = 0 Then
            MsgBox Sprog.A(301), vbOKOnly, Sprog.Error
        Else
        TextBox_punkter2.text = TextBox_punkter2.text & ConvertPixelToCoordX(x) & ";" & ConvertPixelToCoordY(Y)
        OpdaterGraf
        Me.Repaint
        End If
    ElseIf nytmarkerpunkt Then
        If Len(TextBox_markerpunkter.text) > 0 Then
        TextBox_markerpunkter.text = TextBox_markerpunkter.text & VbCrLfMac
        End If
        If Len(TextBox_ymin.text) = 0 Or Len(TextBox_ymax.text) = 0 Then
            MsgBox Sprog.A(301), vbOKOnly, Sprog.Error
        Else
        TextBox_markerpunkter.text = TextBox_markerpunkter.text & ConvertPixelToCoordX(x) & ";" & ConvertPixelToCoordY(Y)
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

xmin = ConvertStringToNumber(TextBox_xmin.text)
xmax = ConvertStringToNumber(TextBox_xmax.text)
ymin = ConvertStringToNumber(TextBox_ymin.text)
ymax = ConvertStringToNumber(TextBox_ymax.text)

s = ConvertPixelToCoordX(gemx)
TextBox_xmax.text = ConvertPixelToCoordX(x)
TextBox_xmin.text = s
If TextBox_ymin.text <> "" And TextBox_ymax.text <> "" Then
    s = ConvertPixelToCoordY(Y)
    TextBox_ymax.text = ConvertPixelToCoordY(gemy)
    TextBox_ymin.text = s
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

xmin = ConvertStringToNumber(TextBox_xmin.text)
xmax = ConvertStringToNumber(TextBox_xmax.text)
ymin = ConvertStringToNumber(TextBox_ymin.text)
ymax = ConvertStringToNumber(TextBox_ymax.text)
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
TextBox_xmax.text = ConvertPixelToCoordX(gemx) + dx
TextBox_xmin.text = s
If TextBox_ymin.text <> "" And TextBox_ymax.text <> "" Then
TextBox_ymin.text = nyy - dy
TextBox_ymax.text = nyy + dy
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
xmin = ConvertStringToNumber(TextBox_xmin.text)
xmax = ConvertStringToNumber(TextBox_xmax.text)
cfakt = (xmax - xmin) / (Image1.Width * 0.9)
x = x - Image1.Width * 0.06
ConvertPixelToCoordX = xmin + cfakt * x
'MsgBox ConvertPixelToCoordX
End Function
Function ConvertPixelToCoordY(Y As Single) As Single
Dim ymin As Single, ymax As Single, cfakt As Single
ymin = ConvertStringToNumber(TextBox_ymin.text)
ymax = ConvertStringToNumber(TextBox_ymax.text)
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
xmin = ConvertStringToNumber(TextBox_xmin.text)
xmax = ConvertStringToNumber(TextBox_xmax.text)
ymin = ConvertStringToNumber(TextBox_ymin.text)
ymax = ConvertStringToNumber(TextBox_ymax.text)

midtx = (xmax + xmin) / 2
midty = (ymax + ymin) / 2
dx = (xmax - xmin) * 0.3
dy = (ymax - ymin) * 0.3

TextBox_xmin.text = betcif(midtx - dx, 5, False)
TextBox_xmax.text = betcif(midtx + dx, 5, False)
If TextBox_ymin.text <> "" And TextBox_ymax.text <> "" Then
TextBox_ymin.text = betcif(midty - dy, 5, False)
TextBox_ymax.text = betcif(midty + dy, 5, False)
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
xmin = ConvertStringToNumber(TextBox_xmin.text)
xmax = ConvertStringToNumber(TextBox_xmax.text)
ymin = ConvertStringToNumber(TextBox_ymin.text)
ymax = ConvertStringToNumber(TextBox_ymax.text)

midtx = (xmax + xmin) / 2
midty = (ymax + ymin) / 2
dx = (xmax - xmin) * 1
dy = (ymax - ymin) * 1

TextBox_xmin.text = betcif(midtx - dx, 5, False)
TextBox_xmax.text = betcif(midtx + dx, 5, False)
If TextBox_ymin.text <> "" And TextBox_ymax.text <> "" Then
TextBox_ymin.text = betcif(midty - dy, 5, False)
TextBox_ymax.text = betcif(midty + dy, 5, False)
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
Dim sep As String
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
sep = "|"
s = "WordMat" & sep & AppVersion & sep & TextBox_definitioner.text & sep & TextBox_titel.text & sep & TextBox_xaksetitel.text & sep & TextBox_yaksetitel.text & sep
s = s & TextBox_xmin.text & sep & TextBox_xmax.text & sep & TextBox_ymin.text & sep & TextBox_ymax.text & sep
s = s & TextBox_ligning1.text & sep & TextBox_var1.text & sep & TextBox_xmin1.text & sep & TextBox_xmax1.text & sep & ComboBox_ligning1.ListIndex & sep
s = s & TextBox_ligning2.text & sep & TextBox_var2.text & sep & TextBox_xmin2.text & sep & TextBox_xmax2.text & sep & ComboBox_ligning2.ListIndex & sep
s = s & TextBox_ligning3.text & sep & TextBox_var3.text & sep & TextBox_xmin3.text & sep & TextBox_xmax3.text & sep & ComboBox_ligning3.ListIndex & sep
s = s & TextBox_ligning4.text & sep & TextBox_var4.text & sep & TextBox_xmin4.text & sep & TextBox_xmax4.text & sep & ComboBox_ligning4.ListIndex & sep
s = s & TextBox_ligning5.text & sep & TextBox_var5.text & sep & TextBox_xmin5.text & sep & TextBox_xmax5.text & sep & ComboBox_ligning5.ListIndex & sep
s = s & TextBox_ligning6.text & sep & TextBox_var6.text & sep & TextBox_xmin6.text & sep & TextBox_xmax6.text & sep & ComboBox_ligning6.ListIndex & sep
s = s & TextBox_lig1.text & sep & TextBox_lig2.text & sep & TextBox_Lig3.text & sep
s = s & TextBox_parametric1x.text & sep & TextBox_parametric1y.text & sep & TextBox_tmin1.text & sep & TextBox_tmax1.text & sep
s = s & TextBox_parametric2x.text & sep & TextBox_parametric2y.text & sep & TextBox_tmin2.text & sep & TextBox_tmax2.text & sep
s = s & TextBox_parametric3x.text & sep & TextBox_parametric3y.text & sep & TextBox_tmin3.text & sep & TextBox_tmax3.text & sep
s = s & TextBox_punkter.text & sep & TextBox_punkter2.text & sep & TextBox_markerpunkter.text & sep & CheckBox_pointsjoined.Value & sep & CheckBox_pointsjoined2.Value & sep & TextBox_pointsize.text & sep & TextBox_pointsize2.text & sep
s = s & TextBox_vektorer.text & sep
s = s & TextBox_labels.text & sep
s = s & CheckBox_gitter.Value & sep & CheckBox_logx.Value & sep & CheckBox_logy.Value & sep & CheckBox_visforklaring.Value & sep


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
Dim defstring As String

omax.InsertKillDef

defstring = GetDefString

If Len(defstring) > 0 Then
'defstring = Replace(defstring, ",", ".")
'defstring = Replace(defstring, ";", ",")
'defstring = Replace(defstring, "=", ":")
If right(defstring, 1) = "," Then defstring = Left(defstring, Len(defstring) - 1)

'omax.MaximaInputStreng = omax.MaximaInputStreng & "[" & defstring & "]$"
omax.MaximaInputStreng = omax.MaximaInputStreng & defstring
End If
End Sub
Function GetDefString()
Dim defstring As String
omax.ResetDefinitions
defstring = TextBox_definitioner.text
If Len(defstring) > 0 Then
defstring = Replace(defstring, VbCrLfMac, ListSeparator)
    defstring = TrimB(defstring, ListSeparator)
Do While InStr(defstring, ListSeparator & ListSeparator) > 0
    defstring = Replace(defstring, ListSeparator & ListSeparator, ListSeparator) ' dobbelt ;; fjernes
Loop
defstring = omax.AddDefinition("definer:" & defstring)
GetDefString = defstring
End If
End Function
Sub OpdaterDefinitioner()
' ser efter variable i textboxene og indsætter under definitioner
Dim vars As String
Dim var As String, var2 As String
Dim ea As New ExpressionAnalyser
Dim ea2 As New ExpressionAnalyser
Dim arr As Variant
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
    
    ea.text = vars
    Do While right(TextBox_definitioner.text, 2) = VbCrLfMac
        TextBox_definitioner.text = Left(TextBox_definitioner.text, Len(TextBox_definitioner.text) - 2)
    Loop
    arr = Split(TextBox_definitioner.text, VbCrLfMac)
    
    Do
    var = ea.GetNextListItem
    var = Replace(var, VbCrLfMac, "")
    For i = 0 To UBound(arr)
        If arr(i) <> "" Then
        var2 = Split(arr(i), "=")(0)
        If var2 = var Then
            var = ""
            Exit For
        End If
        End If
    Next
    If var <> "" Then
'        If Right(TextBox_definitioner.text, 2) <> vbCrLf Then
        If Len(TextBox_definitioner.text) > 0 Then
            TextBox_definitioner.text = TextBox_definitioner.text & VbCrLfMac
        End If
        TextBox_definitioner.text = TextBox_definitioner.text & var & "=1"
    End If
    Loop While ea.pos <= Len(ea.text)

    
End Sub
Function GetTextboxVars(tb As TextBox, tbvar As TextBox) As String
Dim ea As New ExpressionAnalyser
Dim var As String
    If Len(tb.text) > 0 Then
        omax.vars = ""
        omax.FindVariable (tb.text)
        If InStr(omax.vars, "x") > 0 Then
            var = "x"
        ElseIf InStr(omax.vars, "t") > 0 Then
            var = "t"
        Else
            ea.text = omax.vars
            var = ea.GetNextVar(1)
        End If
        If Len(var) > 0 Then
            tbvar.text = var
        End If
        omax.vars = RemoveVar(omax.vars, tbvar.text)
        If Len(omax.vars) > 0 Then
            GetTextboxVars = ";" & omax.vars
        End If
    End If
End Function
Function GetTextboxLignVars(tb As TextBox) As String
Dim ea As New ExpressionAnalyser
Dim var As String
    If Len(tb.text) > 0 Then
        omax.vars = ""
        omax.FindVariable (tb.text)
        omax.vars = RemoveVar(omax.vars, "x")
        omax.vars = RemoveVar(omax.vars, "y")
        If Len(omax.vars) > 0 Then
            GetTextboxLignVars = ";" & omax.vars
        End If
    End If
End Function
Function RemoveVar(text As String, var As String)
' fjerner var fra string
Dim ea As New ExpressionAnalyser

ea.text = text
Call ea.ReplaceVar(var, "")
text = Replace(ea.text, ";;", ";")
If Left(text, 1) = ";" Then text = right(text, Len(text) - 1)
If right(text, 1) = ";" Then text = Left(text, Len(text) - 1)

RemoveVar = text
End Function

Sub opdaterLabels()
    Label_diffy.Caption = TextBox_dfy.text & "'(" & TextBox_dfx.text & ")="
End Sub
Private Sub CommandButton_plotdf_Click()
Dim text As String
Dim pm As String
Dim sl As String
    Label_vent.visible = True
    omax.PrepareNewCommand finddef:=False  ' uden at søge efter definitioner i dokument
    InsertDefinitioner
    If Len(TextBox_skyd1k.text) > 0 And Len(TextBox_skyd1f.text) > 0 And Len(TextBox_skyd1t.text) > 0 Then
        If Len(pm) > 0 Then pm = pm & ","
        If Len(sl) > 0 Then sl = sl & ","
        pm = pm & TextBox_skyd1k.text & "=" & ConvertNumberToMaxima(TextBox_skyd1t.text)
        sl = sl & TextBox_skyd1k.text & "=" & ConvertNumberToMaxima(TextBox_skyd1f.text) & ":" & ConvertNumberToMaxima(TextBox_skyd1t.text)
    End If
    If Len(TextBox_skyd2k.text) > 0 And Len(TextBox_skyd2f.text) > 0 And Len(TextBox_skyd2t.text) > 0 Then
        If Len(pm) > 0 Then pm = pm & ","
        If Len(sl) > 0 Then sl = sl & ","
        pm = pm & TextBox_skyd2k.text & "=" & ConvertNumberToMaxima(TextBox_skyd2t.text)
        sl = sl & TextBox_skyd2k.text & "=" & ConvertNumberToMaxima(TextBox_skyd2f.text) & ":" & ConvertNumberToMaxima(TextBox_skyd2t.text)
    End If
    If Len(TextBox_skyd3k.text) > 0 And Len(TextBox_skyd3f.text) > 0 And Len(TextBox_skyd3t.text) > 0 Then
        If Len(pm) > 0 Then pm = pm & ","
        If Len(sl) > 0 Then sl = sl & ","
        pm = pm & TextBox_skyd3k.text & "=" & ConvertNumberToMaxima(TextBox_skyd3t.text)
        sl = sl & TextBox_skyd3k.text & "=" & ConvertNumberToMaxima(TextBox_skyd3f.text) & ":" & ConvertNumberToMaxima(TextBox_skyd3t.text)
    End If
    If Len(pm) > 0 Then
        pm = "[parameters,""" & pm & """],"
        pm = pm & "[sliders,""" & sl & """]"
    End If
    
    Call omax.PlotDF(omax.CodeForMaxima(TextBox_dfligning.text), TextBox_dfx.text, TextBox_dfy.text, ConvertNumberToMaxima(TextBox_xmin.text), ConvertNumberToMaxima(TextBox_xmax.text), ConvertNumberToMaxima(TextBox_ymin.text), ConvertNumberToMaxima(TextBox_ymax.text), ConvertNumberToMaxima(TextBox_dfsol1x.text), ConvertNumberToMaxima(TextBox_dfsol1y.text), pm)
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
Dim defs As String
Dim pos As Integer
Dim ea As New ExpressionAnalyser
Dim ea2 As New ExpressionAnalyser
Dim s As String, l As String
Dim assumelist As String
ea.SetNormalBrackets
ea2.SetNormalBrackets
    defs = GetDefString()
    TextBox_xmin1.text = ""
    TextBox_xmin2.text = ""
    TextBox_xmin3.text = ""
    TextBox_xmin4.text = ""
    TextBox_xmin5.text = ""
    TextBox_xmin6.text = ""
    TextBox_xmax1.text = ""
    TextBox_xmax2.text = ""
    TextBox_xmax3.text = ""
    TextBox_xmax4.text = ""
    TextBox_xmax5.text = ""
    TextBox_xmax6.text = ""
    
    ea.text = defs
    pos = InStr(ea.text, "assume(")
    Do While pos > 0
        s = ea.GetNextBracketContent(pos)
        ea2.text = s
        l = ea2.GetNextListItem(1, ",")
        Do While Len(l) > 0
            InsertBoundary TextBox_var1.text, l, TextBox_xmin1, TextBox_xmax1
            InsertBoundary TextBox_var2.text, l, TextBox_xmin2, TextBox_xmax2
            InsertBoundary TextBox_var3.text, l, TextBox_xmin3, TextBox_xmax3
            InsertBoundary TextBox_var4.text, l, TextBox_xmin4, TextBox_xmax4
            InsertBoundary TextBox_var5.text, l, TextBox_xmin5, TextBox_xmax5
            InsertBoundary TextBox_var6.text, l, TextBox_xmin6, TextBox_xmax6
            l = ea2.GetNextListItem(ea2.pos, ",")
        Loop
        pos = InStr(pos + 8, ea.text, "assume(")
    Loop
    
End Sub

Sub InsertBoundary(var As String, assumetext As String, tbmin As TextBox, tbmax As TextBox)
Dim dlhs As String, drhs As String
Dim arr As Variant
    arr = Split(assumetext, "<")
    If UBound(arr) > 0 Then
        dlhs = Replace(arr(0), "=", "")
        drhs = Replace(arr(1), "=", "")
        If dlhs = var Then
            tbmax.text = drhs
        ElseIf drhs = var Then
            tbmin.text = dlhs
        End If
    End If
    arr = Split(assumetext, ">")
    If UBound(arr) > 0 Then
        dlhs = Replace(arr(0), "=", "")
        drhs = Replace(arr(1), "=", "")
        If dlhs = var Then
            tbmin.text = drhs
        ElseIf drhs = var Then
            tbmax.text = dlhs
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







