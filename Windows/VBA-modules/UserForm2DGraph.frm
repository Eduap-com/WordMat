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
' only used for plotting with GnuPlot

Private gemx As Single
Private gemy As Single
Private embed As Boolean
Private etikettext As String
Private nytpunkt As Boolean
Private nytmarkerpunkt As Boolean
Private DisableEvents As Boolean
Private Opened As Boolean

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
Private Sub Label_cancel_Click()
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

Private Sub CommandButton_GeoGebraDF_Click()
    Dim s As String, Fundet As Boolean
    s = "SlopeField(" & TextBox_dfligning.Text & ");Xmin=-100;Xmax=100;Tic=0.1;"
    If TextBox_dfsol1x.Text <> vbNullString And TextBox_dfsol1y.Text <> vbNullString Then
        s = s & "A=(" & TextBox_dfsol1x.Text & ", " & TextBox_dfsol1y.Text & ");"
        s = s & "SolveODE(" & TextBox_dfligning.Text & ", x(A), y(A), Xmin, Tic);" ' y(A) virker ikke
        s = s & "SolveODE(" & TextBox_dfligning.Text & ", x(A), y(A), Xmax, Tic);" ' y(A) virker ikke
        Fundet = True
    End If
    If TextBox_dfsol2x.Text <> vbNullString And TextBox_dfsol2y.Text <> vbNullString Then
        s = s & "B=(" & TextBox_dfsol2x.Text & ", " & TextBox_dfsol2y.Text & ");"
        s = s & "SolveODE(" & TextBox_dfligning.Text & ", x(B), y(B), Xmin, Tic);" ' y(A) virker ikke
        s = s & "SolveODE(" & TextBox_dfligning.Text & ", x(B), y(B), Xmax, Tic);" ' y(A) virker ikke
        Fundet = True
    End If
    If TextBox_dfsol3x.Text <> vbNullString And TextBox_dfsol3y.Text <> vbNullString Then
        s = s & "C=(" & TextBox_dfsol3x.Text & ", " & TextBox_dfsol3y.Text & ");"
        s = s & "SolveODE(" & TextBox_dfligning.Text & ", x(C), y(C), Xmin, Tic);"
        s = s & "SolveODE(" & TextBox_dfligning.Text & ", x(C), y(C), Xmax, Tic);"
        Fundet = True
    End If
    If TextBox_dfsol4x.Text <> vbNullString And TextBox_dfsol4y.Text <> vbNullString Then
        s = s & "D=(" & TextBox_dfsol4x.Text & ", " & TextBox_dfsol4y.Text & ");"
        s = s & "SolveODE(" & TextBox_dfligning.Text & ", x(D), y(D), Xmin, Tic);"
        s = s & "SolveODE(" & TextBox_dfligning.Text & ", x(D), y(D), Xmax, Tic);"
        Fundet = True
    End If
    If TextBox_dfsol5x.Text <> vbNullString And TextBox_dfsol5y.Text <> vbNullString Then
        s = s & "E=(" & TextBox_dfsol5x.Text & ", " & TextBox_dfsol5y.Text & ");"
        s = s & "SolveODE(" & TextBox_dfligning.Text & ", x(E), y(E), Xmin, Tic);"
        s = s & "SolveODE(" & TextBox_dfligning.Text & ", x(E), y(E), Xmax, Tic);"
        Fundet = True
    End If
    
    If Not Fundet Then
        s = s & "A=(1, 2);"
        s = s & "SolveODE(" & TextBox_dfligning.Text & ", x(A), y(A), Xmin, Tic);" ' y(A) virker ikke
        s = s & "SolveODE(" & TextBox_dfligning.Text & ", x(A), y(A), Xmax, Tic);" ' y(A) virker ikke
    End If
    OpenGeoGebraWeb s, "Classic", True, True
    
End Sub

Private Sub Label_helpmarker_Click()
MsgBox TT.A(195), vbOKOnly, TT.A(808)
End Sub

Private Sub Label_punkter2_Click()
MsgBox TT.A(196), vbOKOnly, TT.A(808)
End Sub

Private Sub Label_symbol_Click()
Dim Ctrl As control
On Error GoTo fejl
Set Ctrl = Me.ActiveControl
If Left(Ctrl.Name, 7) <> "TextBox" Then Set Ctrl = TextBox_titel
UserFormSymbol.Show
Ctrl.Text = Ctrl.Text & UserFormSymbol.tegn
fejl:
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
Dim D As String, nd As String
Dim xmin As String, xmax As String
Dim Arr As Variant, i As Integer
On Error Resume Next
    SetCaptions
#If Mac Then
    If Opened Then Exit Sub
    Opened = True
    Me.Left = 10
    Me.Top = 80
    Label_wait.Left = 180
    Label_wait.Top = 270
    Label_zoom.visible = False
    Kill GetTempDir() & "WordMatGraf.pdf"
#Else
    Kill GetTempDir() & "\WordMatGraf.gif"
#End If

If Not PicOpen Then
    omax.PrepareNewCommand '
    If Len(omax.DefString) > 1 Then
    D = omax.defstringtext

    D = Replace(D, "assume", "")
    D = Replace(D, ":=", "=")
    D = Replace(D, ":", "=")
'    d = omax.ConvertToAscii(omax.ConvertToWordSymbols(d)) ' fjernet efter defstringtext anvendes
    D = Trim(D)
    
    ' reverse definition order
    Arr = Split(D, "$")
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
    ' insert xmin and xmax if defined
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

    SetEscEvents Me.Controls
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

    etikettext = InputBox(TT.A(299), TT.A(298), "")
    
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

Private Sub Label_opdater_Click()
OpdaterGraf
#If Mac Then
    ShowPreviewMac
#Else
    Me.Repaint
#End If
End Sub

Private Sub Label_punkterhelp_Click()
    MsgBox TT.A(197), vbOKOnly, TT.A(808)
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

    Dim Path As String
    Dim ils As InlineShape
    Dim xmin As Double, xmax As Double
    Dim plinjer As Variant
    Dim linje As Variant
    Dim i As Integer

    If cxl Is Nothing Then Set cxl = New CExcel
    Application.ScreenUpdating = False
    Me.hide
    Dim UFwait2 As New UserFormWaitForMaxima
    UFwait2.Show vbModeless
    DoEvents
    UFwait2.Label_progress = "***"


    If Not embed Then
        cxl.LoadFile ("Graphs.xltm")
        Set WB = cxl.xlwb
        Set ws = cxl.xlwb.Sheets("Tabel")
        UFwait2.Label_progress = UFwait2.Label_progress & "***"


    Else

        Path = """" & GetProgramFilesDir & "\WordMat\ExcelFiles\Graphs.xltm"""
        PrepareMaxima
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph
        UFwait2.Label_progress = UFwait2.Label_progress & "**"

        Set ils = ActiveDocument.InlineShapes.AddOLEObject(fileName:=Path, LinkToFile:=False, DisplayAsIcon:=False, Range:=Selection.Range)

        'Ils.Height = 300
        'Ils.Width = 500
        UFwait2.Label_progress = UFwait2.Label_progress & "***********"


        'Ils.OLEFormat.DoVerb (wdOLEVerbOpen)
        ils.OLEFormat.DoVerb (wdOLEVerbShow)
        Set WB = ils.OLEFormat.Object
        Set ws = WB.Sheets("Tabel")
        'ws.Activate
    End If

    XLapp.Application.EnableEvents = False
    XLapp.Application.ScreenUpdating = False

    UFwait2.Label_progress = UFwait2.Label_progress & "*****"
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
    'xmin og xmax copied over
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
    'variable name copied over
    ws.Range("B1").Value = Me.TextBox_var1.Text
    ws.Range("C1").Value = Me.TextBox_var2.Text
    ws.Range("D1").Value = Me.TextBox_var3.Text
    ws.Range("E1").Value = Me.TextBox_var4.Text
    ws.Range("F1").Value = Me.TextBox_var5.Text
    ws.Range("G1").Value = Me.TextBox_var6.Text
    ' iSettings
    If Radians Then
        ws.Range("A4").Value = "rad"
    Else
        ws.Range("A4").Value = "grad"
    End If

    On Error GoTo slut

    'datapoints
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


slut:
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
    UFwait2.Label_progress = UFwait2.Label_progress & "**"

    'excel.Run ("UpDateAll")
    XLapp.Run ("UpDateAll")
    
    UFwait2.Label_progress = UFwait2.Label_progress & "***"
    WB.Charts(1).Activate
    XLapp.Application.EnableEvents = True
    XLapp.Application.ScreenUpdating = True
slut2:
    Unload UFwait2

End Sub
Sub SetLineStyle(CB As ComboBox, n As Integer)
' sets linestyle according to combobox

If CB.ListIndex = 0 Then
XLapp.ActiveChart.SeriesCollection(n).Border.LineStyle = xlContinuous '
ElseIf CB.ListIndex = 1 Then
XLapp.ActiveChart.SeriesCollection(n).Border.LineStyle = xlDot 'xlContinuous '
ElseIf CB.ListIndex = 2 Then
XLapp.ActiveChart.SeriesCollection(n).Border.LineStyle = xlDash 'xlContinuous '
ElseIf CB.ListIndex = 3 Then
XLapp.ActiveChart.SeriesCollection(n).Border.LineStyle = xlDashDot 'xlContinuous '
ElseIf CB.ListIndex = 4 Then
XLapp.ActiveChart.SeriesCollection(n).Border.LineStyle = xlDashDotDot 'xlContinuous '
Else
XLapp.ActiveChart.SeriesCollection(n).Border.LineStyle = xlContinuous '
End If

End Sub
Sub InsertFormula(ws As Variant, WB As Variant, tb As TextBox, col As Integer)
' inserts formula from textbox in column col
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
    ' find available variable plot
    Do While ws.Range("N6").Offset(i, 0).Value <> ""
      i = i + 1
    Loop
    On Error Resume Next
    Do While varnavn <> ""

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
     
     ' insert function in spreadsheet
     ws.Range("b7").Offset(0, col).Formula = "=" & forskrift

    ' copy formula down
    ws.Activate
    ws.Range("B7").Offset(0, col).Select
    ws.Application.Selection.AutoFill Destination:=ws.Range("b7:b207").Offset(0, col), Type:=0   'xlFillDefault=0
    
    ' Error in any cells?
    For i = 0 To 200
        If TypeName(ws.Range("b7").Offset(i, col).Value) = "Error" Then ws.Range("b7").Offset(i, col).Value = ""
    Next
        
End If
GoTo slut:
fejlindtast:
    MsgBox TT.A(300) & " " & col + 1
slut:
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
    forskrift = Replace(forskrift, VBA.ChrW(12310), "") ' speciel invisible parenthesis removed
    forskrift = Replace(forskrift, VBA.ChrW(12311), "") ' speciel invisible parenthesis removed
'    forskrift = Replace(forskrift, VBA.ChrW(11), "")
    forskrift = Replace(forskrift, vbLf, "") ' shift-enter and enter
    forskrift = Replace(forskrift, vbCrLf, "")
    forskrift = Replace(forskrift, vbCr, "")
    forskrift = Replace(forskrift, """", "") ' apostrof removed
    forskrift = Replace(forskrift, VBA.ChrW(8289), "") ' symbol that defines function removed
    forskrift = Replace(forskrift, VBA.ChrW(8212), "+") 'double minus sign equals plus
    forskrift = Replace(forskrift, VBA.ChrW(183), "*") ' dot replaced by multiplication
    forskrift = Replace(forskrift, VBA.ChrW(215), "*") ' cross replaced by multiplication
    forskrift = Replace(forskrift, VBA.ChrW(8901), "*") ' \cdot replaced by multiplication
    forskrift = Replace(forskrift, VBA.ChrW(8226), "*") ' thick dot replaced by multiplication
    forskrift = Replace(forskrift, "%", "/100") ' percentage sign
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
        forskrift = Left(forskrift, Pos - 1) & "abs(" & Mid(forskrift, Pos + 1, posb - Pos - 1) & ")" & Right(forskrift, Len(forskrift) - posb)
    End If
    Loop While Pos > 0
    
    ' 3 og 4 root
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
    
    'squareroot
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
                forskrift = Left(ea.Text, Pos) & Right(ea.Text, Len(ea.Text) - posog)
               
            End If
        End If
    End If
    Loop While Pos > 0
    
    'trigfunctions if 360 degrees
    If Not (Radians) Then
        forskrift = ConvertDegreeToRad(forskrift, "sin")
        forskrift = ConvertDegreeToRad(forskrift, "cos")
        forskrift = ConvertDegreeToRad(forskrift, "tan")
        forskrift = ConvertDegreeToRad(forskrift, "sec")
        forskrift = ConvertDegreeToRad(forskrift, "cot")
        forskrift = ConvertDegreeToRad(forskrift, "csc")
    End If
    
    ' find understood parenthesis after ^ and / ' (this line must be after diff and other functions with comma)
    ea.Text = forskrift
    ea.InsertBracketAfter ("^")
    ea.InsertBracketAfter ("/")
    forskrift = ea.Text
    
    ' space removed
    forskrift = Replace(forskrift, " ", "")

    ' insert understood multiplication
    ea.Text = forskrift
    ea.Pos = 1
    ea.InsertMultSigns
    forskrift = ea.Text
    
    ConvertToExcelFormula = forskrift

End Function
Function GetDraw2Dtext(Optional highres As Double = 1) As String
On Error GoTo fejl
Dim grafobj As String
Dim xmin As String
Dim xmax As String
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
Dim Arr2 As Variant
Dim i As Integer
Dim vekt As String
Dim parx As String
Dim pary As String
Dim tmin As String
Dim tmax As String
Dim X As String
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

'parametric pplots
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


'points
If TextBox_punkter.Text <> "" Then
    grafobj = grafobj & "key="""",color=black,"
    Arr = Split(TextBox_punkter.Text, VbCrLfMac)
    For i = 0 To UBound(Arr)
    If InStr(Arr(i), ";") > 0 Or InStr(Arr(i), vbTab) > 0 Then
        Arr(i) = Replace(Arr(i), ",", ".")
        Arr(i) = Replace(Arr(i), ";", ",")
    End If
        Arr(i) = Replace(Arr(i), vbTab, ",") ' if tab is copied from Excel
        Arr(i) = Replace(Arr(i), " ", "")
        If Len(Arr(i)) > 0 Then
        If Left(Arr(i), 1) <> "(" Then
            Arr(i) = "(" & Arr(i)
        End If
        If Right(Arr(i), 1) <> ")" Then
            Arr(i) = Arr(i) & ")"
        End If
        Arr(i) = Replace(Arr(i), "),(", "],[")
        Arr(i) = Replace(Arr(i), ");(", "],[")
        Arr(i) = Replace(Arr(i), "(", "[")
        Arr(i) = Replace(Arr(i), ")", "]")
        punkttekst = punkttekst & Arr(i) & ","
        End If
    Next
    If Right(punkttekst, 1) = "," Then punkttekst = Left(punkttekst, Len(punkttekst) - 1)
    
    grafobj = grafobj & "point_type=filled_circle,point_size=" & Replace(highres * ConvertStringToNumber(TextBox_pointsize.Text), ",", ".") & ",points_joined=" & VBA.LCase(CheckBox_pointsjoined.Value) & ",points([" & punkttekst & "]),"
End If

'points 2
If TextBox_punkter2.Text <> "" Then
    punkttekst = ""
    grafobj = grafobj & "key="""",color=blue,"
    Arr = Split(TextBox_punkter2.Text, VbCrLfMac)
    For i = 0 To UBound(Arr)
    If InStr(Arr(i), ";") > 0 Or InStr(Arr(i), vbTab) > 0 Then
        Arr(i) = Replace(Arr(i), ",", ".")
        Arr(i) = Replace(Arr(i), ";", ",")
    End If
        Arr(i) = Replace(Arr(i), vbTab, ",") ' if tab is copied from Excel
        Arr(i) = Replace(Arr(i), " ", "")
        If Len(Arr(i)) > 0 Then
        If Left(Arr(i), 1) <> "(" Then
            Arr(i) = "(" & Arr(i)
        End If
        If Right(Arr(i), 1) <> ")" Then
            Arr(i) = Arr(i) & ")"
        End If
        Arr(i) = Replace(Arr(i), "),(", "],[")
        Arr(i) = Replace(Arr(i), ");(", "],[")
        Arr(i) = Replace(Arr(i), "(", "[")
        Arr(i) = Replace(Arr(i), ")", "]")
        punkttekst = punkttekst & Arr(i) & ","
        End If
    Next
    If Right(punkttekst, 1) = "," Then punkttekst = Left(punkttekst, Len(punkttekst) - 1)
    
    grafobj = grafobj & "point_type=filled_circle,point_size=" & Replace(TextBox_pointsize2.Text, ",", ".") & ",points_joined=" & VBA.LCase(CheckBox_pointsjoined2.Value) & ",points([" & punkttekst & "]),"
End If

'selected points
If TextBox_markerpunkter.Text <> "" Then
    punkttekst = ""
    grafobj = grafobj & "key="""",color=red,"
    Arr = Split(TextBox_markerpunkter.Text, VbCrLfMac)
    For i = 0 To UBound(Arr)
    If InStr(Arr(i), ";") > 0 Or InStr(Arr(i), vbTab) > 0 Then
        Arr(i) = Replace(Arr(i), ",", ".")
        Arr(i) = Replace(Arr(i), ";", ",")
    End If
        Arr(i) = Replace(Arr(i), vbTab, ",") 'if tab is copied from Excel
        Arr(i) = Replace(Arr(i), " ", "")
        If Len(Arr(i)) > 0 Then
        Arr2 = Split(Arr(i), ",")
        If UBound(Arr2) = 1 Then
            X = Arr2(0)
            Y = Arr2(1)
            punkttekst = punkttekst & "points([[" & X & ",0],[" & X & "," & Y & "],[0," & Y & "]]),"
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
                Arr2 = Split(Arr(i), ";")
            Else
                Arr2 = Split(Arr(i), ",")
            End If
            If UBound(Arr2) >= 2 Then
            labeltext = labeltext & "["
            labeltext = labeltext & """" & Arr2(0) & """"
            labeltext = labeltext & "," & Replace(Arr2(1), ",", ".")
            labeltext = labeltext & "," & Replace(Arr2(2), ",", ".")
            labeltext = labeltext & "],"
            End If
        Next
        If Len(labeltext) > 0 Then
            labeltext = Left(labeltext, Len(labeltext) - 1)
            grafobj = "color=black,label(" & ConvertDrawLabel(labeltext) & ")," & grafobj
        End If
    End If
    
    If Len(grafobj) = 0 Then GoTo slut

    
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
    grafobj = "xtics_axis = true," & grafobj
    grafobj = "ytics_axis = true," & grafobj
    grafobj = "line_width=" & Replace(highres, ",", ".") & "," & grafobj
    If Not MaximaComplex Then grafobj = "draw_realpart = false," & grafobj
    
    If CheckBox_logx.Value Then
        If ConvertStringToNumber(TextBox_xmin.Text) > 0 Then
            grafobj = "logx=true," & grafobj
        Else
            MsgBox "xmin must be >0 to use log x-axis."
        End If
    End If
    If CheckBox_logy.Value Then
        If ConvertStringToNumber(TextBox_ymin.Text) > 0 Then
            grafobj = "logy=true," & grafobj
        Else
            MsgBox "ymin must be >0 to use log y-axis."
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
GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
End Function
Private Sub OpdaterGraf(Optional highres As Double = 1)
Dim Text As String
Dim df As String
Dim dfsol As String
On Error GoTo fejl
    Label_wait.Caption = TT.A(826) & "!"
    Label_wait.Font.Size = 36
    Label_wait.visible = True
    omax.PrepareNewCommand FindDef:=False
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
            df = df & ",duration=100,solns_at(" & dfsol & ")" ' duration defaulat is 10. by increasing you can plot closer to asympototes
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
Private Sub GnuPlot()
Dim Text As String
    omax.PrepareNewCommand FindDef:=False
    InsertDefinitioner

    Text = GetDraw2Dtext()
    
    If Len(Text) > 0 Then
    Call omax.Draw2D(Text, "", omax.ConvertToAscii(TextBox_xaksetitel.Text), omax.ConvertToAscii(TextBox_yaksetitel.Text), CheckBox_gitter.Value, False, 1)
    DoEvents
    End If

GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
End Sub

Private Sub CommandButton_linregr_Click()
    Dim Cregr As New CRegression
    On Error GoTo slut
    Cregr.Datatext = TextBox_punkter.Text
    Cregr.ComputeLinRegr
'    Selection.Collapse
'    Selection.TypeParagraph
'    Cregr.InsertEquation

    If TextBox_ligning1.Text = "" Then
        TextBox_ligning1.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning2.Text = "" Then
        TextBox_ligning2.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning3.Text = "" Then
        TextBox_ligning3.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning4.Text = "" Then
        TextBox_ligning4.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning5.Text = "" Then
        TextBox_ligning5.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning6.Text = "" Then
        TextBox_ligning6.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    End If

    OpdaterGraf
    Me.Repaint
slut:
End Sub

Private Sub CommandButton_polregr_Click()
    Dim Cregr As New CRegression
    On Error GoTo slut
    
    Cregr.Datatext = TextBox_punkter.Text
    Cregr.ComputePolRegr
'    Selection.Collapse
'    Selection.TypeParagraph
'    Cregr.InsertEquation

    If TextBox_ligning1.Text = "" Then
        TextBox_ligning1.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning2.Text = "" Then
        TextBox_ligning2.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning3.Text = "" Then
        TextBox_ligning3.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning4.Text = "" Then
        TextBox_ligning4.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning5.Text = "" Then
        TextBox_ligning5.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning6.Text = "" Then
        TextBox_ligning6.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    End If
    
    OpdaterGraf
    Me.Repaint
slut:
End Sub
Private Sub CommandButton_ekspregr_Click()
   Dim Cregr As New CRegression
    On Error GoTo slut
    
    Cregr.Datatext = TextBox_punkter.Text
    Cregr.ComputeExpRegr
'    Selection.Collapse
'    Selection.TypeParagraph
'    Cregr.InsertEquation

    If TextBox_ligning1.Text = "" Then
        TextBox_ligning1.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning2.Text = "" Then
        TextBox_ligning2.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning3.Text = "" Then
        TextBox_ligning3.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning4.Text = "" Then
        TextBox_ligning4.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning5.Text = "" Then
        TextBox_ligning5.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning6.Text = "" Then
        TextBox_ligning6.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    End If
    
    OpdaterGraf
    Me.Repaint
slut:
End Sub

Private Sub CommandButton_potregr_Click()
       Dim Cregr As New CRegression
    On Error GoTo slut
    
    Cregr.Datatext = TextBox_punkter.Text
    Cregr.ComputePowRegr

    If TextBox_ligning1.Text = "" Then
        TextBox_ligning1.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning2.Text = "" Then
        TextBox_ligning2.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning3.Text = "" Then
        TextBox_ligning3.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning4.Text = "" Then
        TextBox_ligning4.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning5.Text = "" Then
        TextBox_ligning5.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    ElseIf TextBox_ligning6.Text = "" Then
        TextBox_ligning6.Text = Right(Cregr.Ligning, Len(Cregr.Ligning) - 2)
    End If

    OpdaterGraf
    Me.Repaint
slut:
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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
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

Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    gemx = X
    gemy = Y
    If Len(etikettext) > 0 Then
        If Len(TextBox_labels.Text) > 0 Then
        TextBox_labels.Text = TextBox_labels.Text & VbCrLfMac
        End If
        If Len(TextBox_ymin.Text) = 0 Or Len(TextBox_ymax.Text) = 0 Then
            MsgBox TT.A(301), vbOKOnly, TT.Error
        Else
        TextBox_labels.Text = TextBox_labels.Text & etikettext & ";" & ConvertPixelToCoordX(X) & ";" & ConvertPixelToCoordY(Y)
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
            MsgBox TT.A(301), vbOKOnly, TT.Error
        Else
        TextBox_punkter2.Text = TextBox_punkter2.Text & ConvertPixelToCoordX(X) & ";" & ConvertPixelToCoordY(Y)
        OpdaterGraf
        Me.Repaint
        End If
    ElseIf nytmarkerpunkt Then
        If Len(TextBox_markerpunkter.Text) > 0 Then
        TextBox_markerpunkter.Text = TextBox_markerpunkter.Text & VbCrLfMac
        End If
        If Len(TextBox_ymin.Text) = 0 Or Len(TextBox_ymax.Text) = 0 Then
            MsgBox TT.A(301), vbOKOnly, TT.Error
        Else
        TextBox_markerpunkter.Text = TextBox_markerpunkter.Text & ConvertPixelToCoordX(X) & ";" & ConvertPixelToCoordY(Y)
        nytmarkerpunkt = False
        OpdaterGraf
        Me.Repaint
        End If
    Else
    Label_zoom.Left = gemx + Image1.Left
    Label_zoom.Top = gemy + Image1.Top
    Label_zoom.Width = 1
    Label_zoom.Height = 1
    Label_zoom.visible = True
    End If
End Sub

Private Sub Image1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
'image1.Picture.Render(image1.Picture.Handle,0,0,600,600,image1.Picture.Width,image1.Picture.Height
'hDC, 0, 0, ScaleWidth, ScaleHeight, 0, p.Height, p.Width, -p.Height, ByVal 0
    If Label_zoom.visible Then
    Label_zoom.Left = gemx + Image1.Left
    Label_zoom.Top = gemy + Image1.Top
    Label_zoom.Width = X - gemx
    Label_zoom.Height = Y - gemy
    End If
End Sub

Private Sub Image1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim xmin As Single
Dim xmax As Single
Dim Ymin As Single
Dim Ymax As Single
Dim s As String

Label_zoom.visible = False
If Abs(X - gemx) < 5 Then GoTo slut

xmin = ConvertStringToNumber(TextBox_xmin.Text)
xmax = ConvertStringToNumber(TextBox_xmax.Text)
Ymin = ConvertStringToNumber(TextBox_ymin.Text)
Ymax = ConvertStringToNumber(TextBox_ymax.Text)

s = ConvertPixelToCoordX(gemx)
TextBox_xmax.Text = ConvertPixelToCoordX(X)
TextBox_xmin.Text = s
If TextBox_ymin.Text <> "" And TextBox_ymax.Text <> "" Then
    s = ConvertPixelToCoordY(Y)
    TextBox_ymax.Text = ConvertPixelToCoordY(gemy)
    TextBox_ymin.Text = s
End If
OpdaterGraf

    Me.Repaint
slut:
End Sub
Private Sub Image1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim xmin As Single
Dim xmax As Single
Dim Ymin As Single
Dim Ymax As Single
Dim dx As Single, dy As Single
Dim nyy As Single
Dim s As String
Label_zoom.visible = False

xmin = ConvertStringToNumber(TextBox_xmin.Text)
xmax = ConvertStringToNumber(TextBox_xmax.Text)
Ymin = ConvertStringToNumber(TextBox_ymin.Text)
Ymax = ConvertStringToNumber(TextBox_ymax.Text)
dx = (xmax - xmin) * 0.3
dy = (Ymax - Ymin) * 0.3
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
GoTo slut
fejl:
    MsgBox TT.A(95), vbOKOnly, TT.Error
slut:
    
End Sub
Function ConvertPixelToCoordX(X As Single) As Single
Dim xmin As Single, xmax As Single, cfakt As Single
xmin = ConvertStringToNumber(TextBox_xmin.Text)
xmax = ConvertStringToNumber(TextBox_xmax.Text)
cfakt = (xmax - xmin) / (Image1.Width * 0.9)
X = X - Image1.Width * 0.06
ConvertPixelToCoordX = xmin + cfakt * X
End Function
Function ConvertPixelToCoordY(Y As Single) As Single
Dim Ymin As Single, Ymax As Single, cfakt As Single
Ymin = ConvertStringToNumber(TextBox_ymin.Text)
Ymax = ConvertStringToNumber(TextBox_ymax.Text)
cfakt = (Ymax - Ymin) / (Image1.Height * 0.9)
Y = Image1.Height - Y
Y = Y - Image1.Height * 0.08
ConvertPixelToCoordY = Ymin + cfakt * Y
End Function
Private Sub CommandButton_zoom_Click()
Dim dx As Single, dy As Single
Dim midtx As Single, midty As Single
Dim xmin As Single
Dim xmax As Single
Dim Ymin As Single
Dim Ymax As Single
On Error GoTo fejl
xmin = ConvertStringToNumber(TextBox_xmin.Text)
xmax = ConvertStringToNumber(TextBox_xmax.Text)
Ymin = ConvertStringToNumber(TextBox_ymin.Text)
Ymax = ConvertStringToNumber(TextBox_ymax.Text)

midtx = (xmax + xmin) / 2
midty = (Ymax + Ymin) / 2
dx = (xmax - xmin) * 0.3
dy = (Ymax - Ymin) * 0.3

TextBox_xmin.Text = betcif(midtx - dx, 5, False)
TextBox_xmax.Text = betcif(midtx + dx, 5, False)
If TextBox_ymin.Text <> "" And TextBox_ymax.Text <> "" Then
TextBox_ymin.Text = betcif(midty - dy, 5, False)
TextBox_ymax.Text = betcif(midty + dy, 5, False)
End If
OpdaterGraf

Me.Repaint
GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:

End Sub

Private Sub CommandButton_zoomud_Click()
Dim dx As Single, dy As Single
Dim midtx As Single, midty As Single
Dim xmin As Single
Dim xmax As Single
Dim Ymin As Single
Dim Ymax As Single
On Error GoTo fejl
xmin = ConvertStringToNumber(TextBox_xmin.Text)
xmax = ConvertStringToNumber(TextBox_xmax.Text)
Ymin = ConvertStringToNumber(TextBox_ymin.Text)
Ymax = ConvertStringToNumber(TextBox_ymax.Text)

midtx = (xmax + xmin) / 2
midty = (Ymax + Ymin) / 2
dx = (xmax - xmin) * 1
dy = (Ymax - Ymin) * 1

TextBox_xmin.Text = betcif(midtx - dx, 5, False)
TextBox_xmax.Text = betcif(midtx + dx, 5, False)
If TextBox_ymin.Text <> "" And TextBox_ymax.Text <> "" Then
TextBox_ymin.Text = betcif(midty - dy, 5, False)
TextBox_ymax.Text = betcif(midty + dy, 5, False)
End If
OpdaterGraf
 
Me.Repaint
GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:

End Sub
Private Sub Label_insertpic_Click()
Dim ils As InlineShape
Dim s As String
Dim Sep As String
On Error GoTo fejl
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
s = AppNavn & Sep & AppVersion & Sep & TextBox_definitioner.Text & Sep & TextBox_titel.Text & Sep & TextBox_xaksetitel.Text & Sep & TextBox_yaksetitel.Text & Sep
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
GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
Application.ScreenUpdating = True
End Sub

Sub InsertDefinitioner()
    ' inserts definitions from textbox in maximainputstring
    Dim DefString As String

    omax.InsertKillDef

    DefString = GetDefString

    If Len(DefString) > 0 Then
        If Right(DefString, 1) = "," Then DefString = Left(DefString, Len(DefString) - 1)
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
' checks for variables in textboxes and inderts under definitions
Dim Vars As String
Dim Var As String, var2 As String
Dim ea As New ExpressionAnalyser
Dim Arr As Variant
Dim i As Integer
    
    
    Vars = Vars & GetTextboxVars(TextBox_ligning1, TextBox_var1)
    Vars = Vars & GetTextboxVars(TextBox_ligning2, TextBox_var2)
    Vars = Vars & GetTextboxVars(TextBox_ligning3, TextBox_var3)
    Vars = Vars & GetTextboxVars(TextBox_ligning4, TextBox_var4)
    Vars = Vars & GetTextboxVars(TextBox_ligning5, TextBox_var5)
    Vars = Vars & GetTextboxVars(TextBox_ligning6, TextBox_var6)
    Vars = Vars & GetTextboxLignVars(TextBox_lig1)
    Vars = Vars & GetTextboxLignVars(TextBox_lig2)
    Vars = Vars & GetTextboxLignVars(TextBox_Lig3)
    
    
    omax.FindVariable Vars, False ' removes doubles
    Vars = omax.Vars
    If Left(Vars, 1) = ";" Then Vars = Right(Vars, Len(Vars) - 1)
    
    ea.Text = Vars
    Do While Right(TextBox_definitioner.Text, 2) = VbCrLfMac
        TextBox_definitioner.Text = Left(TextBox_definitioner.Text, Len(TextBox_definitioner.Text) - 2)
    Loop
    Arr = Split(TextBox_definitioner.Text, VbCrLfMac)
    
    Do
    Var = ea.GetNextListItem
    Var = Replace(Var, VbCrLfMac, "")
    For i = 0 To UBound(Arr)
        If Arr(i) <> "" Then
        var2 = Split(Arr(i), "=")(0)
        If var2 = Var Then
            Var = ""
            Exit For
        End If
        End If
    Next
    If Var <> "" And Var <> "if" And Var <> "then" And Var <> "else" And Var <> "elseif" And Var <> "and" And Var <> "or" Then
'        If Right(TextBox_definitioner.text, 2) <> vbCrLf Then
        If Len(TextBox_definitioner.Text) > 0 Then
            TextBox_definitioner.Text = TextBox_definitioner.Text & VbCrLfMac
        End If
        TextBox_definitioner.Text = TextBox_definitioner.Text & Var & "=1"
    End If
    Loop While ea.Pos <= Len(ea.Text)

    
End Sub
Function GetTextboxVars(tb As TextBox, tbvar As TextBox) As String
Dim ea As New ExpressionAnalyser
Dim Var As String
    If Len(tb.Text) > 0 Then
        omax.Vars = ""
        omax.FindVariable (tb.Text)
        If InStr(omax.Vars, "x") > 0 Then
            Var = "x"
        ElseIf InStr(omax.Vars, "t") > 0 Then
            Var = "t"
        Else
            ea.Text = omax.Vars
            Var = ea.GetNextVar(1)
        End If
        If Len(Var) > 0 Then
            tbvar.Text = Var
        End If
        omax.Vars = RemoveVar(omax.Vars, tbvar.Text)
        If Len(omax.Vars) > 0 Then
            GetTextboxVars = ";" & omax.Vars
        End If
    End If
End Function
Function GetTextboxLignVars(tb As TextBox) As String
    If Len(tb.Text) > 0 Then
        omax.Vars = ""
        omax.FindVariable (tb.Text)
        omax.Vars = RemoveVar(omax.Vars, "x")
        omax.Vars = RemoveVar(omax.Vars, "y")
        If Len(omax.Vars) > 0 Then
            GetTextboxLignVars = ";" & omax.Vars
        End If
    End If
End Function
Function RemoveVar(Text As String, Var As String)
' removes var from string
Dim ea As New ExpressionAnalyser

ea.Text = Text
Call ea.ReplaceVar(Var, "")
Text = Replace(ea.Text, ";;", ";")
If Left(Text, 1) = ";" Then Text = Right(Text, Len(Text) - 1)
If Right(Text, 1) = ";" Then Text = Left(Text, Len(Text) - 1)

RemoveVar = Text
End Function

Sub opdaterLabels()
    Label_diffy.Caption = TextBox_dfy.Text & "'(" & TextBox_dfx.Text & ")="
End Sub
Private Sub CommandButton_plotdf_Click()
Dim pm As String
Dim sl As String
    Label_vent.visible = True
    omax.PrepareNewCommand FindDef:=False
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
' checks if there are any defs in the the def-textbox and uses them to limit xmin and xmax
Dim DefS As String
Dim Pos As Integer
Dim ea As New ExpressionAnalyser
Dim ea2 As New ExpressionAnalyser
Dim s As String, l As String
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

Sub InsertBoundary(Var As String, assumetext As String, tbmin As TextBox, tbmax As TextBox)
Dim dlhs As String, drhs As String
Dim Arr As Variant
    Arr = Split(assumetext, "<")
    If UBound(Arr) > 0 Then
        dlhs = Replace(Arr(0), "=", "")
        drhs = Replace(Arr(1), "=", "")
        If dlhs = Var Then
            tbmax.Text = drhs
        ElseIf drhs = Var Then
            tbmin.Text = dlhs
        End If
    End If
    Arr = Split(assumetext, ">")
    If UBound(Arr) > 0 Then
        dlhs = Replace(Arr(0), "=", "")
        drhs = Replace(Arr(1), "=", "")
        If dlhs = Var Then
            tbmin.Text = drhs
        ElseIf drhs = Var Then
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
    Me.Caption = TT.A(799)
    Label_wait.Caption = TT.A(826)
    Label_opdater.Caption = TT.A(813)
    MultiPage1.Pages("Page1").Caption = TT.A(832)
    MultiPage1.Pages("Page2").Caption = TT.A(834)
    MultiPage1.Pages("Page3").Caption = TT.A(835)
    MultiPage1.Pages("Page5").Caption = TT.A(836)
    MultiPage1.Pages("Page7").Caption = TT.A(462)
    MultiPage1.Pages("Page4").Caption = TT.A(444)
    Label29.Caption = TT.A(823)
    Label45.Caption = TT.A(837)
    Label_ligninger.Caption = TT.A(832) & "  f(x)=..."
    CommandButton_nulstil1.Caption = TT.Reset
    CommandButton_nulstil2.Caption = TT.Reset
    CommandButton_nulstil3.Caption = TT.Reset
    CommandButton_nulstil4.Caption = TT.Reset
    CommandButton_nulstil5.Caption = TT.Reset
    CommandButton_nulstil6.Caption = TT.Reset
    CommandButton_nulstil.Caption = TT.A(800)
    Label_insertpic.Caption = TT.A(93) 'TT.OK
    Label_cancel.Caption = TT.Cancel
    Label2.Caption = "x-" & TT.A(801)
    Label3.Caption = "y-" & TT.A(801)
    Label4.Caption = TT.A(802)
    CommandButton_nulstillign1.Caption = TT.Reset
    CommandButton_nulstilligning2.Caption = TT.Reset
    CommandButton_nulstillign3.Caption = TT.Reset
    Label50.Caption = TT.A(833) & " 1"
    Label51.Caption = TT.A(833) & " 2"
    Label52.Caption = TT.A(833) & " 3"
    
    Label1.Caption = TT.A(198)
    CommandButton_kugle.Caption = TT.A(199) 'insert circle
    CommandButton_insertplan.Caption = TT.A(200) ' insert line
    CommandButton_parlinje.Caption = TT.A(200)
    CommandButton_nulstilpar1.Caption = TT.Reset
    Label10.Caption = TT.A(835)
    Label48.Caption = TT.A(835) & " 2"
    Label53.Caption = TT.A(201) ' selected points
    Label28.Caption = TT.A(202)
    
End Sub

Private Sub Label_cancel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorPress
End Sub
Private Sub Label_cancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorHover
End Sub
Private Sub Label_insertpic_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_insertpic.BackColor = LBColorPress
End Sub
Private Sub Label_insertpic_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_insertpic.BackColor = LBColorHover
End Sub
Private Sub Label_opdater_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_opdater.BackColor = LBColorPress
End Sub
Private Sub Label_opdater_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_opdater.BackColor = LBColorHover
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_insertpic.BackColor = LBColorInactive
    Label_cancel.BackColor = LBColorInactive
    Label_opdater.BackColor = LBColorInactive
End Sub
