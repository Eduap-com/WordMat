Attribute VB_Name = "Excelsubs"
Option Explicit
Function InsertIndlejretExcel() As Object
    '    InsertIndlejret "TomExcel.xlsx"
    On Error GoTo slut

    Dim vers As String
    Dim oWS As Object ' Worksheet Object
    Dim ils As InlineShape
    Dim oWD As Document ' Word Document Object (Use Microsoft Word Reference)
    Set oWD = ActiveDocument
    Set ils = Selection.InlineShapes.AddOLEObject(ClassType:="Excel.Sheet", LinkToFile:=False, DisplayAsIcon:=False)

    Set oWS = ils.OLEFormat.Object
    Set InsertIndlejretExcel = oWS
slut:
End Function

Sub ExcelIns()
    Application.Run macroname:="TableExcelSpreadsheetInsert"
End Sub
Sub Chi2Test()
#If Mac Then
    InsertOpenExcel "chi2test.xltm"
#Else
    Dim s As String
    Dim Arr As Variant
    Dim r As Integer
    Dim c As Integer
    
On Error GoTo Fejl

Dim signiv As Integer
signiv = InputBox(Sprog.A(349), Sprog.A(350), "5")
If Selection.Tables.Count = 0 Then
    s = InputBox(Sprog.A(351), Sprog.A(352), "2x2")
    Arr = Split(s, "x")
    If UBound(Arr) < 1 Then Arr = Split(s, ",")
    r = Arr(0)
    c = Arr(1)
End If

'Application.ScreenUpdating = False

If signiv <= 0 Then Exit Sub
Dim cxl As New CExcel
'cxl.vis = False
cxl.vis = True
cxl.NewFile
If r = 0 Or c = 0 Then
    cxl.GetTableData
End If
    cxl.Chi2Test signiv / 100, r, c

If r > 0 And c > 0 Then GoTo slut

Selection.Collapse
Selection.InsertAfter Sprog.A(353) & vbCrLf
Selection.InsertAfter Sprog.A(354) & ": " & vbTab & cxl.p & " = " & cxl.p * 100 & "%" & vbCrLf

If cxl.p * 100 < signiv Then
    Selection.InsertAfter Sprog.A(355) & " " & signiv & Sprog.A(356) & vbCrLf
Else
    Selection.InsertAfter Sprog.A(355) & " " & signiv & Sprog.A(357) & vbCrLf
End If
If cxl.Below5 Or cxl.sum < 50 Then
    Selection.InsertAfter Sprog.A(358) & vbCrLf
End If

GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
Application.ScreenUpdating = True
#End If
End Sub
Sub Chi2GrafNoLoad()
    'On Error GoTo fejl
    Dim XLapp As Object
    Dim xlwb As Object
    Dim xlsh As Object
    Dim ch As Variant ' chart?
    Dim i As Integer

    ' If independent sheet
    On Error Resume Next
    Set XLapp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set XLapp = CreateObject("Excel.Application")
    End If
    Set xlwb = XLapp.Workbooks.Add

    Set xlsh = xlwb.Sheets(1)
    xlsh.Activate
    xlsh.Columns("A:E").ColumnWidth = 0.5
    xlsh.Cells(1, 1).Formula = "0"
    xlsh.Cells(2, 1).Formula = "=G$3/100+A1"
    xlsh.Cells(1, 2).Formula = "=CHIDIST(A1,G$2)"

    xlsh.Range("F1").Value = Sprog.A(359) '"Chifordeling"
    xlsh.Range("F2").Value = Sprog.A(360) & ":"
    xlsh.Range("G2").Value = 7
    xlsh.Range("F3").Value = "Xmax:"
    xlsh.Range("G3").Value = 15

    xlsh.Range("F5").Value = Sprog.A(361) & ":"
    xlsh.Range("G5").Value = 0.05

    xlsh.Range("H5").Formula = "=CHIINV(G5,G2)"

    xlsh.Cells(1, 4).Value = 0
    xlsh.Cells(2, 4).Formula = "=H5"
    xlsh.Cells(3, 4).Formula = "=H5"
    xlsh.Cells(1, 5).Formula = "=G5"
    xlsh.Cells(2, 5).Formula = "=G5"
    xlsh.Cells(3, 5).Value = 0
    XLapp.visible = True
    On Error Resume Next
    Err.Number = 0
    xlsh.Range("A2").AutoFill Destination:=xlsh.Range("A2:A100"), Type:=0 'xlFillDefault
    xlsh.Range("B1").AutoFill Destination:=xlsh.Range("B1:B100"), Type:=0 'xlFillDefault
    If Err.Number <> 0 Then ' For unknown reasons, autofill may fail every other time, so it must be done manually, but it is slower.
        Err.Number = 0
        For i = 1 To 100
            xlsh.Cells(1 + i, 1).Formula = "=G$3/100+A" & i
            xlsh.Cells(1 + i, 2).Formula = "=CHIDIST(A" & i + 1 & ",G$2)"
        Next
    End If
    Set ch = xlsh.Shapes.AddChart(xlXYScatterSmoothNoMarkers, 5, 100, 400, 200).Chart '    xlsh.Shapes.AddChart.Select
    ch.SetElement (msoElementLegendNone)
    ch.Axes(xlValue).MaximumScale = 1
    ch.Axes(xlValue).MajorUnit = 0.25
    ch.Axes(xlValue).MinorTickMark = xlInside
    
    '    xlsh.ChartObjects(1).ChartType = xlXYScatterSmoothNoMarkers
    ch.SetSourceData Source:=xlsh.Range("'Ark1'!$A$1:$B$100")
    ch.SeriesCollection.NewSeries
    ch.SeriesCollection(2).Name = "=""Marker"""
    ch.SeriesCollection(2).ChartType = xlXYScatterLinesNoMarkers
    ch.SeriesCollection(2).Format.Line.DashStyle = msoLineSysDash
    ch.SeriesCollection(2).Format.Line.EndArrowheadStyle = msoArrowheadTriangle
    ch.SeriesCollection(2).XValues = "='Ark1'!$D$1:$D$3"
    ch.SeriesCollection(2).Values = "='Ark1'!$E$1:$E$3"


    xlsh.visible = -1 'xlSheetVisible
    'xlsh.visible = xlSheetVisible

    'xlapp.EnableEvents = False

    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
    XLapp.ScreenUpdating = True
    XLapp.EnableEvents = True
End Sub

Function betcif(Optional ByVal tal As Double = 1, Optional ByVal cif As Integer = 100, Optional ByVal rundop As Integer = 0) As Double
' Returns numbers with cif significant digits.
' if no number of significant digits is specified, 5 is used
    Dim p As Integer
    On Error GoTo Fejl
  
    If tal = 0 Then
        betcif = 0
        GoTo slut
    End If
    p = Int(Log(Abs(tal)) / Log(10))
    If rundop = 0 Then ' normal rounding
        betcif = Round(tal * 10 ^ (cif - 1 - p)) / 10 ^ -(p - cif + 1)
    ElseIf rundop = 1 Then ' always round up
        betcif = tal * 10 ^ (cif - 1 - p)
        If Round(betcif) < betcif Then
            betcif = Round(betcif) + rundop
        Else
            betcif = Round(betcif)
        End If
        betcif = betcif * 10 ^ (p - cif + 1)
    ElseIf rundop = 2 Then ' always round down
        betcif = Int(tal * 10 ^ (cif - 1 - p)) / 10 ^ -(p - cif + 1)
    Else
        betcif = tal * 10 ^ (cif - 1 - p)
        If Round(betcif) > betcif Then
            betcif = Int(betcif) + rundop
        Else
            betcif = Round(betcif)
        End If
        betcif = betcif * 10 ^ (p - cif + 1)
    End If

    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Function
Function OpenExcelWB(FilNavn As String, Optional startark As String, Optional WorkBookName As String) As Object
On Error Resume Next
#If Mac Then
    OpenExcelMac FilNavn & ";" & startark
#Else
Dim xclapp As Object 'Excel.Application
Set xclapp = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
    Set xclapp = CreateObject("Excel.Application")
End If
Dim wordmatsti As String
xclapp.visible = True
wordmatsti = GetProgramFilesDir & "\WordMat\Excelfiles\" & FilNavn
If Dir(wordmatsti) = "" Then
    wordmatsti = Environ("AppData") & "\WordMat\Excelfiles\" & FilNavn
End If

If Dir(wordmatsti) <> "" Then
    Set OpenExcelWB = xclapp.Workbooks.Add(wordmatsti)
    If WorkBookName <> "" Then
        OpenExcelWB.Windows(1).Caption = WorkBookName
    End If
'    OpenExcelWB.Windows(1).WindowState = xlMaximized
    OpenExcelWB.Windows(1).WindowState = -4137 'xlMaximized
    xclapp.Run ("Auto_open")
    If startark <> "" Then
        OpenExcelWB.Sheets(startark).Activate
    End If
End If
#End If
End Function
Function GetExcelSti() As String
    GetExcelSti = GetProgramFilesDir & "\Microsoft Office\Office14"
    If Dir(GetExcelSti, vbDirectory) = "" Then
        GetExcelSti = GetProgramFilesDir & "\Microsoft Office\Office12"
    End If
    If Dir(GetExcelSti, vbDirectory) <> "" Then
        GetExcelSti = """" & GetExcelSti & "\Excel.exe"""
    Else
        GetExcelSti = ""
    End If
End Function
Function InsertOpenExcel(FilNavn As String, Optional startark As String = "", Optional WorkBookName As String) As Object
' inserts embedded or opens depending on setting
On Error GoTo Fejl
    If ExcelIndlejret Then
        Set InsertOpenExcel = InsertIndlejret(FilNavn, startark)
    Else
        Set InsertOpenExcel = OpenExcelWB(FilNavn, startark, WorkBookName)
    End If

Fejl:
End Function
Sub Chi2Graf()
    InsertOpenExcel FilNavn:="Chi2Fordeling.xltm", WorkBookName:=Sprog.A(483)
End Sub
Sub NormalFordelingGraf()
    InsertOpenExcel FilNavn:="NormalFordeling.xltm", WorkBookName:=Sprog.A(482)
End Sub
Sub BinomialFordeling()
    InsertOpenExcel FilNavn:="BinomialFordeling.xltm", WorkBookName:=Sprog.A(585)
End Sub
Sub GoodnessofFit()
    InsertOpenExcel "statistik.xltm", "GOF", "Goodness of fit"
End Sub
Sub BinomialTest()
    If Sprog.SprogNr = 1 Then
        InsertOpenExcel "BinomialFordeling.xltm", "Binomialtest", Sprog.A(593)
    Else
        InsertOpenExcel "BinomialFordeling.xltm", "Binomial test", Sprog.A(593)
    End If
End Sub

