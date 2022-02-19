VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Public XLapp As excel.Application
Public XLapp As Object ' getobject metoden gav problemer med at man ikke vidste hvilken
                        ' excel app der kom retur. Fidusen var ellers at gøre det hurtigere
                        ' hvis Excel allerede var åben.
                        ' men det virker ikke stabilt. WordMat skal have sin egen excel.app kørende
                        
'Public xlwb As Excel.Workbook
Public xlwb As Object
'Public ch As excel.Chart
Public ch As Object
Public p As Double
Public sum As Double
Public Below5 As Boolean
Public vis As Boolean ' visible?
'Private tabel() As Double
Private Tabel As DocData

Private Sub Class_Initialize()

'Set xlapp = CreateObject("Excel.Application")
'Set XLapp = New Excel.Application
vis = True
Set Tabel = New DocData
'Statistik.xltm

End Sub
Private Sub Class_Terminate()
'    CloseExcel
End Sub
Sub NewFile()
If XLapp Is Nothing Then
'    Set XLapp = New excel.Application
    Set XLapp = CreateObject("Excel.Application") 'excel.Application
End If
XLapp.visible = vis
Set xlwb = XLapp.Workbooks.Add
End Sub
Sub LoadFile(filnavn As String)
Dim wordmatsti As String
On Error Resume Next
If XLapp Is Nothing Then
'    Set XLapp = New excel.Application
    Set XLapp = CreateObject("Excel.Application")
End If
XLapp.visible = vis

'If XLapp Is Nothing Then
'    Set XLapp = CreateObject("Excel.Application")
'End If
'Set XLapp = GetObject(, "Excel.Application")
'If Err.Number <> 0 Then
'    Set XLapp = CreateObject("Excel.Application")
'End If
    XLapp.visible = True
wordmatsti = GetProgramFilesDir & "\WordMat\ExcelFiles\" & filnavn & ""
If Dir(wordmatsti) <> "" Then
    Set xlwb = XLapp.Workbooks.Add(wordmatsti)
End If


End Sub
Public Sub CloseExcel()
On Error GoTo fejl
xlwb.Close False ' close the workbook without saving
XLapp.quit ' close the Excel application

fejl:
Set xlwb = Nothing
Set XLapp = Nothing

End Sub
Public Sub RestartExcel()
    On Error Resume Next
    CloseExcel
    Set XLapp = CreateObject("Excel.Application")
'    Set XLapp = New excel.Application

End Sub
Public Sub UpdateAll()
On Error Resume Next
    If Not XLapp Is Nothing Then
        XLapp.Run ("UpDateAll")
    End If
End Sub
Public Sub PrePareExcel()
On Error Resume Next

If XLapp Is Nothing Then
    Set XLapp = GetObject(, "Excel.Application")
'    Set XLapp = New excel.Application
End If
If Err.Number <> 0 Then
    Set XLapp = CreateObject("Excel.Application")
'    Set XLapp = New excel.Application
End If
GoTo slut
fejl:
    RestartExcel
slut:
End Sub
Public Sub Freeze()
On Error Resume Next
    XLapp.EnableEvents = False
    XLapp.ScreenUpdating = False
    XLapp.DisplayAlerts = False
End Sub
Public Sub UnFreeze()
On Error Resume Next
    XLapp.EnableEvents = True
    XLapp.ScreenUpdating = True
    XLapp.DisplayAlerts = True
End Sub
Function InsertIndlejret(filnavn As String, Optional startark As String)
'indsætter exceldokument som indlejret dokument
Dim path As String
Dim ils As InlineShape
Dim vers As String
On Error Resume Next
'Set XLapp = GetObject(, "Excel.Application")
'If Err.Number <> 0 Then
'    Set XLapp = CreateObject("Excel.Application")
'    Set XLapp = New Excel.Application
'End If
On Error GoTo slut
Application.ScreenUpdating = False
EnableExcelMacros
    
'    Dim ufwait2 As New UserFormWaitForMaxima
'    ufwait2.CommandButton_stop.visible = False
'    ufwait2.Label_tip.Caption = "      Indsætter indlejret..."
'    ufwait2.Label_progress.Caption = "Indsætning af indlejrede objekter kan tage tid. Dobbeltklik på objektet for at redigere det."
'    ufwait2.Show vbModeless
        DoEvents
'        ufwait2.Label_progress = "***"
path = """" & GetProgramFilesDir & "\WordMat\ExcelFiles\" & filnavn & """"
    If Selection.Range.Tables.Count > 0 Then
'        Selection.Copy
        Selection.Tables(Selection.Tables.Count).Select
        Selection.Collapse (wdCollapseEnd)
        Selection.TypeParagraph
'        startark = "Data"
    End If
    If Selection.OMaths.Count > 0 Then
        Selection.OMaths(Selection.OMaths.Count).Range.Select
        Selection.Collapse (wdCollapseEnd)
        Selection.TypeParagraph
    End If
If val(Application.Version) = 12 Then
    vers = ".12"
Else
    vers = ""
End If

Set ils = ActiveDocument.InlineShapes.AddOLEObject(ClassType:="Excel.Sheet" & vers & Application.Version, _
FileName:=path, LinkToFile:=False, DisplayAsIcon:=False, Range:=Selection.Range)
'        ufwait2.Label_progress = "***************************************"
'Ils.OLEFormat.DoVerb (wdOLEVerbOpen)
'ils.OLEFormat.DoVerb (wdOLEVerbInPlaceActivate)
'ils.OLEFormat.DoVerb (wdOLEVerbShow)
If startark <> "" Then
    ils.OLEFormat.DoVerb (wdOLEVerbInPlaceActivate)
    DoEvents
    Set xlwb = ils.OLEFormat.Object

'    Dim oWS As Object ' Worksheet Object
'    Set oWS = ils.OLEFormat.Object
'    ils.OLEFormat.Object.Sheets(startark).Activate
'    oWS.ActiveSheet.Cells(4, 1).Activate
'    oWS.ActiveSheet.Application.Selection.Paste ' virker ikke
'    oWS.Selection.Paste
End If
'Unload ufwait2
'Ils.OLEFormat.DoVerb (wdOLEVerbUIActivate)
'Ils.OLEFormat.DoVerb (wdOLEVerbInPlaceActivate)
'Ils.OLEFormat.DoVerb (wdOLEVerbHide)
DisableExcelMacros
GoTo slut
fejl:
    On Error Resume Next
    MsgBox Sprog.A(98), vbOKOnly, Sprog.Error
'    Unload ufwait2
slut:
End Function


Public Sub Chi2Test(Optional signiv As Double = 0.05, Optional ByVal r As Integer = 0, Optional ByVal c As Integer = 0)
'    Dim c As Integer
'    Dim r As Integer
    Dim i As Integer
    Dim j As Integer
    On Error Resume Next
'XLapp.ScreenUpdating = False
    If r = 0 Or c = 0 Then
        r = Tabel.nrows
        c = Tabel.ncolumns
    End If
    
    
'    Dim ws As Worksheet
    Dim ws As Object 'Worksheet
    Set ws = xlwb.ActiveSheet
    
    For i = 1 To r
    For j = 1 To c 'tabel.ncolumns
        ws.Cells(i + 2, j + 1).Formula = Tabel.TabelCelle(i, j)
        ws.Range(ws.Cells(i + 2, j + 1), ws.Cells(i + 2, j + 1)).BorderAround ColorIndex:=1, Weight:=xlThin
'        ws.Range("A1").Offset(i, j).Text = tabel(i, j)
    Next
    Next
    
'    For i = 1 To tabel.nrows
'    For j = 1 To tabel.ncolumns
'        ws.Cells(i, j).Formula = tabel.TabelCelle(i, j)
''        ws.Range("A1").Offset(i, j).Text = tabel(i, j)
'    Next
'    Next
    
    'tæl rækker sammen
    For i = 1 To r
        ws.Cells(i + 2, c + 2).Formula = "=sum(R" & i + 2 & "C2:R" & i + 2 & "C" & c + 1 & ")"
        ws.Range(ws.Cells(i + 2, c + 2), ws.Cells(i + 2, c + 2)).BorderAround ColorIndex:=1, Weight:=xlThin
    Next
    
    'tæl kolonner sammen
    For i = 1 To c
        ws.Cells(r + 3, i + 1).Formula = "=sum(R3C" & i + 1 & ":R" & r + 2 & "C" & i + 1 & ")"
        ws.Range(ws.Cells(r + 3, i + 1), ws.Cells(r + 3, i + 1)).BorderAround ColorIndex:=1, Weight:=xlThin
    Next
    'total
    ws.Cells(r + 3, c + 2).Formula = "=sum(R" & r + 3 & "C2:R" & r + 3 & "C" & c + 1 & ")"
    ws.Range(ws.Cells(r + 3, c + 2), ws.Cells(r + 3, c + 2)).BorderAround ColorIndex:=1, Weight:=xlThin
    
    ' forventede værdier
    For i = 3 To r + 2
    For j = 2 To c + 1
'        ws.Cells(i, c + 2 + j).Formula = ws.Cells(r + 1, j).Value * ws.Cells(i, c + 1).Value / ws.Cells(r + 1, c + 1).Value
        ws.Cells(i, c + 2 + j).Formula = "=if(R" & r + 3 & "C" & c + 2 & ">0,R" & r + 3 & "C" & j & "*R" & i & "C" & c + 2 & "/R" & r + 3 & "C" & c + 2 & ",0)"
        ws.Range(ws.Cells(i, c + 2 + j), ws.Cells(i, c + 2 + j)).BorderAround ColorIndex:=1, Weight:=xlThin
'        ws.Range("A1").Offset(i, j).Text = tabel(i, j)
    Next
    Next
    
    'tæl rækker sammen
    For i = 1 To r
        ws.Cells(i + 2, 2 * c + 4).Formula = "=sum(R" & i + 2 & "C" & c + 4 & ":R" & i + 2 & "C" & 2 * c + 3 & ")"
        ws.Range(ws.Cells(i + 2, 2 * c + 4), ws.Cells(i + 2, 2 * c + 4)).BorderAround ColorIndex:=1, Weight:=xlThin
    Next
    
    'tæl kolonner sammen
    For i = 1 To c + 1
        ws.Cells(r + 3, i + 3 + c).Formula = "=sum(R3C" & i + 3 + c & ":R" & r + 2 & "C" & i + 3 + c & ")"
        ws.Range(ws.Cells(r + 3, i + 3 + c), ws.Cells(r + 3, i + 3 + c)).BorderAround ColorIndex:=1, Weight:=xlThin
    Next
    'total
'    ws.Cells(r + 3, c + 2).Formula = "=sum(R" & r + 3 & "C2:R" & r + 3 & "C" & c + 1 & ")"
    
    ws.Calculate
    ' Chi2 beregning af p-værdi
'    ws.Cells(2, 2 * c + 4).Formula = "=a1"
'    ws.Cells(2, 2 * c + 4).Formula = "=CHITEST(R1C1:R2C2;R1C5:R2C6)"
'    ws.Cells(2, 8).Formula = "=sum(A1:B2)"
'On Error Resume Next
'    ws.Cells(2, 8).FormulaR1C1 = "=CHITEST(R1C1:R2C2,R1C5:R2C6)"
'    ws.Cells(2, 8).FormulaR1C1 = "=CHITEST(RC[-5]:R[1]C[-4],RC[-3]:R[1]C[-2])"

'    ws.Cells(2, 8).Formula = "=CHITEST(A1:B2,E1:F2)"
    ws.Cells(r + 5, 2).Formula = "=CHITEST(R3C2:R" & r + 2 & "C" & c + 1 & ",R3C" & c + 4 & ":R" & r + 2 & "C" & 2 * c + 3 & ")"
    
    p = ws.Cells(r + 5, 2).Value
    p = betcif(p, MaximaCifre)
'    ws.Cells(r + 5, 2).Style = "Percent"
    ws.Cells(r + 5, 2).NumberFormat = "0.00%"
    ws.Cells(r + 5, 2).Font.Bold = True
    
    
    ws.Cells(1, 1).Value = VBA.ChrW(&H3C7) & VBA.ChrW(&HB2) & "-test"
    ws.Cells(2, 2).Value = Sprog.A(400)
    ws.Cells(2, c + 4).Value = Sprog.A(401)
    ws.Cells(r + 3, 1).Value = Sprog.A(402)
    ws.Cells(r + 3, 1).HorizontalAlignment = xlRight
    
    ws.Cells(2, c + 2).Value = Sprog.A(402)
    ws.Cells(2, 2 * c + 4).Value = Sprog.A(402)
    
    ws.Cells(r + 5, 1).Value = Sprog.A(403) & ":"
    ws.Cells(r + 6, 1).Value = Sprog.A(350) & ":"
    ws.Cells(r + 6, 2).Value = signiv
    ws.Cells(r + 6, 2).Style = Sprog.RibPercentage
    
    ws.Cells(r + 5, 4).Formula = "=if(R" & r + 5 & "C2<R" & r + 6 & "C2,""" & Sprog.A(404) & """,""" & Sprog.A(405) & """)"
    ws.Cells(r + 6, 4).Formula = "=if(COUNTIF(R3C2:R" & r + 2 & "C" & c + 1 & ",""<5"")>0,""" & Sprog.A(358) & ""","""")"
    ws.Cells(r + 7, 4).Formula = "=if(sum(R3C2:R" & r + 2 & "C" & c + 1 & ")<50,""" & Sprog.A(406) & ""","""")"
    ws.Cells(r + 5, 4).Font.Bold = True
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    
    
    ws.Range(ws.Cells(3, 2), ws.Cells(r + 2, c + 1)).BorderAround ColorIndex:=1, Weight:=xlThick
    ws.Range(ws.Cells(3, 2), ws.Cells(r + 3, c + 2)).BorderAround ColorIndex:=1, Weight:=xlThick
    
    ws.Range(ws.Cells(3, c + 4), ws.Cells(r + 2, 2 * c + 3)).BorderAround ColorIndex:=1, Weight:=xlThick
    ws.Range(ws.Cells(3, c + 4), ws.Cells(r + 3, 2 * c + 4)).BorderAround ColorIndex:=1, Weight:=xlThick
    
    ws.Range(ws.Cells(r + 3, c + 2), ws.Cells(r + 3, c + 2)).BorderAround ColorIndex:=1, Weight:=xlThick
    ws.Range(ws.Cells(r + 3, 2 * c + 4), ws.Cells(r + 3, 2 * c + 4)).BorderAround ColorIndex:=1, Weight:=xlThick
    
    ws.Columns("A:A").ColumnWidth = 15
    
    ws.Cells(r + 8, 1).Value = Sprog.A(360) & ":"
    ws.Cells(r + 8, 2).Value = (r - 1) * (c - 1)
    ws.Cells(r + 9, 1).Value = Sprog.A(407) & ":"
    ws.Cells(r + 9, 2).Formula = "=sum(R" & r + 9 & "C3:R" & r + 9 & "C" & r * c + 2 & ")"
    
    Dim k As Integer
    For i = 1 To r
    For j = 1 To c
'        ws.Cells(r + 9, 3 + k).Formula = "=if(R" & 2 + i & "C" & c + 3 + j & ">0,(R" & 2 + i & "C" & 1 + j & "-R" & 2 + i & "C" & c + 3 + j & ")^2/R" & 2 + i & "C" & c + 3 + j & ",0)"
'        MsgBox "=if(R" & 2 + i & "C" & c + 3 + j & ">0,(R" & 2 + i & "C" & 1 + j & "-R" & 2 + i & "C" & c + 3 + j & ")^2/R" & 2 + i & "C" & c + 3 + j & ",0)"
        ws.Cells(r + 9, 3 + k).Formula = "=if(R" & 2 + i & "C" & c + 3 + j & ">0,(R" & 2 + i & "C" & 1 + j & "-R" & 2 + i & "C" & c + 3 + j & ")^2/R" & 2 + i & "C" & c + 3 + j & ",0)"
        k = k + 1
    Next
    Next
    
    With ws.Range(ws.Cells(r + 9, 3), ws.Cells(r + 9, r * c + 2)).Font
        .Size = 8
        .Color = RGB(170, 170, 170)
    End With
    
    
    ws.Cells(r + 10, 1).Value = Sprog.A(408) & ":"
    ws.Cells(r + 10, 2).Formula = "=chiinv(R" & r + 6 & "C2,R" & r + 8 & "C2)"
    XLapp.ScreenUpdating = True

End Sub
Public Sub GetTableDataregr()
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
        cellval = rws(i).Cells(j).Range.text
        cellval = Left(cellval, Len(cellval) - 1)
        cellval = Replace(cellval, "*10^", "E")
        tal = CDbl(cellval)
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

        cellval = Replace(rws(j).Cells(1 + oFS).Range.text, ",", ".")
'        cellval = tabel.Cell(j, 1).Range.text
        cellval = Left(cellval, Len(cellval) - 1)
        cellval = Replace(cellval, "*10^", "E")
        Tabel(i, 1) = CDbl(cellval) ' skal med for at melde fejl , men alene duer ikke da den er afh af sprog indstill
        Tabel(i, 1) = val(cellval) 'melder ikke fejl ved bogstaver
        cellval = Replace(rws(j).Cells(2 + oFS).Range.text, ",", ".")
'        cellval = tabel.Cell(j, 2).Range.text
        cellval = Left(cellval, Len(cellval) - 1)
        cellval = Replace(cellval, "*10^", "E")
        Tabel(i, 2) = CDbl(cellval)
        Tabel(i, 2) = val(cellval)
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
Public Function GetTableData() As Boolean
Dim i As Integer
Dim j As Integer
    GetTableData = Tabel.ReadSelection()
    
    For i = 1 To Tabel.nrows
    For j = 1 To Tabel.ncolumns
        sum = sum + Tabel.TabelCelle(i, j)
        If Tabel.TabelCelle(i, j) < 5 Then Below5 = True
    Next
    Next
    
End Function

Public Sub CopyTableToXL()
    Dim c As Integer
    Dim r As Integer
    Dim i As Integer
    Dim j As Integer
'    r = UBound(tabel, 1)
'    c = UBound(tabel, 2)
'    Dim ws As Worksheet
    Dim ws As Object 'Worksheet
    Set ws = xlwb.ActiveSheet

    For i = 1 To Tabel.ncolumns
    For j = 1 To Tabel.nrows
        ws.Range("a1").Offset(j - 1, i - 1).Value = Tabel(j, i)
    Next
    Next
    
    ws.Range("a1:b5").Select
End Sub




