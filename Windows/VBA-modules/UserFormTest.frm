VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormTest 
   Caption         =   "UserForm1"
   ClientHeight    =   7650
   ClientLeft      =   30
   ClientTop       =   165
   ClientWidth     =   13215
   OleObjectBlob   =   "UserFormTest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ErrorCount As Long, OkCount As Long
Dim TabNo As Integer, RowNo As Integer
Dim StopTest As Boolean

Private Sub CommandButton_luk_Click()
    Me.hide
End Sub


Sub TestTable(Tabel As Table)
Dim s As String
        If Tabel.Columns.Count >= 5 Then
            For RowNo = 1 To Tabel.Rows.Count
                If Not StopTest Then
                    s = VBA.LCase(Tabel.Cell(RowNo, 1).Range.text)
                    If InStr(s, "auto") > 0 Then
                        MaximaExact = 0
                    ElseIf InStr(s, "exact") > 0 Then
                        MaximaExact = 1
                    ElseIf InStr(s, "num") > 0 Then
                        MaximaExact = 2
                    End If
                    If InStr(s, "beregn") > 0 Then
                        PrepareRow Tabel, RowNo
                        beregn
                        SetRow Tabel, RowNo
                    ElseIf InStr(s, "desolve") > 0 Then
                        PrepareRow Tabel, RowNo
                        SolveDEpar "y", "x"
                        SetRow Tabel, RowNo
                    ElseIf InStr(s, "solve") > 0 Then
                        PrepareRow Tabel, RowNo
                        MaximaSolvePar "x"
                        SetRow Tabel, RowNo
                    End If
                Else
                    GoTo slut
                End If
            Next
        End If
slut:
End Sub

Private Sub CommandButton_nexttable_Click()
Dim Tabel As Table
    Nulstil
    PrepareMaxima
    Selection.GoToNext wdGoToTable
    
    Set Tabel = Selection.Tables(1)
    If Not Tabel Is Nothing Then
        TestTable Tabel
    End If
    
End Sub
'
Private Sub CommandButton_start_Click()
    Nulstil
    Application.ScreenUpdating = False
    PrepareMaxima
    AllTables
End Sub

Sub AllTables()
    Dim Tabel As Table

    TextBox_status.text = "ok/fejl | Tabel | Række | Kommando " & vbCrLf
    For TabNo = 1 To ActiveDocument.Tables.Count
        Set Tabel = ActiveDocument.Tables(TabNo)
        TestTable Tabel
    Next

    TextBox_status.text = TextBox_status.text & vbCrLf & "Test afsluttet. " & vbCrLf & "Der blev gennemført " & OkCount + ErrorCount & " test med " & ErrorCount & " fejl."
Fejl:
slut:
End Sub

Sub SetRow(t As Table, r As Integer)
Dim s As String
    
    s = " | " & TabNo & " | " & r & " | " & Replace(Replace(t.Cell(r, 3).Range.text, vbCrLf, ""), vbCr, "") & vbCrLf
    If t.Cell(RowNo, 3).Range.text = t.Cell(RowNo, 4).Range.text Then
        t.Cell(r, 5).Range.text = "OK"
        t.Cell(r, 5).Range.Font.ColorIndex = wdGreen
        s = " OK " & s
        OkCount = OkCount + 1
    Else
        t.Cell(r, 5).Range.text = Sprog.Error
        t.Cell(r, 5).Range.Font.ColorIndex = wdRed
        s = Sprog.Error & s
        ErrorCount = ErrorCount + 1
    End If
    TextBox_status.text = TextBox_status.text & s
    OpdaterAntal
End Sub
Sub PrepareRow(t As Table, r As Integer)
    t.Cell(r, 2).Range.Select
    Selection.Copy
    t.Cell(r, 3).Range.Paste
    t.Cell(r, 3).Range.Select
End Sub
Sub OpdaterAntal()
    Label_antalok.Caption = OkCount
    Label_antalfejl.Caption = ErrorCount
End Sub

Private Sub CommandButton_stop_Click()
    StopTest = True
End Sub
Sub Nulstil()
    TextBox_status.text = ""
    StopTest = False
    ErrorCount = 0
    OkCount = 0
    OpdaterAntal
End Sub
Private Sub UserForm_Activate()
    Nulstil
End Sub
