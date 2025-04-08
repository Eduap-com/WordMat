Attribute VB_Name = "Regression"
Option Explicit

Sub linregression()
' is executed from the menu. Table must be selected
    Dim Cregr As New CRegression
    Application.ScreenUpdating = False
On Error GoTo Fejl
    SaveBackup
    If Selection.OMaths.Count > 0 And Selection.Tables.Count = 0 Then
        Cregr.GetSetData
        Selection.OMaths(Selection.OMaths.Count).Range.Select
        Selection.Collapse wdCollapseEnd
        Selection.TypeParagraph
    ElseIf Selection.Tables.Count > 0 Then
        Cregr.GetTableData
    Else
        InsertTabel
        GoTo slut
    End If
    Cregr.ComputeLinRegr
    Cregr.InsertEquation
GoTo slut
Fejl:
    MsgBox Sprog.A(26), vbOKOnly, Sprog.Error
slut:
End Sub
Sub ekspregression()
    Dim Cregr As New CRegression
    Application.ScreenUpdating = False
On Error GoTo Fejl
    SaveBackup
    If Selection.OMaths.Count > 0 And Selection.Tables.Count = 0 Then
        Cregr.GetSetData
        Selection.OMaths(Selection.OMaths.Count).Range.Select
        Selection.Collapse wdCollapseEnd
        Selection.TypeParagraph
    ElseIf Selection.Tables.Count > 0 Then
        Cregr.GetTableData
    Else
        InsertTabel
        GoTo slut
    End If
    Cregr.ComputeExpRegr
    Cregr.InsertEquation
GoTo slut
Fejl:
    MsgBox Sprog.A(26), vbOKOnly, Sprog.Error
slut:
End Sub
Sub potregression()

On Error GoTo Fejl
    Dim Cregr As New CRegression
    SaveBackup
    Application.ScreenUpdating = False
        
    If Selection.OMaths.Count > 0 And Selection.Tables.Count = 0 Then
        Cregr.GetSetData
        Selection.OMaths(Selection.OMaths.Count).Range.Select
        Selection.Collapse wdCollapseEnd
        Selection.TypeParagraph
    ElseIf Selection.Tables.Count > 0 Then
        Cregr.GetTableData
    Else
        InsertTabel
        GoTo slut
    End If
    Cregr.ComputePowRegr
    Cregr.InsertEquation

GoTo slut
Fejl:
    MsgBox Sprog.A(26), vbOKOnly, Sprog.Error
slut:
End Sub
Sub polregression()

On Error GoTo Fejl
    Dim Cregr As New CRegression
    SaveBackup
    Application.ScreenUpdating = False
    
    If Selection.OMaths.Count > 0 And Selection.Tables.Count = 0 Then
        Cregr.GetSetData
        Selection.OMaths(Selection.OMaths.Count).Range.Select
        Selection.Collapse wdCollapseEnd
        Selection.TypeParagraph
    ElseIf Selection.Tables.Count > 0 Then
        Cregr.GetTableData
    Else
        InsertTabel
        GoTo slut
    End If
    Cregr.ComputePolRegr
    Cregr.InsertEquation
GoTo slut
Fejl:
    MsgBox Sprog.A(26), vbOKOnly, Sprog.Error
slut:
End Sub
Sub UserRegression()
On Error GoTo Fejl
    Dim Cregr As New CRegression
    Dim sslut As Long
    Application.ScreenUpdating = False
    sslut = Selection.End
    
    PrepareMaxima
    If Selection.Tables.Count > 0 Then
        If Selection.OMaths.Count > 0 Then
'            Selection.OMaths(Selection.OMaths.count).Range.Select
            omax.ReadSelection
            omax.Kommando = omax.ConvertToAscii(omax.Kommando)
        End If
        Cregr.GetTableData
        Selection.Collapse wdCollapseEnd
        Selection.TypeParagraph
    Else
        InsertTabel
        GoTo slut
    End If
    Cregr.ComputeUserRegr
    If Selection.OMaths.Count > 0 Then
        Selection.End = sslut
        Selection.start = sslut
        Selection.OMaths(Selection.OMaths.Count).Range.Select
        Selection.Collapse wdCollapseEnd
        Selection.TypeParagraph
    End If
    Cregr.InsertEquation
GoTo slut
Fejl:
    MsgBox Sprog.A(26), vbOKOnly, Sprog.Error
slut:
End Sub
Sub InsertTabel()
    Dim antalp As Integer
    Application.ScreenUpdating = False
    SaveBackup
    antalp = val(InputBox(Sprog.A(24), Sprog.A(202), ""))
    If antalp = 0 Then Exit Sub
        
    If antalp > 200 Then
        MsgBox Sprog.A(25)
    ElseIf antalp > 0 Then
        Selection.Collapse wdCollapseEnd
        
        If Selection.OMaths.Count > 0 Then
            Selection.OMaths(1).Range.Select
            Selection.Collapse wdCollapseEnd
            Selection.TypeParagraph
        End If
        
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
        
        If antalp <= 10 Then
            ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=2, NumColumns:=antalp + 1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
            With Selection.Tables(1)
#If Mac Then
#Else
                .ApplyStyleHeadingRows = True
                .ApplyStyleLastRow = False
                .ApplyStyleFirstColumn = True
                .ApplyStyleLastColumn = False
                .ApplyStyleRowBands = True
                .ApplyStyleColumnBands = False
#End If
                .Cell(1, 1).Range.Text = "x"
                .Cell(1, 1).Range.Bold = True
                .Cell(2, 1).Range.Text = "y"
                .Cell(2, 1).Range.Bold = True
                .Cell(1, 2).Range.Select
                .Columns(1).Width = 30
            End With
        Else
            ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=antalp + 1, NumColumns:=2, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
            With Selection.Tables(1)
#If Mac Then
#Else
                .ApplyStyleHeadingRows = True
                .ApplyStyleLastRow = False
                .ApplyStyleFirstColumn = True
                .ApplyStyleLastColumn = False
                .ApplyStyleRowBands = True
                .ApplyStyleColumnBands = False
#End If
                .Cell(1, 1).Range.Text = "x"
                .Cell(1, 1).Range.Bold = True
                .Cell(1, 2).Range.Text = "y"
                .Cell(1, 2).Range.Bold = True
                .Cell(2, 1).Range.Select
                .Columns(1).Width = 65
                .Columns(2).Width = 65
            End With
        End If
    End If
        
    Oundo.EndCustomRecord

End Sub

