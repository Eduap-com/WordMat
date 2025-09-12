Attribute VB_Name = "Regression"
Option Explicit

Sub linregression()
' is executed from the menu. Table must be selected
    Dim Cregr As New CRegression
    Application.ScreenUpdating = False
On Error GoTo fejl
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
fejl:
    MsgBox TT.A(26), vbOKOnly, TT.Error
slut:
End Sub
Sub ekspregression()
    Dim Cregr As New CRegression
    Application.ScreenUpdating = False
On Error GoTo fejl
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
fejl:
    MsgBox TT.A(26), vbOKOnly, TT.Error
slut:
End Sub
Sub potregression()

On Error GoTo fejl
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
fejl:
    MsgBox TT.A(26), vbOKOnly, TT.Error
slut:
End Sub
Sub polregression()

On Error GoTo fejl
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
fejl:
    MsgBox TT.A(26), vbOKOnly, TT.Error
slut:
End Sub
Sub FitSin()
    If GraphApp = 2 Then
    Else
        FitSinGeoGebraSuite
    End If
End Sub

Sub UserRegression()
On Error GoTo fejl
    Dim Cregr As New CRegression
    Dim sslut As Long, fkt As String, r As Range
    Application.ScreenUpdating = False
        
    sslut = Selection.End
    Set r = Selection.Range
    PrepareMaxima
    
    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord
    
    If Selection.Tables.Count > 0 Then
        If Selection.OMaths.Count > 0 Then
            omax.ReadSelection
            fkt = omax.ConvertToAscii(omax.Kommando)
        End If
        Cregr.GetTableData
        r.Select
        Selection.Collapse wdCollapseEnd
        Selection.TypeParagraph
    Else
        InsertTabel
        GoTo slut
    End If
    omax.Kommando = fkt
    If Not Cregr.ComputeUserRegr Then GoTo slut
    
    If Selection.OMaths.Count > 0 Then
        Selection.End = sslut
        Selection.start = sslut
        Selection.OMaths(Selection.OMaths.Count).Range.Select
        Selection.Collapse wdCollapseEnd
        Selection.TypeParagraph
    End If
    Selection.TypeText TT.A(33) & ":  "
    Selection.OMaths.Add Selection.Range
    Selection.OMaths(1).Range.text = Replace(fkt, "*", MaximaGangeTegn)
    Selection.OMaths(1).BuildUp
    Selection.OMaths(1).Range.Select
    Selection.Collapse wdCollapseEnd
    Selection.MoveRight wdCharacter, 1
    Selection.TypeText "  " & TT.A(34) & ": "
    Oundo.EndCustomRecord
    Cregr.InsertEquation
GoTo slut
fejl:
    Oundo.EndCustomRecord
    MsgBox TT.A(26), vbOKOnly, TT.Error
slut:
    Oundo.EndCustomRecord
End Sub
Sub InsertTabel()
    Dim antalp As Integer, s As String
    Application.ScreenUpdating = False
    SaveBackup
'    antalp = val(InputBox(TT.A(24), TT.A(202), ""))

    UserFormInputBox.MsgBoxStyle = vbOKCancel
    UserFormInputBox.prompt = TT.A(24)
    UserFormInputBox.Title = TT.A(202)
    UserFormInputBox.MultiLine = False
    UserFormInputBox.SetDefaultInput 10
    UserFormInputBox.Show
    If UserFormInputBox.MsgBoxResult = vbCancel Then Exit Sub
    s = UserFormInputBox.InputString
    If IsNumeric(s) Then
        antalp = CInt(s)
    End If
    If antalp <= 0 Then Exit Sub
        
    If antalp > 200 Then
        MsgBox2 TT.A(25)
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
                .Cell(1, 1).Range.text = "x"
                .Cell(1, 1).Range.Bold = True
                .Cell(2, 1).Range.text = "y"
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
                .Cell(1, 1).Range.text = "x"
                .Cell(1, 1).Range.Bold = True
                .Cell(1, 2).Range.text = "y"
                .Cell(1, 2).Range.Bold = True
                .Cell(2, 1).Range.Select
                .Columns(1).Width = 65
                .Columns(2).Width = 65
            End With
        End If
    End If
        
    Oundo.EndCustomRecord

End Sub

