Attribute VB_Name = "ModuleCode"
Option Explicit

Function GetCodeFileName() As String
#If Mac Then
    GetCodeFileName = DataFolder & "WordMatCodeFile.mac"
#Else
    GetCodeFileName = Environ("AppData") & "\WordMat\WordMatCodeFile.mac"
#End If
End Function
Function GetCodeFileText() As String
    GetCodeFileText = ReadTextfileToString(GetCodeFileName)
End Function

Sub SaveCodeFileText(t As String)
    WriteTextfileToString GetCodeFileName, t
End Sub

Sub InsertCodeBlock()
    Dim cc As ContentControl
    Dim codeText As String
    On Error GoTo fejl
    
    codeText = TT.A(907) & vbCrLf & vbCrLf & " "

    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord

    Selection.TypeParagraph
    ' Add a rich text content control at the current selection
    Set cc = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
    ' Set tag and title for identification
    cc.Title = "CodeBlock"
    cc.Tag = "CodeBlock"
    
    ' Insert code text
    cc.Range.text = codeText

    ' Apply Consolas font and optional styling
    With cc.Range.Font
        .Name = "Consolas"
        .Bold = False
        .ColorIndex = wdAuto
        .Size = 10
    End With
    
    cc.Range.ParagraphFormat.SpaceAfter = 0
    
    cc.Range.Shading.BackgroundPatternColor = RGB(240, 240, 240)
    
    With cc.Range.ParagraphFormat.Borders(wdBorderTop)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth050pt
        .Color = wdColorGray25
    End With
    With cc.Range.ParagraphFormat.Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth050pt
        .Color = wdColorGray25
    End With
    With cc.Range.ParagraphFormat.Borders(wdBorderLeft)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth050pt
        .Color = wdColorGray25
    End With
    With cc.Range.ParagraphFormat.Borders(wdBorderRight)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth050pt
        .Color = wdColorGray25
    End With
    cc.Range.NoProofing = True
    cc.Range.Select
    Oundo.EndCustomRecord
    
    If Not UseCodeBlocks Then
        MsgBox2 TT.A(908), vbOKOnly
    End If
    
    GoTo slut
fejl:
    ActiveDocument.Undo
    MsgBox2 TT.A(910), vbOKOnly, TT.Error
slut:

End Sub

Function GetAllPreviousCodeBlocks() As String
    Dim selStart As Long
    Dim cc As ContentControl
    Dim i As Long
    Dim result As String, s As String

    selStart = Selection.Range.start
    result = ""

    ' Loop through all content controls from last to first
    For i = ActiveDocument.ContentControls.Count To 1 Step -1
        Set cc = ActiveDocument.ContentControls(i)

        ' Check if it's a code block and before the cursor
        If cc.Tag = "CodeBlock" And cc.Range.End < selStart Then
            ' Prepend the code block text to maintain order from nearest to farthest
            
            s = cc.Range.text
            s = TrimR(Trim$(s), vbCrLf)
            s = TrimR(Trim$(s), vbCr)
            s = TrimR(Trim$(s), vbLf)
            If s <> vbNullString Then
                If right$(s, 1) <> ";" And right$(s, 1) <> "$" Then s = s & "$"
                If result = "" Then
                    result = s
                Else
                    result = s & vbCrLf & result
                End If
            End If
        End If
    Next i

    GetAllPreviousCodeBlocks = result
End Function

