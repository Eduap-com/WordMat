Attribute VB_Name = "LatexModule"
Option Explicit
Public HiddenDoc As Document
Public MainDoc As Document
Private LatexFilePath As String
Public ImageCounter As Long
Public HTMLindex As Integer ' whenever a html-file is generated it needs a new filename or images vil be overwritten. Then they cannot be used for preview

Public LatexCode As String
Private LAlign As Integer
Sub SaveDocToLatexPdf()
   PrepareMaxima
   SaveFile 0
End Sub
Sub SaveDocToLatexTex()
   PrepareMaxima
   SaveFile 2
End Sub


Public Sub SaveFile(doctype As Integer)
   ' 0 - pdf
   ' 1 - dvi
   ' 2 - Tex
   Dim i As Integer, s As String, p As Long
   Dim SaveSel As Range, HasBib As Boolean
   
   ' check if miktex installed
   If Not latexfil.IsMikTexInstalled Then
      MsgBox "MikTex is not installed. You will now be sent to miktex.org where you can download. After the download you will also be prompted to download some packages the first time you run the converter. Just click ok.", vbOKOnly, TT.Error
      OpenLink "https://miktex.org/download"
      GoTo slut
   End If
   '***
   
   'Check if document saved
   If ActiveDocument.path = "" Then
      MsgBox "Save your document before attempting to convert to Latex. The Latex files will be placed in a folder next to your document-file.", vbOKOnly, "File not saved"
      GoTo slut
   End If
   
   Dim UfWait As New UserFormWaitForMaxima
   Set latexfil.UfWait = UfWait
   UfWait.Label_tip.Caption = "Coverting"
   UfWait.Show vbModeless
   UserFormLatex.EventsOn = True
   latexfil.Reset
   latexfil.TitlePage = UserFormLatex.CheckBox_title.Value
   latexfil.toc = UserFormLatex.CheckBox_contents.Value
   latexfil.Titel = Split(ActiveDocument.Name, ".")(0)
   latexfil.Author = ActiveDocument.BuiltInDocumentProperties(wdPropertyAuthor)
    
   UserFormLatex.Label_status.visible = True
   '    Dim d2 As Document
   Set SaveSel = Selection.Range
   Set MainDoc = ActiveDocument
   LatexFilePath = Local_Document_Path(MainDoc) & "\" & Split(MainDoc.Name, ".")(0) & "-Latex"
   If Dir(LatexFilePath, vbDirectory) = "" Then MkDir LatexFilePath
   Selection.WholeStory 'Select whole document
   Selection.Expand wdParagraph 'Expands your selection to current paragraph
   Selection.Copy
   'd.Range.Copy
   OpretHiddendoc
   UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
   '    If HiddenDoc Is Nothing Then Set HiddenDoc = Documents.Add(, , , False)
   '    HiddenDoc.BuiltInDocumentProperties("Title") = "WordMatLatexHiddenDoc"
   DoEvents
   HiddenDoc.Activate
   HiddenDoc.Range.Select
   '    Selection.EndKey wdStory 'Move to end of document
   DoEvents
   Selection.PasteAndFormat wdPasteDefault  ' can cause error if d2 is not ready
   'Selection.Paste

   UfWait.Label_tip.Caption = "Converting equations"
   UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
   ConvertAllEquations False
    
   UfWait.Label_tip.Caption = "Converting formatting"
   UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
   ConvertFormattingToLatex HiddenDoc.Range
    
   '    ConvertRangeToLatex HiddenDoc.Range
    
   UfWait.Label_tip.Caption = "Converting images"
   UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
   ConvertImagesToLatex HiddenDoc
    
   Dim bm As Bookmark ' does not work
   For Each bm In HiddenDoc.Bookmarks
      bm.Range.InsertAfter "\ref{" & bm.Name & "}"
      bm.Delete
   Next
   
   Dim footn As Footnote
   For Each footn In HiddenDoc.Footnotes
      footn.Reference.InsertAfter "\footnote{" & footn.Range.text & "}"
      footn.Delete
   Next
   
   Dim toc As TableOfContents
   For Each toc In HiddenDoc.TablesOfContents
      toc.Range.InsertAfter "\tableofcontents" & vbCrLf
      toc.Delete ' det hele bliver ikke slettet
   Next
   
   UfWait.Label_tip.Caption = "Converting bibliography"
   ' Fields
   Dim F As Field, CiteName As String, CiteP As String, p2 As Integer
   For Each F In HiddenDoc.Fields
      If F.Type = wdFieldCitation Then
         HasBib = True
         CiteName = Split(Trim$(F.Code), " ")(1)
         F.Select
         p = InStr(F.Code, "\p")
         If p > 0 Then
            p2 = InStr(p + 3, F.Code, " ")
            CiteP = Mid$(F.Code, p + 3, p2 - p - 3)
         End If
         Selection.Collapse wdCollapseEnd
         F.Delete
         If CiteP = "" Then
            Selection.TypeText "\cite{" & CiteName & "}"
         Else
            Selection.TypeText "\cite[p.~" & CiteP & "]{" & CiteName & "}"
         End If
      ElseIf F.Type = wdFieldAuthor Then
         latexfil.Author = F.Code ' needs to be adjusted the name is in code
      ElseIf F.Type = wdFieldBibliography Then
         F.Select
         Selection.MoveStart wdLine, -2
         Selection.Delete
      End If
   Next
   
   UfWait.Label_tip.Caption = "Converting tables"
   UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
   Dim t As Table, r As Row, c As Cell
   For Each t In HiddenDoc.Tables
      s = ""
      If t.Rows.Alignment = wdAlignRowCenter Then s = s & "\begin{center}" & vbCrLf
      s = s & "\begin{tabular}{"
      If t.Columns(1).Borders.Item(wdBorderLeft).LineStyle <> wdLineStyleNone Then s = s & "|"
      For i = 1 To t.Columns.Count
         s = s & "c"
         If t.Columns(i).Borders.Item(wdBorderRight).LineStyle <> wdLineStyleNone Then
            s = s & "|"
         Else
            s = s & " "
         End If
      Next
      s = s & "}" & vbCrLf
      If t.Rows(1).Borders.Item(wdBorderTop).LineStyle <> wdLineStyleNone Then s = s & "\hline" & vbCrLf
      For Each r In t.Rows
         For Each c In r.Cells
            s = s & Left$(c.Range.text, Len(c.Range.text) - 2) & "&"
         Next
         s = Left$(s, Len(s) - 1) & "\\ "
         If r.Borders.Item(wdBorderBottom).LineStyle <> wdLineStyleNone Then s = s & "\hline"
         s = s & vbCrLf
      Next
      s = s & "\end{tabular}"
      If t.Rows.Alignment = wdAlignRowCenter Then s = s & vbCrLf & "\end{center}"
      '      t.Range.InsertAfter s
      t.Select
      Selection.Collapse wdCollapseEnd
      Selection.TypeText s
      t.Delete
   Next
    
   UfWait.Label_tip.Caption = "Konverterer sektioner, paragrafer, ..."
   UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
   HiddenDoc.Activate
    
   '    For Each l In HiddenDoc.Lists ' causes problems as a list can be split into two, with a normal paragraph in between. However, the two parts are still one list, so the paragraph in between is deleted. Replaced by another method
   '      ConvertList l
   '    Next
        
   For i = 1 To ActiveDocument.Paragraphs.Count - 1
      '        MsgBox ActiveDocument.Paragraphs(i).Range.Style & vbCrLf & ActiveDocument.Paragraphs(i).Range.text
      If ActiveDocument.Paragraphs(i).Range.OMaths.Count > 0 Then
        
      ElseIf ActiveDocument.Paragraphs(i).Range.Style = ActiveDocument.Styles(wdStyleTitle) Or ActiveDocument.Paragraphs(i).Range.Style = "Title" Then
         If LatexDocumentclass = 0 Then
            latexfil.Titel = Replace(ActiveDocument.Paragraphs(i).Range.text, vbCr, "")
         Else
            latexfil.InsertChapter ActiveDocument.Paragraphs(i).Range.text
         End If
      ElseIf ActiveDocument.Paragraphs(i).Range.Style = ActiveDocument.Styles(wdStyleHeading1) Or ActiveDocument.Paragraphs(i).Range.Style = "Heading 1" Or ActiveDocument.Paragraphs(i).Range.Style = "Overskrift 1" Then
         latexfil.InsertSection (ActiveDocument.Paragraphs(i).Range.text)
      ElseIf ActiveDocument.Paragraphs(i).Style = ActiveDocument.Styles(wdStyleHeading2) Or ActiveDocument.Paragraphs(i).Range.Style = "Heading 2" Then
         latexfil.InsertSubSection (ActiveDocument.Paragraphs(i).Range.text)
      ElseIf ActiveDocument.Paragraphs(i).Range.Style = ActiveDocument.Styles(wdStyleHeading3) Or ActiveDocument.Paragraphs(i).Range.Style = "Heading 3" Then
         latexfil.InsertSubSubSection (ActiveDocument.Paragraphs(i).Range.text)
      ElseIf ActiveDocument.Paragraphs(i).Range.Style = ActiveDocument.Styles(wdStyleNormal) And InStr(ActiveDocument.Paragraphs(i).Range.text, "\") <= 0 Then
         latexfil.InsertParagraph (ActiveDocument.Paragraphs(i).Range.text)
      ElseIf HiddenDoc.Paragraphs(i).Range.ListFormat.ListType = wdListBullet Or HiddenDoc.Paragraphs(i).Range.ListFormat.ListType = wdListMixedNumbering Or HiddenDoc.Paragraphs(i).Range.ListFormat.ListType = wdListOutlineNumbering Or HiddenDoc.Paragraphs(i).Range.ListFormat.ListType = wdListPictureBullet Then  ' Or HiddenDoc.Paragraphs(i).Range.ListFormat.ListType = wdListNoNumbering
         If HiddenDoc.Paragraphs(i - 1).Range.ListFormat.ListType <> HiddenDoc.Paragraphs(i).Range.ListFormat.ListType Then
            latexfil.InsertText "\begin{itemize}" & vbCrLf
            latexfil.InsertText " \item " & HiddenDoc.Paragraphs(i).Range.text
         ElseIf HiddenDoc.Paragraphs(i).Range.ListFormat.ListLevelNumber > HiddenDoc.Paragraphs(i - 1).Range.ListFormat.ListLevelNumber Then
            latexfil.InsertText " \begin{itemize}" & vbCrLf & " \item " & HiddenDoc.Paragraphs(i).Range.text
         ElseIf HiddenDoc.Paragraphs(i).Range.ListFormat.ListLevelNumber < HiddenDoc.Paragraphs(i - 1).Range.ListFormat.ListLevelNumber Then
            latexfil.InsertText " \end{itemize}" & vbCrLf
            latexfil.InsertText " \item " & HiddenDoc.Paragraphs(i).Range.text
         Else
            latexfil.InsertText " \item " & HiddenDoc.Paragraphs(i).Range.text
         End If
         If HiddenDoc.Paragraphs(i + 1).Range.ListFormat.ListType <> HiddenDoc.Paragraphs(i).Range.ListFormat.ListType Then
            latexfil.InsertText "\end{itemize}"
         End If
      ElseIf HiddenDoc.Paragraphs(i).Range.ListFormat.ListType = wdListSimpleNumbering Or HiddenDoc.Paragraphs(i).Range.ListFormat.ListType = wdListListNumOnly Then
         If HiddenDoc.Paragraphs(i - 1).Range.ListFormat.ListType <> HiddenDoc.Paragraphs(i).Range.ListFormat.ListType Then
            latexfil.InsertText "\begin{enumerate}" & vbCrLf
            latexfil.InsertText " \item " & HiddenDoc.Paragraphs(i).Range.text
         ElseIf HiddenDoc.Paragraphs(i).Range.ListFormat.ListLevelNumber > HiddenDoc.Paragraphs(i - 1).Range.ListFormat.ListLevelNumber Then
            latexfil.InsertText " \begin{enumerate}" & vbCrLf & " \item " & HiddenDoc.Paragraphs(i).Range.text
         ElseIf HiddenDoc.Paragraphs(i).Range.ListFormat.ListLevelNumber < HiddenDoc.Paragraphs(i - 1).Range.ListFormat.ListLevelNumber Then
            latexfil.InsertText " \end{enumerate}" & vbCrLf
            latexfil.InsertText " \item " & HiddenDoc.Paragraphs(i).Range.text
         Else
            latexfil.InsertText " \item " & HiddenDoc.Paragraphs(i).Range.text
         End If
         If HiddenDoc.Paragraphs(i + 1).Range.ListFormat.ListType <> HiddenDoc.Paragraphs(i).Range.ListFormat.ListType Then
            latexfil.InsertText "\end{enumerate}"
         End If
      Else
         '            MsgBox ActiveDocument.Paragraphs(i).Range.Style
         latexfil.InsertText (ActiveDocument.Paragraphs(i).Range.text)
      End If
   Next
    
   UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
   If HasBib Then
      Dim src As Source, srcI As Integer, sXML As String, SrcTitle As String, SrcAuthor As String, SrcYear As String, SrcPubl As String, SrcEdition As String, nl As String, pn As String, FN As String, ln As String
      If HiddenDoc.Bibliography.Sources.Count > 0 Then
         srcI = 1
         latexfil.InsertText "\newpage" & vbCrLf
         latexfil.InsertText "\begin{thebibliography}{" & HiddenDoc.Bibliography.Sources.Count & "}" & vbCrLf
         For Each src In HiddenDoc.Bibliography.Sources
            latexfil.InsertText "\bibitem{" & src.Tag & "}" & vbCrLf
            sXML = src.XML
            SrcTitle = ExtractTag(sXML, "<b:Title>", "</b:Title>")
            SrcYear = ExtractTag(sXML, "<b:Year>", "</b:Year>")
            SrcPubl = ExtractTag(sXML, "<b:Publisher>", "</b:Publisher>")
            SrcEdition = ExtractTag(sXML, "<b:Edition>", "</b:Edition>")
            nl = ExtractTag(sXML, "<b:NameList>", "</b:NameList>")
            Do
               pn = ExtractTag(nl, "<b:Person>", "</b:Person>")
               FN = ExtractTag(pn, "<b:First>", "</b:First>")
               ln = ExtractTag(pn, "<b:Last>", "</b:Last>")
               If FN <> "" Or ln <> "" Then
                  SrcAuthor = SrcAuthor & FN & " " & ln & ", "
                  p = InStr(nl, "</b:Person>")
                  nl = Right$(nl, Len(nl) - p - 10)
               End If
            Loop While FN <> "" And ln <> "" And nl <> ""
            
            '         MsgBox sXML
            If SrcAuthor <> "" Then latexfil.InsertText "  " & SrcAuthor & vbCrLf
            If SrcTitle <> "" Then latexfil.InsertText "  \textit{" & SrcTitle & "}," & vbCrLf
            If SrcPubl <> "" Then latexfil.InsertText "  " & SrcPubl & "," & vbCrLf
            If SrcEdition <> "" Then latexfil.InsertText "  " & SrcEdition & "," & vbCrLf
            If SrcYear <> "" Then latexfil.InsertText "  " & SrcYear & "." & vbCrLf
            latexfil.InsertText vbCrLf
            srcI = srcI + 1
         Next
         latexfil.InsertText "\end{thebibliography}" & vbCrLf
      End If
   End If
   
    latexfil.ValidateLatexDoc
    
   UfWait.Label_tip.Caption = "Gemmer fil"
   latexfil.CreateHeader
   UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
   If doctype = 0 Then
      latexfil.SavePdf LatexFilePath, Split(MainDoc.Name, ".")(0) 'Environ("TEMP")
   ElseIf doctype = 1 Then
      latexfil.Savedvi LatexFilePath, Split(MainDoc.Name, ".")(0) 'Environ("TEMP")
   ElseIf doctype = 2 Then
      latexfil.SaveTex LatexFilePath, Split(MainDoc.Name, ".")(0) & ".tex" 'Environ("TEMP")
#If Mac Then
#Else
      RunDefaultProgram Split(MainDoc.Name, ".")(0) & ".tex", LatexFilePath 'Environ("TEMP")
#End If
   End If
    
   UserFormLatex.Label_status.visible = False
   '    d2.Close False
   '    MainDoc.Activate
   SaveSel.Select
slut:
   Unload UfWait
End Sub
Sub ConvertList(l As List)
   Dim s As String, i As Long
   If l.Range.ListFormat.ListType = wdListBullet Or l.Range.ListFormat.ListType = wdListMixedNumbering Or l.Range.ListFormat.ListType = wdListNoNumbering Or l.Range.ListFormat.ListType = wdListOutlineNumbering Or l.Range.ListFormat.ListType = wdListPictureBullet Then
      s = "\begin{itemize}" & vbCrLf
      For i = 1 To l.ListParagraphs.Count
         If i > 1 Then
            If l.ListParagraphs(i).Range.ListFormat.ListLevelNumber > l.ListParagraphs(i - 1).Range.ListFormat.ListLevelNumber Then
               s = s & " \begin{itemize}" & vbCrLf
            ElseIf l.ListParagraphs(i).Range.ListFormat.ListLevelNumber < l.ListParagraphs(i - 1).Range.ListFormat.ListLevelNumber Then
               s = s & " \end{itemize}" & vbCrLf
            End If
         End If
         s = s & " \item " & l.ListParagraphs(i).Range.text
      Next
      s = s & "\end{itemize}" & vbCrLf
      l.Range.InsertAfter s
      l.Range.Delete
   ElseIf l.Range.ListFormat.ListType = wdListSimpleNumbering Or l.Range.ListFormat.ListType = wdListListNumOnly Then
      s = "\begin{enumerate}" & vbCrLf
      For i = 1 To l.ListParagraphs.Count
         If i > 1 Then
            If l.ListParagraphs(i).Range.ListFormat.ListLevelNumber > l.ListParagraphs(i - 1).Range.ListFormat.ListLevelNumber Then
               s = s & " \begin{enumerate}" & vbCrLf
            ElseIf l.ListParagraphs(i).Range.ListFormat.ListLevelNumber < l.ListParagraphs(i - 1).Range.ListFormat.ListLevelNumber Then
               s = s & " \end{enumerate}" & vbCrLf
            End If
         End If
         s = s & " \item " & l.ListParagraphs(i).Range.text
      Next
      s = s & "\end{enumerate}" & vbCrLf
      l.Range.InsertAfter s
      l.Range.Delete
   End If

End Sub
Function Get3DigitImageNo(n As Integer) As String
   If n < 10 Then
      Get3DigitImageNo = "00" & n
   ElseIf n < 100 Then
      Get3DigitImageNo = "0" & n
   Else
      Get3DigitImageNo = n
   End If
End Function

Sub ConvertImagesToLatex(D As Document)
'   On Error GoTo Fejl
   Dim FilNavn As String, ImagFilDir As String, tDoc As Document, sh As InlineShape, si As Integer, sh2 As Shape, sha As Variant
   Dim ImagCol As New Collection
   If D.InlineShapes.Count = 0 And D.Shapes.Count = 0 Then
      latexfil.ImagDir = ""
      Exit Sub
   End If
   D.Range.Copy
   Set tDoc = Documents.Add(, , , False)
'   Wait (1)
   DoEvents
   
   
   tDoc.Activate
   tDoc.Range.Select
   DoEvents
   Selection.PasteAndFormat wdPasteDefault  ' kan fejle hvis d2 ikke er klar
'   Selection.PasteAndFormat (wdSingleCellText)
   '   Selection.Paste
    
   FilNavn = LatexFilePath & "\" & "Images.htm"
   If Dir(FilNavn) <> "" Then Kill FilNavn
   ImagFilDir = LatexFilePath & "\" & Dir(LatexFilePath & "\Images*", vbDirectory)
   If Len(Dir(ImagFilDir, vbDirectory)) > 1 Then
      Kill ImagFilDir & "\*.*"
      RmDir ImagFilDir
   End If
   tDoc.SaveAs2 fileName:=FilNavn, FileFormat:=wdFormatFilteredHTML, LockComments:=False, Password:="", AddToRecentFiles:=False, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:=False, CompatibilityMode:=0
'   ImagFilDir = LatexFilePath & "\" & Dir(LatexFilePath & "\Images-*", vbDirectory)
   latexfil.ImagDir = Dir(LatexFilePath & "\Images-*", vbDirectory)
   tDoc.Close
   Set tDoc = Nothing

   If Dir(FilNavn, vbNormal) <> "" Then
      Kill FilNavn
   End If

   For Each sh In D.InlineShapes
      ImagCol.Add sh
   Next
   For Each sh2 In D.Shapes
      ImagCol.Add sh2
   Next
   
   SortImagCol ImagCol, 1, ImagCol.Count
   
   si = 1
   For Each sha In ImagCol
      If TypeOf sha Is InlineShape Then
         If sha.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter Then
            sha.Range.InsertAfter ("\begin{center}\includegraphics{image" & Get3DigitImageNo(si) & "}\end{center}")
         Else
            sha.Range.InsertAfter ("\includegraphics{image" & Get3DigitImageNo(si) & "}")
         End If
      Else
         If sha.Anchor.ParagraphFormat.Alignment = wdAlignParagraphCenter Then
            sha.Anchor.InsertAfter ("\begin{center}\includegraphics{image" & Get3DigitImageNo(si) & "}\end{center}")
         Else
            sha.Anchor.InsertAfter ("\includegraphics{image" & Get3DigitImageNo(si) & "}")
         End If
      End If
      si = si + 1
   Next
       
   For Each sh In D.InlineShapes
      sh.Delete
   Next
   For Each sh2 In D.Shapes
      sh2.Delete
   Next
   
   GoTo slut
fejl:
   MsgBox "Fejl " & Err.Number & " (" & Err.Description & ") i procedure ConvertImagesToLatex, linje " & Erl & ".", vbOKOnly Or vbCritical Or vbSystemModal, "Fejl"
slut:
End Sub

Sub SortImagCol(coll As Collection, first As Long, last As Long)
' There is no built-in function to sort a collection.
' QuickSort(coll,1,coll.Count)

  Dim vCentreVal As Variant, vTemp As Variant
  
  Dim lTempLow As Long
  Dim lTempHi As Long
  lTempLow = first
  lTempHi = last
  
  Set vCentreVal = coll.Item((first + last) \ 2)
  Do While lTempLow <= lTempHi
  
    Do While GetShapePos(coll(lTempLow)) < GetShapePos(vCentreVal) And lTempLow < last
      lTempLow = lTempLow + 1
    Loop
    
    Do While GetShapePos(vCentreVal) < GetShapePos(coll(lTempHi)) And lTempHi > first
      lTempHi = lTempHi - 1
    Loop
    
    If lTempLow <= lTempHi Then
    
      ' Swap values
      Set vTemp = coll(lTempLow)
      
      coll.Add coll(lTempHi), After:=lTempLow
      coll.Remove lTempLow
      
      coll.Add vTemp, Before:=lTempHi
      coll.Remove lTempHi + 1
      
      ' Move to next positions
      lTempLow = lTempLow + 1
      lTempHi = lTempHi - 1
      
    End If
    
  Loop
  
  If first < lTempHi Then SortImagCol coll, first, lTempHi
  If lTempLow < last Then SortImagCol coll, lTempLow, last
End Sub
Function GetShapePos(sh As Variant) As Long
' sh can be shape or inlineshape
   If TypeOf sh Is Shape Then
      GetShapePos = sh.Anchor.start
   ElseIf TypeOf sh Is InlineShape Then
      GetShapePos = sh.Range.start
   End If
End Function
Sub ConvertFormattingToLatex(r As Range)
    With r.Find
        .text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .ClearFormatting
        .Font.Bold = True
        With .Replacement
            .ClearFormatting
            .text = "\textbf{^&}"
            .Font.Bold = False
        End With
        .Execute Replace:=wdReplaceAll
        .ClearFormatting
        .Font.Italic = True
        With .Replacement
            .ClearFormatting
            .text = "\textit{^&}"
            .Font.Italic = False
        End With
        .Execute Replace:=wdReplaceAll
        .ClearFormatting
        .Font.Underline = True
        With .Replacement
            .ClearFormatting
            .text = "\underline{^&}"
            .Font.Underline = False
        End With
        .Execute Replace:=wdReplaceAll
        .ClearFormatting
        .text = "^m"
        With .Replacement
            .ClearFormatting
            .text = "^&\newpage"
        End With
        .Execute Replace:=wdReplaceAll
        .ClearFormatting
        .text = "^l" ' softbreak
        With .Replacement
            .ClearFormatting
            .text = "\newline^10"
        End With
        .Execute Replace:=wdReplaceAll
    End With

End Sub

Function ContainsFormatting(r As Range) As Boolean
    If Selection.Range.ParagraphStyle Is Nothing Then
        ContainsFormatting = True
    ElseIf r.Tables.Count = 0 And r.Bold = 0 And r.Italic = 0 And r.Underline = 0 And r.Hyperlinks.Count = 0 And r.InlineShapes.Count = 0 And r.ShapeRange.Count = 0 And r.Font.Name <> vbNullString And r.Font.Size > 1000 And r.Font.Superscript = 0 And r.Font.Subscript = 0 And r.ListParagraphs.Count = 0 Then
        ContainsFormatting = False
    Else
        ContainsFormatting = True
    End If

End Function

Sub OpretHiddendoc()
' but only if does not already exist
#If Mac Then
'    Call tempDoc
#Else
Dim D As Document
If HiddenDoc Is Nothing Then
For Each D In Application.Documents
    If D.BuiltInDocumentProperties("Title") = "WordMatLatexHiddenDoc" Then
        Set HiddenDoc = D
        Exit For
    End If
Next
End If

If HiddenDoc Is Nothing Then
    Set HiddenDoc = Documents.Add(, , , False)
    HiddenDoc.BuiltInDocumentProperties("Title") = "WordMatLatexHiddenDoc"
End If

If Not HiddenDoc.BuiltInDocumentProperties("Title") = "WordMatLatexHiddenDoc" Then
    HiddenDoc.Close SaveChanges:=wdDoNotSaveChanges
    Set HiddenDoc = Documents.Add(, , , False)
    HiddenDoc.BuiltInDocumentProperties("Title") = "WordMatLatexHiddenDoc"
End If
#End If
End Sub

Public Sub ConvertAllEquations(Optional KeepOriginal As Boolean = False)
   Dim mi As Integer, i As Integer
   Dim antal As Integer
   Application.ScreenUpdating = False
   UserFormLatex.SaveSet
   LAlign = 0
   antal = ActiveDocument.OMaths.Count

   i = 1
   For mi = 1 To antal
      If KeepOriginal Then i = mi
      If mi = antal Then
         If LAlign > 0 Then
            LAlign = 3
         End If
      ElseIf MainDoc.OMaths(mi).AlignPoint > 0 And LAlign = 0 Then
         LAlign = 1 ' start on align
      ElseIf MainDoc.OMaths(mi).AlignPoint > 0 And MainDoc.OMaths(mi + 1).AlignPoint > 0 Then
         LAlign = 2 ' continue
      ElseIf MainDoc.OMaths(mi).AlignPoint > 0 And MainDoc.OMaths(mi + 1).AlignPoint < 0 Then
         LAlign = 3 ' stop
      Else
         LAlign = 0
      End If
'      MainDoc.OMaths(mi).Range.Select ' destroys alignment
      HiddenDoc.OMaths(i).Range.Select
      omax.ReadSelection
      HiddenDoc.OMaths(i).Range.Select
      ConvertEquationToLatex
      '        UpDateLatex
   Next
   Application.ScreenUpdating = True

End Sub
Sub TestEQ()
   MsgBox Selection.OMaths(1).AlignPoint & vbCrLf & Selection.OMaths(1).Justification & vbCrLf & Selection.OMaths(1).Type & vbCrLf
   ' justification 1=centreret, 2=gruppe
      omax.ReadSelection
     ConvertEquationToLatex
End Sub
Sub ConvertEquationToLatex(Optional KeepOriginal As Boolean = False)
    ' for miktex
    Dim t As Table, s As String, p As Integer, EqStart As String, EqEnd As String, eq As OMath
    If Not UserFormLatex.EventsOn Then Exit Sub
    If Selection.OMaths.Count = 0 Then Exit Sub
    Set eq = Selection.OMaths(1)
    UserFormLatex.Label_input.Caption = omax.Kommando
    LatexCode = omax.ConvertToLatex(omax.Kommando)
    '   If OptionButton_visstor.Value = True Then
    '      LatexCode = "\displaystyle " & LatexCode
    '   ElseIf OptionButton_visinline.Value = True Then
    '      LatexCode = "\inline " & LatexCode
    '   End If
   
    If Selection.OMaths(1).Breaks.Count > 0 Then
        p = Selection.OMaths(1).Breaks(1).Range.start
    End If
    If LAlign > 0 Then
        If LAlign = 1 Then
            EqStart = "\begin{align}"
            EqEnd = "\\"
            LatexCode = Replace(LatexCode, "=", "&=", 1, 1)
        ElseIf LAlign = 2 Then
            EqStart = " "
            EqEnd = "\\"
            LatexCode = Replace(LatexCode, "=", "&=", 1, 1)
        Else
            EqStart = " "
            EqEnd = "\end{align}"
            LatexCode = Replace(LatexCode, "=", "&=", 1, 1)
        End If
    Else
        EqStart = "\begin{equation}"
        EqEnd = "\end{equation}"
    End If

    If UserFormLatex.OptionButton_omslutauto.Value = True Then
        If Selection.OMaths(1).Justification = wdOMathJcInline Then
            s = "$" & LatexCode & "$"
        Else
            If Selection.OMaths(1).Range.Tables.Count > 0 Then
                Set t = Selection.OMaths(1).Range.Tables(1)
                If t.Rows.Count = 1 And t.Columns.Count = 3 Then
                    If t.Cell(1, 2).Range.OMaths.Count > 0 And t.Cell(1, 3).Range.Fields.Count > 0 Then
                        If t.Range.Bookmarks.Count > 0 Then
                            s = EqStart & "\label{" & t.Range.Bookmarks(1).Name & "}" & LatexCode & EqEnd
                        Else
                            s = EqStart & LatexCode & EqEnd
                        End If
                        t.Delete
                    Else
                        s = "$" & LatexCode & "$"
                    End If
                Else
                    s = "$" & LatexCode & "$"
                End If
            Else
                If InStr(EqStart, "\begin") > 0 Then EqStart = Left$(EqStart, Len(EqStart) - 1) & "*" & Right$(EqStart, 1)
                If InStr(EqEnd, "\end") > 0 Then EqEnd = Left$(EqEnd, Len(EqEnd) - 1) & "*" & Right$(EqEnd, 1)
                s = EqStart & LatexCode & EqEnd
            End If
        End If
    Else
        s = LatexStart & LatexCode & LatexSlut
    End If
   
    If Not (KeepOriginal) Then
        '            ActiveDocument.OMaths(i).Range.text = ""
        '      On Error Resume Next
        '      Selection.OMaths(1).Range.Delete
        eq.Range.Delete
    Else
        omax.GoToEndOfSelectedMaths
        Selection.TypeParagraph
    End If
    If Selection.OMaths.Count > 0 Then
        Selection.MoveLeft
        If Selection.OMaths.Count > 0 Then Selection.MoveLeft
        Selection.TypeParagraph
    End If
    Selection.InsertAfter s
End Sub

Sub KonverterTilLaTex()
    PrepareMaxima
    Dim uflatex As New UserFormLatex
    uflatex.Show
End Sub
Sub ToggleLatex()
Dim mtext As String
Dim r As Range
On Error GoTo slut
    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord
    If Selection.OMaths.Count > 0 Then
        PrepareMaxima
        omax.ReadSelection
        Selection.OMaths(1).Range.text = ""
        Selection.InsertAfter LatexStart & omax.ConvertToLatex(omax.Kommando) & LatexSlut
    Else
        PrepareMaxima
        
        mtext = omax.ConvertLatexToWord(RemoveLatexOmslut(Selection.Range.text))
        Selection.Range.Delete
        Selection.Collapse wdCollapseEnd
        Set r = Selection.OMaths.Add(Selection.Range)
        Selection.TypeText mtext
        r.OMaths(1).BuildUp
        Selection.TypeParagraph
    End If
    Oundo.EndCustomRecord

slut:
End Sub
Function RemoveLatexOmslut(text As String)

    text = TrimB(text, "$")
    text = TrimL(text, "\[")
    text = TrimR(text, "\]")
    RemoveLatexOmslut = text
End Function

Function ReadTextfileToString(FilNavn As String) As String
#If Mac Then
   Dim filnr As Integer
   filnr = FreeFile()
   Open FilNavn For Input As filnr   ' Open file
   ReadTextfileToString = Input$(LOF(1), 1)
   Close #filnr
#Else
   Dim fsT As Object
   'On Error GoTo fejl

   Set fsT = CreateObject("ADODB.Stream")
   fsT.Type = 2 'Specify stream type - we want To save text/string data.
   fsT.Charset = "iso-8859-1" 'Specify charset For the source text data. (Alternate: utf-8)
   fsT.Open 'Open the stream
   fsT.LoadFromFile FilNavn
   ReadTextfileToString = fsT.ReadText()
   fsT.Close
   Set fsT = Nothing
#End If

   GoTo slut
fejl:
   MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:

End Function

Sub WriteTextfileToString(FilNavn As String, WriteText As String)
#If Mac Then
   Dim filnr As Integer
   filnr = FreeFile()
   Open FilNavn For Output As filnr   ' Open file for output.
   
   Print #filnr, WriteText  ' print skriver uden " "
   Close #filnr    ' Close file.
#Else
   Dim fsT As Object
   'On Error GoTo fejl

   If FilNavn = vbNullString Then GoTo slut
   If WriteText = vbNullString Then
      If Dir(FilNavn) <> vbNullString Then Kill FilNavn
         GoTo slut
   End If
   Set fsT = CreateObject("ADODB.Stream")
   fsT.Type = 2 'Specify stream type - we want To save text/string data.
   fsT.Charset = "iso-8859-1" 'Specify charset For the source text data. utf-8
   fsT.Open 'Open the stream And write binary data To the object
   fsT.WriteText WriteText
   On Error GoTo 0
   fsT.SaveToFile FilNavn, 2 'Save binary data To disk
   fsT.Close
   Set fsT = Nothing
#End If

   GoTo slut
fejl:
   MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:

End Sub
