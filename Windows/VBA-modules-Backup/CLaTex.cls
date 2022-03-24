VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLaTex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private LaTexDoc As String
Public Titel As String
Public toc As Boolean
Public TitlePage As Boolean
Public Author As String
Public ImagDir As String
Public Documentclass As String
Public Fontsize As String
Public UseWordMargins As Boolean
Private Log As String

Public ufwait As UserFormWaitForMaxima

Private MikTex64bit As Boolean

Private Sub Class_Initialize()
    toc = True
    TitlePage = True
    Titel = Split(ActiveDocument.Name, ".")(0)
    Author = ActiveDocument.BuiltInDocumentProperties(wdPropertyAuthor)
    UseWordMargins = LatexWordMargins
End Sub

Public Sub Reset()
'    toc = True
'    TitlePage = True
   Titel = Split(ActiveDocument.Name, ".")(0)
   Author = ActiveDocument.BuiltInDocumentProperties(wdPropertyAuthor)
   LaTexDoc = ""

End Sub

Public Sub CreateHeader()

If Titel = "" Then
    Titel = Split(ActiveDocument.Name, ".")(0)
End If

LaTexDoc = FixedLatexPreamble1 & LatexPreamble & FixedLatexPreamble2 & LaTexDoc
End Sub

Public Function FixedLatexPreamble1() As String
   Dim s As String, pageNo As String, LID As Long, dClass As String, ps As String
   
   s = s & "% !TEX encoding = ISO-8859-1" & vbCrLf ' kun til TexMaker?
   If LatexDocumentclass = 0 Then
      dClass = "article"
   ElseIf LatexDocumentclass = 1 Then
      dClass = "book"
   ElseIf LatexDocumentclass = 2 Then
      dClass = "report"
   Else
      dClass = "article"
   End If
   
   If LatexFontsize = "" Then LatexFontsize = "11"
   
   s = s & "\documentclass[" & LatexFontsize & "pt]{" & dClass & "}" & vbCrLf
   s = s & "\usepackage[T1]{fontenc}" & vbCrLf
   s = s & "\usepackage[latin1]{inputenc}" & vbCrLf
   s = s & "\usepackage{geometry}" & vbCrLf
   If ActiveDocument.PageSetup.PaperSize = wdPaperA4 Then
      ps = "a4paper"
   ElseIf ActiveDocument.PageSetup.PaperSize = wdPaperLetter Then
      ps = "letterpaper"
   ElseIf ActiveDocument.PageSetup.PaperSize = wdPaperA5 Then
      ps = "a5paper"
   ElseIf ActiveDocument.PageSetup.PaperSize = wdPaperA3 Then
      ps = "a3paper"
   ElseIf ActiveDocument.PageSetup.PaperSize = wdPaperB4 Then
      ps = "b4paper"
   ElseIf ActiveDocument.PageSetup.PaperSize = wdPaperB5 Then
      ps = "b5paper"
   ElseIf ActiveDocument.PageSetup.PaperSize = wdPaperLegal Then
      ps = "legalpaper"
   End If
   If ActiveDocument.PageSetup.Orientation = wdOrientLandscape Then
      If ps <> "" Then
         ps = ps & ", landscape"
      Else
         ps = ps & "landscape"
      End If
   End If
   If UseWordMargins Then
      s = s & " \geometry{" & ps & ","
      s = s & "left=" & Round(PointsToCentimeters(ActiveDocument.PageSetup.LeftMargin), 1) * 10 & "mm,"
      s = s & "right=" & Round(PointsToCentimeters(ActiveDocument.PageSetup.RightMargin), 1) * 10 & "mm,"
      s = s & "top=" & Round(PointsToCentimeters(ActiveDocument.PageSetup.TopMargin), 1) * 10 & "mm,"
      s = s & "bottom=" & Round(PointsToCentimeters(ActiveDocument.PageSetup.BottomMargin), 1) * 10 & "mm"
      s = s & "}" & vbCrLf
   ElseIf ps <> "" Then
      s = s & "\geometry{" & ps & "}" & vbCrLf
   End If
   
   LID = ActiveDocument.Paragraphs(1).Range.LanguageID

   If LID = wdEnglishUK Then
      s = s & "\usepackage[UKenglish]{babel}" & vbCrLf
   ElseIf LID = wdDanish Then
      s = s & "\usepackage[danish]{babel}" & vbCrLf
   ElseIf LID = wdEnglishAUS Then
      s = s & "\usepackage[australian]{babel}" & vbCrLf
   ElseIf LID = wdEnglishUS Then
      s = s & "\usepackage[USenglish]{babel}" & vbCrLf
   ElseIf LID = wdEnglishCanadian Then
      s = s & "\usepackage[canadian]{babel}" & vbCrLf
   ElseIf LID = wdEnglishNewZealand Then
      s = s & "\usepackage[newzealand]{babel}" & vbCrLf
   ElseIf LID = wdEnglishIreland Then
      s = s & "\usepackage[irish]{babel}" & vbCrLf
   ElseIf LID = wdSpanish Then
      s = s & "\usepackage[spanish]{babel}" & vbCrLf
   ElseIf LID = wdCatalan Then
      s = s & "\usepackage[catalan]{babel}" & vbCrLf
   ElseIf LID = wdGerman Then
      s = s & "\usepackage[german]{babel}" & vbCrLf
   ElseIf LID = wdGermanAustria Then
      s = s & "\usepackage[austrian]{babel}" & vbCrLf
   ElseIf LID = wdFrench Then
      s = s & "\usepackage[french]{babel}" & vbCrLf
   ElseIf LID = wdGreek Then
      s = s & "\usepackage[greek]{babel}" & vbCrLf
   ElseIf LID = wdRussian Then
      s = s & "\usepackage[russian]{babel}" & vbCrLf
   ElseIf LID = wdUkrainian Then
      s = s & "\usepackage[ukrainian]{babel}" & vbCrLf
   ElseIf LID = wdBulgarian Then
      s = s & "\usepackage[bulgarian]{babel}" & vbCrLf
   ElseIf LID = wdHungarian Then
      s = s & "\usepackage[hungarian]{babel}" & vbCrLf
   ElseIf LID = wdBelgianDutch Then
      s = s & "\usepackage[belgian]{babel}" & vbCrLf
   ElseIf LID = wdBelgianFrench Then
      s = s & "\usepackage[french]{babel}" & vbCrLf
   ElseIf LID = wdItalian Then
      s = s & "\usepackage[italian]{babel}" & vbCrLf
   ElseIf LID = wdFinnish Then
      s = s & "\usepackage[finnish]{babel}" & vbCrLf
   ElseIf LID = wdSwedish Then
      s = s & "\usepackage[swedish]{babel}" & vbCrLf
   ElseIf LID = wdNorwegianBokmol Then
      s = s & "\usepackage[norsk]{babel}" & vbCrLf
   ElseIf LID = wdNorwegianNynorsk Then
      s = s & "\usepackage[nynorsk]{babel}" & vbCrLf
   ElseIf LID = wdRomanian Then
      s = s & "\usepackage[romanian]{babel}" & vbCrLf
   ElseIf LID = wdGaelicScotland Then
      s = s & "\usepackage[scottish]{babel}" & vbCrLf
   ElseIf LID = wdTurkish Then
      s = s & "\usepackage[turkish]{babel}" & vbCrLf
   ElseIf LID = wdDutch Then
      s = s & "\usepackage[dutch]{babel}" & vbCrLf
   ElseIf LID = wdCzech Then
      s = s & "\usepackage[czech]{babel}" & vbCrLf
   ElseIf LID = wdPortuguese Then
      s = s & "\usepackage[portuguese]{babel}" & vbCrLf
   ElseIf LID = wdPortugueseBrazil Then
      s = s & "\usepackage[brazil]{babel}" & vbCrLf
   ElseIf LID = wdWelsh Then
      s = s & "\usepackage[welsh]{babel}" & vbCrLf
   ElseIf LID = wdPolish Then
      s = s & "\usepackage[polish]{babel}" & vbCrLf
   ElseIf LID = wdSlovak Then
      s = s & "\usepackage[slovak]{babel}" & vbCrLf
   ElseIf LID = wdSerbianLatin Then
      s = s & "\usepackage[serbian]{babel}" & vbCrLf
   ElseIf LID = wdIndonesian Then
      s = s & "\usepackage[indonesian]{babel}" & vbCrLf
   ElseIf LID = wdIcelandic Then
      s = s & "\usepackage[icelandic]{babel}" & vbCrLf
   ElseIf LID = wdEstonian Then
      s = s & "\usepackage[estonian]{babel}" & vbCrLf
   ElseIf LID = wdCroatian Then
      s = s & "\usepackage[croatian]{babel}" & vbCrLf
   ElseIf LID = wdAlbanian Then
      s = s & "\usepackage[albanian]{babel}" & vbCrLf
   Else
      s = s & "\usepackage[english]{babel}" & vbCrLf
   End If
   
   s = s & "\usepackage{amsmath}" & vbCrLf
   
   If ImagDir <> "" Then
      s = s & "\usepackage{graphicx}" & vbCrLf
      s = s & "\graphicspath{{./" & ImagDir & "/}}" & vbCrLf
   End If
   
   s = s & "" & vbCrLf
   FixedLatexPreamble1 = s

End Function

Public Function FixedLatexPreamble2() As String
   Dim s As String, pageNo As String
   If Titel = "" Then
      Titel = Split(ActiveDocument.Name, ".")(0)
   End If
   
   If ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.Count = 0 And ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).PageNumbers.Count = 0 Then
      pageNo = 0 ' no pagenumbers
   Else
   End If

   s = s & "\title{" & Titel & "}" & vbCrLf
   s = s & "\author{" & Author & "}" & vbCrLf

   s = s & "" & vbCrLf
   s = s & "\begin{document}" & vbCrLf

   If pageNo = 0 Then
      s = s & "\pagestyle{empty}" & vbCrLf
      s = s & "\thispagestyle{empty}" & vbCrLf
   End If
   If TitlePage Then
      s = s & "\maketitle" & vbCrLf
      If pageNo = 0 Then s = s & "\thispagestyle{empty}" & vbCrLf
      s = s & "\newpage" & vbCrLf
   End If
   If toc Then
      s = s & "\tableofcontents" & vbCrLf
      s = s & "\newpage" & vbCrLf
   End If
   's = s & "" & vbCrLf
   s = s & "" & vbCrLf
   FixedLatexPreamble2 = s
End Function

Public Sub SaveTex(ByVal Mappe As String, ByVal filnavn As String)
Dim filnr As Integer, filren As String, s As String
Dim fsT As Object
'On Error GoTo fejl

   ufwait.Label_tip.Caption = "Gemmer Tex-fil"
LaTexDoc = LaTexDoc & "\end{document}" & vbCrLf
filren = Split(filnavn, ".")(0)
On Error Resume Next
s = Mappe & "\" & filren & ".log"
If Dir(s) <> "" Then Kill Mappe & "\" & filren & ".*"
's = Mappe & "\" & filren & ".toc"
'If Dir(s) <> "" Then
'Kill Mappe & "\*.*"
On Error GoTo 0

WriteTextfileToString Mappe & "\" & filnavn, LaTexDoc

'Set fsT = CreateObject("ADODB.Stream")
'fsT.Type = 2 'Specify stream type - we want To save text/string data.
'fsT.Charset = "iso-8859-1" 'Specify charset For the source text data. utf-8
'fsT.Open 'Open the stream And write binary data To the object
'fsT.WriteText LaTexDoc
'fsT.SaveToFile Mappe & "\" & filnavn, 2 'Save binary data To disk
'fsT.Close
'Set fsT = Nothing

'filnr = FreeFile()
'Open Mappe & "\" & filnavn For Output As filnr   ' Open file for output.
'
'Print #filnr, LaTexDoc  ' print skriver uden " "
'Close #filnr    ' Close file.


GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error '"Der skete en fejl i forsøget på at gemme LaTex-filen"
slut:

End Sub
Sub Savedvi(ByVal Mappe As String, ByVal filnavn As String)
   ' skal ikke bruges. pdf er bedre. fx includegraphics skal have flere parametre med denne
   Dim shellcmd As String
   Dim appnr As Integer
   Dim i As Integer, j As Integer
   Dim mtpath As String
#If Mac Then
   MsgBox Sprog.A(410), vbOKOnly, Sprog.Error
#Else
   Dim x64 As String
   MikTexPath ' called to set miktex64bit
   If MikTex64bit Then
      x64 = "x64\"
   End If
   mtpath = MikTexPath
   If mtpath = "" Then
      MsgBox Sprog.A(409)
      OpenLink "http://miktex.org"
      GoTo slut
   End If

   SaveTex Mappe, filnavn & ".tex"

   'shellcmd = "cmd /C """ & mtpath & "\bin\" & x64 & "latex.exe"" --output-directory=" & Mappe & " " & Mappe & "\" & filnavn '/K holder cmd åben /C lukker
   shellcmd = "cmd /C """ & mtpath & "\bin\" & x64 & "latex.exe"" --output-directory=" & Mappe & " " & filnavn & ".tex" '/K holder cmd åben /C lukker
   ' x64 er for 64bit installtion af miktex

   For j = 0 To 1
      On Error Resume Next
      If Dir(Mappe & "\" & filnavn & ".dvi") <> "" Then Kill Mappe & "\" & filnavn & ".dvi"
      'Kill Mappe & "\" & filnavn & ".dvi"
      On Error GoTo fejl

      appnr = Shell(shellcmd, vbNormalFocus) 'vbhide vbNormalFocus vbMinimizedNoFocus vbhide

      ' vent på maxima har genereret outputfil
      i = 0
      Do
         i = i + 1
         Wait (0.1)
      Loop Until Dir(Mappe & "\" & filnavn & ".dvi") <> "" Or i > 80
   Next

   Wait 1
   If Dir(Mappe & "\" & filnavn & ".dvi") > "" Then
      RunDefaultProgram Mappe & "\" & filnavn & ".dvi", Mappe
   Else
      MsgBox "Could not create dvi-file. The tex-file is now opening so you can identify the error", vbOKOnly, Sprog.Error
      RunDefaultProgram filnavn & ".tex", Mappe
   End If

   'RunDefaultProgram Mappe & "\" & filnavn & ".dvi", Mappe
   GoTo slut
fejl:
   MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
#End If
End Sub

Sub SavePdf(ByVal Mappe As String, ByVal filnavn As String)
   ' filnavn uden extension
   Dim shellcmd As String
   Dim appnr As Integer, p As Long, p2 As Long, ErrorS As String
   Dim i As Integer, j As Integer
   Dim mtpath As String, ReRun As Boolean
   
   '   On Error GoTo Fejl
#If Mac Then
   MsgBox Sprog.A(410), vbOKOnly, Sprog.Error
#Else
   Dim x64 As String
   
   ufwait.Label_tip.Caption = "Gemmer pdf-fil"
   
   MikTexPath ' called to set miktex64bit
   If MikTex64bit Then
      x64 = "x64\"
   End If
   
   Set ufwait = New UserFormWaitForMaxima
   
   ufwait.Label_progress.Caption = ufwait.Label_progress.Caption & "*"
   mtpath = MikTexPath
   If mtpath = "" Then
      MsgBox Sprog.A(409), vbOKOnly
      OpenLink "http://miktex.org"
      GoTo slut
   End If

   SaveTex Mappe, filnavn & ".tex"
   
   'shellcmd = "cmd /C """ & mtpath & "\bin\" & x64 & "pdflatex.exe"" --output-directory=" & Mappe & " " & Mappe & "\" & filnavn '/K holder cmd åben /C lukker
   shellcmd = "cmd /C """ & mtpath & "\bin\" & x64 & "pdflatex.exe"" --interaction=nonstopmode --halt-on-error --output-directory=" & Mappe & " " & filnavn & ".tex" '/K holder cmd åben /C lukker
   Do
      ufwait.Label_progress.Caption = ufwait.Label_progress.Caption & "**"
      ufwait.Label_tip = "Kører pdflatex " & j + 1 & ". gang"
      ReRun = False
      On Error Resume Next
      If Dir(Mappe & "\" & filnavn & ".pdf") <> "" Then Kill Mappe & "\" & filnavn & ".pdf"
      If Dir(Mappe & "\" & filnavn & ".log") <> "" Then Kill Mappe & "\" & filnavn & ".log"
      On Error GoTo 0 'Fejl

      Shell shellcmd, vbNormal
      'appnr = Shell(shellcmd, vbNormal) 'vbhide vbNormalFocus vbMinimizedNoFocus vbhide
      ' vent på maxima har genereret outputfil
      i = 0
      Do
         i = i + 1
         Wait (0.1)
      Loop Until Dir(Mappe & "\" & filnavn & ".log") > "" Or i > 80
      Wait 1
      Log = ReadTextFile(Mappe & "\" & filnavn & ".log")
      If InStr(Log, "Rerun to get cross-references right") > 0 Then ReRun = True
      ErrorS = ""
      p2 = 1
      p2 = 1
      p = InStr(p2, Log, "! ")
      p2 = InStr(p + 1, Log, "Here is how much of TeX")
      If p > 0 And p2 > p Then
         ErrorS = ErrorS & Mid(Log, p, p2 - p) & vbCrLf
      ElseIf p > 0 Then
         ErrorS = ErrorS & Mid(Log, p, Len(Log) - p - 1) & vbCrLf
      End If
      j = j + 1
   Loop While ReRun And j < 6
   
   If ErrorS <> "" Then
      If MsgBox("The tex-document failed to compile with the following error:" & vbCrLf & vbCrLf & ErrorS & vbCrLf & vbCrLf & "Click OK to open the tex-file", vbOKCancel, Sprog.Error) = vbOK Then RunDefaultProgram filnavn & ".tex", Mappe
   Else
      Wait 1
      If Dir(Mappe & "\" & filnavn & ".pdf") > "" Then
         RunDefaultProgram filnavn & ".pdf", Mappe
      Else
         If MsgBox("Could not locate the pdf-file. Click OK to open the tex-file so you can identify the error", vbOKOnly, Sprog.Error) = vbOK Then RunDefaultProgram filnavn & ".tex", Mappe
      End If
      If Dir(Mappe & "\" & filnavn & ".log") <> "" Then Kill Mappe & "\" & filnavn & ".log"
      If Dir(Mappe & "\" & filnavn & ".aux") <> "" Then Kill Mappe & "\" & filnavn & ".aux"
      If Dir(Mappe & "\" & filnavn & ".aux") <> "" Then Kill Mappe & "\" & filnavn & ".toc"
      
   End If
#End If
   GoTo slut
fejl:
   MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Function MikTexPath() As String
Dim path As String
     
path = Environ("ProgramW6432") & "\MikTex*"
If Dir(path, vbDirectory) <> "" Then
   MikTex64bit = True
   path = Environ("ProgramW6432") & "\" & Dir(path, vbDirectory) & "\miktex"
   GoTo slut
End If
path = Environ("ProgramFiles") & "\MikTex*"
If Dir(path, vbDirectory) <> "" Then
    MikTex64bit = False
   path = Environ("ProgramFiles") & "\" & Dir(path, vbDirectory) & "\miktex"
   GoTo slut
End If

path = Environ("USERPROFILE") & "\AppData\Local\Programs\MiKTeX 2.9\miktex" ' hvis installeret i userprofile
If Dir(path, vbDirectory) <> "" Then GoTo slut
'path = GetProgramFilesDir() & "\MikTex*"
'If Dir(path, vbDirectory) <> "" Then GoTo Slut

path = ""
slut:
MikTexPath = path
End Function
Function IsMikTexInstalled() As Boolean
If MikTexPath = "" Then
    IsMikTexInstalled = False
Else
    IsMikTexInstalled = True
End If
End Function
Sub InsertText(text As String)
    LaTexDoc = LaTexDoc & text ' & vbCr
End Sub
Sub InsertParagraph(text As String)
   If text = "" Or text = vbCrLf Or text = vbCr Or text = vbLf Then
      LaTexDoc = LaTexDoc & vbCrLf
   Else
    LaTexDoc = LaTexDoc & vbCrLf & text & vbCrLf
'    LaTexDoc = LaTexDoc & "\paragraph{" & text & "}" & vbCrLf
   End If
End Sub
Sub InsertChapter(text As String)
   If right(text, 1) = vbCr Then text = Left(text, Len(text) - 1)
    
   If InStr(text, "*") > 0 Then
      LaTexDoc = LaTexDoc & "\chapter*{" & Replace(text, "*", "") & "}" & vbCrLf
   Else
      LaTexDoc = LaTexDoc & "\chapter{" & text & "}" & vbCrLf
   End If
End Sub
Sub InsertSection(text As String)
   If right(text, 1) = vbCr Then text = Left(text, Len(text) - 1)
    
   If LatexSectionNumbering Then
      If InStr(text, "*") > 0 Then
         LaTexDoc = LaTexDoc & "\section*{" & Replace(text, "*", "") & "}" & vbCrLf
      Else
         LaTexDoc = LaTexDoc & "\section{" & text & "}" & vbCrLf
      End If
   Else
      LaTexDoc = LaTexDoc & "\section*{" & text & "}" & vbCrLf
   End If
End Sub
Sub InsertSubSection(text As String)
   If right(text, 1) = vbCr Then text = Left(text, Len(text) - 1)
    
   If LatexSectionNumbering Then
      If InStr(text, "*") > 0 Then
      LaTexDoc = LaTexDoc & "\subsection*{" & Replace(text, "*", "") & "}" & vbCrLf
      Else
      LaTexDoc = LaTexDoc & "\subsection{" & text & "}" & vbCrLf
      End If
   Else
      LaTexDoc = LaTexDoc & "\subsection*{" & text & "}" & vbCrLf
   End If
End Sub
Sub InsertSubSubSection(text As String)
   If right(text, 1) = vbCr Then text = Left(text, Len(text) - 1)
    
   If LatexSectionNumbering Then
      If InStr(text, "*") > 0 Then
         LaTexDoc = LaTexDoc & "\subsubsection*{" & Replace(text, "*", "") & "}" & vbCrLf
      Else
         LaTexDoc = LaTexDoc & "\subsubsection{" & text & "}" & vbCrLf
      End If
   Else
      LaTexDoc = LaTexDoc & "\subsubsection*{" & text & "}" & vbCrLf
   End If
End Sub

