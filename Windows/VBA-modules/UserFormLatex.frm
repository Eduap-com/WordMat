VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormLatex 
   Caption         =   "LaTex"
   ClientHeight    =   10590
   ClientLeft      =   -75
   ClientTop       =   -75
   ClientWidth     =   11115
   OleObjectBlob   =   "UserFormLatex.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormLatex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This form can convert math equations in Word to Latex, and convert the entire document to Tex
Public EventsOn As Boolean

Private EventsCol As New Collection
Sub SetEscEvents(ControlColl As Controls)
' SetEscEvents Me.Controls     in Initialize
    Dim CE As CEvents, c As control, TN As String, F As MSForms.Frame
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
        ElseIf TN = "MultiPage" Then
            Set CE = New CEvents: Set CE.MultiPageControl = c: EventsCol.Add CE
        ElseIf TN = "Frame" Then
            Set F = c
            SetEscEvents F.Controls
        End If
    Next
End Sub
Private Sub CheckBox_contents_Change()
   If EventsOn Then SaveSet2
   ShowFixedPreamble
End Sub

Private Sub CheckBox_forceMargins_Click()
   If EventsOn Then SaveSet2
   ShowFixedPreamble
End Sub

Private Sub CheckBox_sectionnumbers_Change()
   If EventsOn Then SaveSet2
   ShowFixedPreamble
End Sub

Private Sub CheckBox_title_Change()
   If EventsOn Then SaveSet2
   ShowFixedPreamble
End Sub

Private Sub CheckBox_units_Click()
    UpDateLatex
End Sub

Private Sub ComboBox_documentclass_Change()
   If EventsOn Then SaveSet2
   ShowFixedPreamble
End Sub

Private Sub ComboBox_fontsize_Change()
   If EventsOn Then SaveSet2
   ShowFixedPreamble
End Sub

Private Sub CommandButton_convertall_Click()
    Me.hide
    SaveSet
    ConvertAllEquations
    Me.hide

End Sub

Sub ShowFixedPreamble()
   latexfil.UseWordMargins = LatexWordMargins
'   latexfil.ImagDir = ""
   TextBox_FixedPreamble.Text = latexfil.FixedLatexPreamble1 & vbCrLf & "... Custom ..." & vbCrLf & latexfil.FixedLatexPreamble2
End Sub

Sub SaveSet2()
   If Not EventsOn Then Exit Sub
    LatexPreamble = TextBox_preamble.Text
    LatexSectionNumbering = CheckBox_sectionnumbers.Value
    LatexDocumentclass = ComboBox_documentclass.ListIndex
    LatexFontsize = ComboBox_fontsize.ListIndex + 10
    LatexWordMargins = CheckBox_forceMargins.Value
    TextBox_FixedPreamble.Text = latexfil.FixedLatexPreamble1
    LatexTOC = CInt(CheckBox_contents.Value)
    LatexTitlePage = CInt(CheckBox_title.Value)
End Sub
Sub SaveSet()
    LatexUnits = CheckBox_units.Value
    ConvertTexWithMaxima = CheckBox_convertwithmaxima.Value
    
If OptionButton_omslutdollar.Value = True Then
    LatexStart = "$"
    LatexSlut = "$"
ElseIf OptionButton_omslutdobbeltdollar.Value = True Then
    LatexStart = "$$"
    LatexSlut = "$$"
ElseIf OptionButton_omslutsqbrackets.Value = True Then
    LatexStart = "\[ "
    LatexSlut = " \]"
ElseIf OptionButton_omslutdispmath.Value = True Then
    LatexStart = vbCrLf & "\displaymath" & vbCrLf
    LatexSlut = vbCrLf & "\displaymath" & vbCrLf
ElseIf OptionButton_omsluteqn.Value = True Then
    LatexStart = vbCrLf & "\begin{equation}" & vbCrLf
    LatexSlut = vbCrLf & "\end{equation}" & vbCrLf
ElseIf OptionButton_omsluteqnstar.Value = True Then
    LatexStart = vbCrLf & "\begin{equation*}" & vbCrLf
    LatexSlut = vbCrLf & "\end{equation*}" & vbCrLf
ElseIf OptionButton_omsluturl.Value = True Then
    LatexStart = vbCrLf & "<img src=""https://latex.codecogs.com/gif.latex?"
    LatexSlut = """ title=""LaTex"" />" & vbCrLf
ElseIf OptionButton_omslutingen.Value = True Then
    LatexStart = ""
    LatexSlut = ""
ElseIf OptionButton_omslutlatex.Value = True Then
    LatexStart = "[latex]"
    LatexSlut = "[\latex]"
ElseIf OptionButton_omslutuser.Value = True Then
    LatexStart = TextBox_for.Text
    LatexSlut = TextBox_efter.Text
ElseIf OptionButton_omslutauto.Value = True Then
   If Selection.OMaths.Count > 0 Then
    If Selection.OMaths(1).Justification = wdOMathJcInline Then
        LatexStart = "$"
        LatexSlut = "$"
    Else
        LatexStart = "\[ "
        LatexSlut = " \]"
    End If
    End If
Else
    LatexStart = ""
    LatexSlut = ""
End If
    
End Sub
Private Sub CommandButton_copy_Click()
Dim Obj As New DataObject

Obj.SetText TextBox_latex.Text
Obj.PutInClipboard
End Sub

Private Sub CommandButton_latex_Click()
   Me.hide
    SaveFile 2
    'open latex
End Sub

Private Sub CommandButton_next_Click()
    If Selection.OMaths.Count > 0 Then
        Selection.OMaths(1).Range.Text = ""
        Selection.InsertAfter TextBox_latex.Text
    End If
    Me.hide
    Selection.End = ActiveDocument.Range.End
    If Selection.OMaths.Count > 0 Then
        Selection.OMaths(1).Range.Select
        omax.ReadSelection
        Label_input.Caption = omax.Kommando
    Else
        Label_input.Caption = ""
        TextBox_latex.Text = ""
    End If
    UpDateLatex
    Me.Show
End Sub

Private Sub CommandButton_ok_Click()
   SaveSet2
   SaveSet

Me.hide
End Sub
Private Sub CommandButton_onlinelatex_Click()
'https://latex.codecogs.com/emf.latex?%5Cint_0%5E1%20x%5E2%20dx
'https://www.codecogs.com/latex/eqneditor.php?latex=x^2+1
Dim Text As String

Text = LatexCode

Text = Replace(Text, "^", "%5E")
Text = Replace(Text, "&", "%26")
Text = Replace(Text, "\", "%5C")
Text = Replace(Text, " ", "%20")
Text = Replace(Text, "+", "@plus;")
OpenLink "https://www.codecogs.com/latex/eqneditor.php?latex=" & Text  '"%5Cint_0%5E1%20x%5E2%20dx"

End Sub

Private Sub CommandButton_dvi_Click()
    SaveFile (1)
End Sub

Private Sub CommandButton_pdflatex_Click()
   Me.hide
    SaveFile (0)
End Sub

Private Sub OptionButton_omslutauto_click()
    UpDateLatex
End Sub

Private Sub OptionButton_omslutdispmath_click()
    UpDateLatex
End Sub

Private Sub OptionButton_omslutdobbeltdollar_click()
    UpDateLatex
End Sub

Private Sub OptionButton_omslutdollar_click()
    UpDateLatex
End Sub

Private Sub OptionButton_omsluteqn_click()
    UpDateLatex
End Sub

Private Sub OptionButton_omsluteqnstar_click()
    UpDateLatex
End Sub

Private Sub OptionButton_omslutingen_click()
    UpDateLatex
End Sub

Private Sub OptionButton_omslutlatex_click()
    UpDateLatex
End Sub

Private Sub OptionButton_omslutsqbrackets_click()
    UpDateLatex
End Sub

Private Sub OptionButton_omsluturl_click()
    UpDateLatex
End Sub

Private Sub OptionButton_omslutuser_click()
    If OptionButton_omslutuser.Value = True Then
        TextBox_for.visible = True
        TextBox_efter.visible = True
        Label_for.visible = True
        Label_efter.visible = True
    Else
        TextBox_for.visible = False
        TextBox_efter.visible = False
        Label_for.visible = False
        Label_efter.visible = False
    End If
    UpDateLatex
End Sub

Private Sub OptionButton_visauto_click()
    UpDateLatex
End Sub
Private Sub OptionButton_visstor_click()
    UpDateLatex
End Sub
Private Sub OptionButton_visinline_click()
    UpDateLatex
End Sub

Private Sub TextBox_efter_click()
    UpDateLatex
End Sub

Private Sub TextBox_for_click()
    UpDateLatex
End Sub

Private Sub UserForm_Activate()
    SaveBackup
    Application.ScreenUpdating = True
    SetCaptions
    FillComboboxDocumentclass
    FillComboboxFontsize
    
    EventsOn = False
    Selection.End = ActiveDocument.Range.End
    TextBox_for.Text = ""
    TextBox_efter.Text = ""
    If Selection.OMaths.Count = 0 Then
'        MsgBox TT.A(84), vbOKOnly, TT.Error
    Else
        Selection.OMaths(1).Range.Select
        omax.ReadSelection
        Label_input.Caption = omax.Kommando
    End If
    
    CheckBox_units.Value = LatexUnits
    CheckBox_convertwithmaxima.Value = ConvertTexWithMaxima
    CheckBox_sectionnumbers.Value = LatexSectionNumbering
    OptionButton_omslutauto.Value = True
    ComboBox_documentclass.ListIndex = LatexDocumentclass
    If LatexFontsize = "10" Then
       ComboBox_fontsize.ListIndex = 0
    ElseIf LatexFontsize = "11" Then
       ComboBox_fontsize.ListIndex = 1
    ElseIf LatexFontsize = "12" Then
       ComboBox_fontsize.ListIndex = 2
    Else
       ComboBox_fontsize.ListIndex = 0
    End If
    TextBox_preamble.Text = LatexPreamble
    ShowFixedPreamble
    CheckBox_forceMargins.Value = LatexWordMargins
    CheckBox_title.Value = CBool(LatexTitlePage)
    CheckBox_contents.Value = CBool(LatexTOC)
'    If LatexStart = "$" And LatexSlut = "$" Then
'        OptionButton_omslutdollar.Value = True
'    ElseIf LatexStart = "$$" And LatexSlut = "$$" Then
'        OptionButton_omslutdobbeltdollar.Value = True
'    Else
'        OptionButton_omslutuser.Value = True
'        TextBox_for.Text = LatexStart
'        TextBox_efter.Text = LatexSlut
'    End If
    EventsOn = True
    UpDateLatex
End Sub

Sub UpDateLatex()
   Dim t As Table
   If Not EventsOn Then Exit Sub
   If Selection.OMaths.Count = 0 Then Exit Sub
   SaveSet

   Label_input.Caption = omax.Kommando
   LatexCode = omax.ConvertToLatex(omax.Kommando)
   If OptionButton_visstor.Value = True Then
      LatexCode = "\displaystyle " & LatexCode
   ElseIf OptionButton_visinline.Value = True Then
      LatexCode = "\inline " & LatexCode
   End If

   If OptionButton_omslutauto.Value = True Then
      If Selection.OMaths(1).Justification = wdOMathJcInline Then
         TextBox_latex.Text = "$" & LatexCode & "$"
      Else
         If Selection.OMaths(1).Range.Tables.Count > 0 Then
            Set t = Selection.OMaths(1).Range.Tables(1)
            If t.Rows.Count = 1 And t.Columns.Count = 3 And t.Cell(1, 2).Range.OMaths.Count > 0 And t.Cell(1, 3).Range.Fields.Count Then
               
               TextBox_latex.Text = "\begin{equation}" & LatexCode & "\end{equation}"
            Else
               TextBox_latex.Text = "\begin{equation*}" & LatexCode & "\end{equation*}"
            End If
         Else
            TextBox_latex.Text = "\begin{equation*}" & LatexCode & "\end{equation*}"
         End If
      End If
   Else
      TextBox_latex.Text = LatexStart & LatexCode & LatexSlut
   End If

End Sub

Sub SetCaptions()
    Me.Caption = "LaTex"
    Label2.Caption = ChrW(&H2192)
    CommandButton_pdflatex.Caption = ChrW(&H2192) & " PDF"
'    CommandButton_dvi.Caption = ChrW(&H2192) & " dvi (YAP)"
    CommandButton_latex.Caption = ChrW(&H2192) & " Tex"
    CommandButton_ok.Caption = TT.A(661)
    Label1.Caption = TT.A(72)
    CommandButton_copy.Caption = TT.A(73)
    Label_status.Caption = TT.A(826)
    Frame1.Caption = TT.A(83)
    Frame2.Caption = TT.A(74)
    CheckBox_units.Caption = TT.A(75)
    CheckBox_convertwithmaxima.Caption = TT.A(76)
    CommandButton_convertall.Caption = TT.A(77)
    CommandButton_next.Caption = TT.A(78)
    Label_for.Caption = TT.A(79)
    Label_efter.Caption = TT.A(80)
    OptionButton_omslutingen.Caption = TT.A(81)
    OptionButton_omslutuser.Caption = TT.A(82)
    CheckBox_convertwithmaxima.ControlTipText = TT.A(659)
    CheckBox_units.ControlTipText = TT.A(660)
    Frame3.ControlTipText = TT.A(662)
    CommandButton_onlinelatex.ControlTipText = TT.A(663)
    CommandButton_latex.ControlTipText = TT.A(664)
End Sub

Sub FillComboboxDocumentclass()
   ComboBox_documentclass.Clear
   ComboBox_documentclass.AddItem "Article"
   ComboBox_documentclass.AddItem "Report"
   ComboBox_documentclass.AddItem "Book"
End Sub
Sub FillComboboxFontsize()
   ComboBox_fontsize.Clear
   ComboBox_fontsize.AddItem "10pt"
   ComboBox_fontsize.AddItem "11pt"
   ComboBox_fontsize.AddItem "12pt"
End Sub

Private Sub CommandButton_resetpreamble_Click()
   Dim s As String

'   s = s & "\documentclass[11pt]{article}" & vbCrLf
'   s = s & "\usepackage[T1]{fontenc}" & vbCrLf
'   s = s & "\usepackage[latin1]{inputenc}" & vbCrLf
'   s = s & "\usepackage{geometry}" & vbCrLf
'   s = s & "\geometry{a4paper}" & vbCrLf
'   s = s & "\usepackage{graphicx}" & vbCrLf
'   If TT.LangNo = 1 Then
'      s = s & "\graphicspath{{" & Replace(ActiveDocument.path, "\", "/") & "/" & Split(ActiveDocument.Name, ".")(0) & "Images-filer/}}" & vbCrLf
'   Else
'      s = s & "\graphicspath{{" & Replace(ActiveDocument.path, "\", "/") & "/" & Split(ActiveDocument.Name, ".")(0) & "Images-files/}}" & vbCrLf
'   End If
   
'   s = s & "\usepackage{booktabs} % added functionality to tables" & vbCrLf
'   s = s & "\usepackage{array}" & vbCrLf
'   s = s & "\usepackage{paralist} % extented functionality for list within paragrahs etc." & vbCrLf
'   s = s & "\usepackage{verbatim} % \begin{verbatim} used for entering latex commands in the text. This package fixes issues with the buitin" & vbCrLf
'   s = s & "\usepackage{subfig} % used to caption subfigues within figure environment " & vbCrLf
   
'   s = s & "\usepackage{fancyhdr}" & vbCrLf
'   s = s & "\renewcommand{\headrulewidth}{0pt}" & vbCrLf
'   s = s & "\lhead{}\chead{}\rhead{}" & vbCrLf
'   s = s & "\lfoot{}\cfoot{\thepage}\rfoot{}" & vbCrLf
'   s = s & "\usepackage{sectsty}" & vbCrLf
'   s = s & "\allsectionsfont{\sffamily\mdseries\upshape}" & vbCrLf
'   s = s & "\usepackage[nottoc,notlof,notlot]{tocbibind}" & vbCrLf
'   s = s & "\usepackage[titles,subfigure]{tocloft}" & vbCrLf
'   s = s & "\renewcommand{\cftsecfont}{\rmfamily\mdseries\upshape}" & vbCrLf
'   s = s & "\renewcommand{\cftsecpagefont}{\rmfamily\mdseries\upshape}" & vbCrLf

'   s = s & "\title{" & Titel & "}" & vbCrLf
'   s = s & "\author{" & Author & "}" & vbCrLf
   

   LatexPreamble = s
   TextBox_preamble.Text = s

End Sub

Private Sub UserForm_Initialize()
    SetEscEvents Me.Controls
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then
        Me.hide
    End If
End Sub
