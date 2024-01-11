VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormSolveNumeric 
   Caption         =   "Løsning af ligning med grafiske og numeriske metoder"
   ClientHeight    =   7395
   ClientLeft      =   -15
   ClientTop       =   75
   ClientWidth     =   14250
   OleObjectBlob   =   "UserFormSolveNumeric.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormSolveNumeric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit
Public udtryk As String
Public dispudtryk As String
Public vars As String
Public SelectedVar As String
Public Method As String
Private gemx As Single
Private gemy As Single
Private gemslutr As Integer
Private gemstartr As Integer
Private gemr As Range

Private Sub SaveVar()
On Error GoTo fejl
'    If TextBox_variabel.text = "" Then
'        SelectedVar = ListBox_vars.value
'    Else
'        SelectedVar = TextBox_variabel.text
'    End If
        SelectedVar = TextBox_variabel.Text
    
    TextBox_guess.Text = Replace(TextBox_guess.Text, ",", ".")
    TextBox_lval.Text = Replace(TextBox_lval.Text, ",", ".")
    TextBox_hval.Text = Replace(TextBox_hval.Text, ",", ".")
    
    GoTo slut
fejl:
    SelectedVar = ""
slut:
End Sub

Private Sub CommandButton_findroot_Click()
    SaveVar
    Method = "findroot"
    
    GoTo slut
fejl:
slut:
    Selection.start = gemstartr
    Selection.End = gemslutr
    Application.ScreenUpdating = False
    Me.hide

End Sub

Private Sub CommandButton_insertpic_Click()
Dim ils As InlineShape
Dim s As String, Arr As Variant, Sep As String
On Error GoTo fejl
omax.GoToEndOfSelectedMaths
Selection.TypeParagraph
Arr = Split(Label_ligning.Caption, "=")
#If Mac Then
    Set ils = Selection.InlineShapes.AddPicture(GetTempDir() & "WordMatGraf.pdf", False, True)
#Else
    Set ils = Selection.InlineShapes.AddPicture(GetTempDir() & "WordMatGraf.gif", False, True)
#End If
'Set ils = Selection.InlineShapes.AddPicture(Environ("TEMP") & "\WordMatGraf.gif", False, True)
Sep = "|"
s = "WordMat" & Sep & AppVersion & Sep & "" & Sep & "" & Sep & TextBox_variabel.Text & Sep & "" & Sep
s = s & TextBox_xmin.Text & Sep & TextBox_xmax.Text & Sep & "" & Sep & "" & Sep
s = s & Arr(0) & Sep & TextBox_variabel.Text & Sep & "" & Sep & "" & Sep & "" & Sep
s = s & Arr(1) & Sep & TextBox_variabel.Text & Sep & "" & Sep & "" & Sep & "" & Sep
ils.AlternativeText = s
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton_ok_Click()
    ' ok til newton
    SaveVar
    Method = "newton"
    
    GoTo slut
fejl:
slut:
    Selection.start = gemstartr
    Selection.End = gemslutr
    Application.ScreenUpdating = False
    Me.hide

End Sub

Private Sub CommandButton_opdater_Click()
OpdaterGraf
#If Mac Then
    ShowPreviewMac
#Else
    Me.Repaint
#End If
End Sub

Private Sub CommandButton_visgraf_Click()
Dim Text As String
On Error GoTo fejl
'Dim omax As New CMaxima
    If omax Is Nothing Then
        Set omax = New CMaxima
        If MaxProc Is Nothing Then
'        Set MaxProc = New MathMenu.MaximaProcessClass
        Set MaxProc = GetMaxProc() 'CreateObject("MaximaProcessClass")
        End If
    End If
    omax.PrepareNewCommand ' nulstiller og finder definitioner
    
    If TextBox_xmin.Text = "" Then TextBox_xmin = -5
    If TextBox_xmax.Text = "" Then TextBox_xmax = 5
    
    udtryk = Replace(udtryk, "=", ";")
'    Me.Hide
    Call omax.Plot2D(udtryk, "", TextBox_variabel.Text, Replace(TextBox_xmin.Text, ",", "."), Replace(TextBox_xmax.Text, ",", "."), "", "", "", "")
       
    omax.PrepareNewCommand ' nødvendigt da der efterfølgende skal køres newton el lign eller vises grafen igen
    DoEvents
'    Me.Show
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Private Sub CommandButton_zoom_Click()
Dim dx As Single
Dim midt As Single
Dim xmin As Single
Dim xmax As Single
On Error GoTo fejl
xmin = CSng(TextBox_xmin.Text)
xmax = CSng(TextBox_xmax.Text)

midt = (xmax + xmin) / 2
dx = (xmax - xmin) * 0.3

TextBox_xmin.Text = xmin + dx
TextBox_xmax.Text = xmax - dx
OpdaterGraf

Me.Repaint
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:


End Sub

Private Sub CommandButton_zoomud_Click()
Dim dx As Single
Dim midt As Single
Dim xmin As Single
Dim xmax As Single
On Error GoTo fejl
xmin = CSng(TextBox_xmin.Text)
xmax = CSng(TextBox_xmax.Text)

midt = (xmax + xmin) / 2
dx = (xmax - xmin) * 0.4

TextBox_xmin.Text = xmin - dx
TextBox_xmax.Text = xmax + dx
OpdaterGraf

Me.Repaint
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:

End Sub

Private Sub Image1_Click()
'    Image1.PictureSizeMode = fmPictureSizeModeZoom
'    Image1.Picture = LoadPicture("c:/WordMatGraf.gif")
'    Call Image1.Picture.Render(Image1.Picture.Handle, 0, 0, 300, 300, 2, 2, 300, 300, vbNull)
    
'    Image1.Move 1, 1, 1, 1, 0
End Sub

Private Sub Image1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim xmin As Single
Dim xmax As Single
Dim cfakt As Single
Dim dx As Single
Dim midt As Single
Dim X As Single
Label_zoom.visible = False

xmin = CSng(TextBox_xmin.Text)
xmax = CSng(TextBox_xmax.Text)
dx = (xmax - xmin) * 0.3
cfakt = (xmax - xmin) / (Image1.Width * 0.85)
'MsgBox Image1.Picture.Width
'MsgBox "xmax  -  " & xmax & "  xmin-  " & xmin & " gemx-" & gemx & "x-" & X & " cfakt " & cfakt & "cfakt*x" & cfakt * X
X = gemx - Image1.Width * 0.1
gemx = gemx - Image1.Width * 0.1
TextBox_xmin.Text = xmin + cfakt * X - dx
TextBox_xmax.Text = xmin + cfakt * X + dx
OpdaterGraf

    Me.Repaint
GoTo slut
fejl:
    MsgBox "Der skete en fejl. Se på tallene ved xmin og xmax", vbOKOnly, Sprog.Error
slut:
    
End Sub

Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    gemx = X
    gemy = Y
    Label_zoom.Left = gemx + Image1.Left
    Label_zoom.Top = gemy + Image1.Top
    Label_zoom.Width = 1
    Label_zoom.Height = 1
    Label_zoom.visible = True
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
Dim cfakt As Single

Label_zoom.visible = False
If Abs(X - gemx) < 5 Then GoTo slut

xmin = CSng(TextBox_xmin.Text)
xmax = CSng(TextBox_xmax.Text)
cfakt = (xmax - xmin) / (Image1.Width * 0.85)
'MsgBox Image1.Picture.Width
'MsgBox "xmax  -  " & xmax & "  xmin-  " & xmin & " gemx-" & gemx & "x-" & X & " cfakt " & cfakt & "cfakt*x" & cfakt * X
X = X - Image1.Width * 0.1
gemx = gemx - Image1.Width * 0.1
TextBox_xmax.Text = xmin + cfakt * X
TextBox_xmin.Text = xmin + cfakt * gemx
OpdaterGraf

    Me.Repaint
slut:
End Sub

Private Sub Label_intervalhelp_Click()
    MsgBox Sprog.A(675), vbOKOnly, Sprog.A(676)
End Sub

Private Sub Label_newtonhelp_Click()
    MsgBox Sprog.A(677), vbOKOnly, Sprog.A(678)
End Sub

Private Sub TextBox_xmax_AfterUpdate()
'    OpdaterGraf
End Sub

Private Sub TextBox_xmin_AfterUpdate()
'    OpdaterGraf

End Sub
Sub OpdaterGraf()
On Error GoTo fejl
#If Mac Then
Dim Text As String
Dim Arr As Variant
    Arr = Split(Label_ligning.Caption, "=")
    
    Text = "line_width=2,color=green,explicit(" & Arr(0) & "," & TextBox_variabel.Text & "," & TextBox_xmin.Text & "," & TextBox_xmax.Text & "),"
    Text = Text & "color=red,explicit(" & Arr(1) & "," & TextBox_variabel.Text & "," & TextBox_xmin.Text & "," & TextBox_xmax.Text & ")"
'    If Len(TextBox_xmin.text) > 0 And Len(TextBox_xmax.text) > 0 Then
'        text = "xrange=[" & ConvertNumberToMaxima(TextBox_xmin.text) & "," & ConvertNumberToMaxima(TextBox_xmax.text) & "]"
'    End If
'    If Len(TextBox_ymin.text) > 0 And Len(TextBox_ymax.text) > 0 And Len(TextBox_dfligning.text) = 0 Then
'        text = text & ",yrange=[" & ConvertNumberToMaxima(TextBox_ymin.text) & "," & ConvertNumberToMaxima(TextBox_ymax.text) & "]"
'    End If
    Call omax.Draw2D(Text, "", TextBox_variabel.Text, "y", True, True, 3)
#Else
    Call omax.Plot2D(udtryk, "", TextBox_variabel.Text, Replace(TextBox_xmin.Text, ",", "."), Replace(TextBox_xmax.Text, ",", "."), "", "", "", "", True)
#End If
    
    TextBox_guess.Text = (CSng(TextBox_xmax.Text) + CSng(TextBox_xmin.Text)) / 2
    TextBox_lval.Text = TextBox_xmin.Text
    TextBox_hval.Text = TextBox_xmax.Text
    
    omax.PrepareNewCommand ' nødvendigt da der efterfølgende skal køres newton el lign eller vises grafen igen
    DoEvents
#If Mac Then
    ShowPreviewMac
#Else
    Image1.Picture = LoadPicture(GetTempDir() & "WordMatGraf.gif")
#End If
'    Image1.Picture = LoadPicture(Environ("TEMP") & "\WordMatGraf.gif")
GoTo slut
fejl:
    MsgBox "Der skete en fejl. Prøv at trykke Opdater.", vbOKOnly, Sprog.Error
slut:

End Sub
Private Sub UserForm_Activate()
Dim Arr As Variant
Dim i As Integer
On Error Resume Next
    SetCaptions
#If Mac Then
    Me.Left = 10
    Me.Top = 80
    Me.Width = 113
    Me.Height = 500
    CommandButton_insertpic.Left = 5
    CommandButton_insertpic.Top = 380
    TextBox_xmin.Left = 5
    TextBox_xmin.Top = 315
    TextBox_xmax.Left = 50
    TextBox_xmax.Top = 315
    Label5.Left = 5
    Label5.Top = 300
    Label6.Left = 50
    Label6.Top = 300
    CommandButton_opdater.Left = 25
    CommandButton_opdater.Top = 350
    
'    Label_zoom.visible = False
    Kill GetTempDir() & "WordMatGraf.pdf"
#Else
    Kill GetTempDir() & "\WordMatGraf.gif"
#End If

On Error GoTo fejl
    SelectedVar = ""
    ListBox_vars.Clear
    TextBox_guess.Text = "1"
    TextBox_xmin.Text = "-5"
    TextBox_xmax.Text = "5"
    Label_ligning.Caption = omax.ConvertToAscii(udtryk)
    Arr = Split(vars, ";")
    Set gemr = Selection.Range
    gemstartr = Selection.Range.start
    gemslutr = Selection.Range.End
'    For i = 0 To UBound(arr)
'        If arr(i) <> "" Then
'            ListBox_vars.AddItem (arr(i))
'        End If
'    Next
'    If ListBox_vars.ListCount > 0 Then
'        ListBox_vars.ListIndex = SelVarIndex
'    End If
    If UBound(Arr) >= 0 And TextBox_variabel.Text = vbNullString Then
        TextBox_variabel.Text = Arr(0)
    End If
    
'    If omax Is Nothing Then
'        Set omax = New CMaxima
'        If MaxProc Is Nothing Then
''        Set MaxProc = New MathMenu.MaximaProcessClass
'        Set MaxProc = CreateObject("MaximaProcessClass")
'        End If
'    End If
    omax.PrepareNewCommand ' nulstiller og finder definitioner
        
    udtryk = Replace(udtryk, "=", ";")
'    Me.Hide

    OpdaterGraf
    
    CommandButton_ok.SetFocus

GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
'    newton.Select

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'  If CloseMode = 0 Then
'    Cancel = 1
'    SelectedVar = ""
'    Me.Hide
'  End If
Unload Me
End Sub

Sub SetCaptions()
    Me.Caption = Sprog.A(235)
    Label14.Caption = Sprog.Equation
    Label_variabel.Caption = Sprog.Variable
    Label1.Caption = Sprog.A(236)
    Frame3.Caption = Sprog.A(237)
    Label9.Caption = Sprog.A(238)
    Label10.Caption = Sprog.A(239)
    Label12.Caption = Sprog.A(240)
    Label13.Caption = Sprog.A(241)
    CommandButton_opdater.Caption = Sprog.Update
    Label15.Caption = Sprog.A(242)
    CommandButton_visgraf.Caption = Sprog.A(243)
    CommandButton_insertpic.Caption = Sprog.A(93)
End Sub
Sub ShowPreviewMac()
#If Mac Then
    RunScript "OpenPreview", GetTempDir() & "WordMatGraf.pdf"
#End If
End Sub


