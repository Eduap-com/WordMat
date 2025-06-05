VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormDiffEq 
   Caption         =   "Solve differential equation"
   ClientHeight    =   4920
   ClientLeft      =   -30
   ClientTop       =   75
   ClientWidth     =   6855
   OleObjectBlob   =   "UserFormDiffEq.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormDiffEq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Form to enter variables and conditions for solving differential equations

Public DefS As String
Public Vars As String
Public TempDefs As String
Public luk As Boolean
Private Svars As Variant ' array to hold variablenames to be returned (without asciiconvertion)

Private EventsCol As New Collection
Sub SetEscEvents(ControlColl As Controls)
' SetEscEvents Me.Controls     in Initialize
    Dim CE As CEvents, c As control, TN As String, F As MSForms.Frame
    On Error Resume Next
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
        ElseIf TN = "Frame" Then
            Set F = c
            SetEscEvents F.Controls
        End If
    Next
End Sub
Private Sub Label_cancel_Click()
    luk = True
    Me.hide
End Sub

Private Sub CommandButton_ok_Click()
Dim Arr As Variant
Dim i As Integer
    
    luk = False
    TempDefs = TextBox_def.Text
    TempDefs = Trim(TempDefs)
    If Len(TempDefs) > 2 Then
    TempDefs = Replace(TempDefs, ",", ".")
    Arr = Split(TempDefs, VbCrLfMac)
    TempDefs = ""
    For i = 0 To UBound(Arr)
        If Len(Arr(i)) > 2 And Not Right(Arr(i), 1) = "=" Then
            If Split(Arr(i), "=")(0) <> TextBox_funktion.Text Then ' cant define variable to be solved for
                TempDefs = TempDefs & omax.CodeForMaxima(Arr(i)) & ListSeparator
            Else
                MsgBox TT.A(252) & " " & TextBox_funktion.Text & " " & TT.A(253), vbOKOnly, TT.Error
                Exit Sub
            End If
        End If
    Next
    If Right(TempDefs, 1) = ListSeparator Then
        TempDefs = Left(TempDefs, Len(TempDefs) - 1)
    End If
    End If
    
    Me.hide
End Sub

Private Sub Label_solvenum_Click()
   Dim F As String, Arr() As String, sx As String, sy As String
   Arr = Split(Label_ligning.Caption, "=")
   If UBound(Arr) > 0 Then F = Trim(Arr(1))
   If Len(Trim(Arr(0))) > 2 Then
      MsgBox TT.A(847), vbOKOnly, TT.Error
      Exit Sub
   End If
   luk = True
   Me.hide
   
   If TextBox_startx.Text = vbNullString Then
      sx = "1"
   Else
      sx = TextBox_startx.Text
   End If
   If TextBox_starty.Text = vbNullString Then
      sy = "1"
   Else
      sy = TextBox_starty.Text
   End If
   
   PlotDF F, TextBox_variabel.Text, TextBox_funktion.Text, "(" & sx & ", " & sy & ")"
   
Exit Sub
    ' this was previously used
   UserFormDeSolveNumeric.TextBox_varx.Text = TextBox_variabel.Text
   UserFormDeSolveNumeric.TextBox_var1.Text = TextBox_funktion.Text
   UserFormDeSolveNumeric.TextBox_eq1.Text = F
   If TextBox_starty.Text = vbNullString Then
      UserFormDeSolveNumeric.TextBox_init1.Text = "1"
   Else
      UserFormDeSolveNumeric.TextBox_init1.Text = TextBox_starty.Text
   End If
   UserFormDeSolveNumeric.TextBox_xmin.Text = TextBox_startx.Text
   UserFormDeSolveNumeric.TextBox_xmax.Text = TextBox_startx.Text + 10
   UserFormDeSolveNumeric.TextBox_definitioner.Text = TextBox_def.Text
   UserFormDeSolveNumeric.Show
End Sub


Private Sub Label_ok_Click()
    CommandButton_ok_Click
End Sub

Private Sub TextBox_funktion_Change()
    opdaterLabels
End Sub

Private Sub TextBox_startx_Change()
    opdaterLabels
End Sub

Private Sub UserForm_Activate()
Dim i As Integer
Dim svar As String
    SetCaptions
    
    luk = True
    Label_ligning.Caption = FormatDefinitions(Replace(Label_ligning.Caption, "=", " = "))
    Label_ligning.Caption = Replace(Label_ligning.Caption, ChrW(180), "'") ' converttoascii introduces 180. It may be neccesary, but this looks better

    If InStr(Label_ligning.Caption, "''") > 0 Then ' "´´" two accents goint up
        Label_diffy.visible = True
        TextBox_starty2.visible = True
        Label_y2.visible = True
        TextBox_bcx.visible = True
        Label7.visible = True
        Label8.visible = True
        TextBox_bcy.visible = True
        Label_solvenum.visible = False
    Else
        Label_diffy.visible = False
        TextBox_starty2.visible = False
        Label_y2.visible = False
        TextBox_bcx.visible = False
        Label7.visible = False
        Label8.visible = False
        TextBox_bcy.visible = False
        Label_solvenum.visible = True
    End If

    Svars = Split(Vars, ";")
    For i = 0 To UBound(Svars)
        If Svars(i) <> "" Then
            svar = omax.ConvertToWordSymbols(Svars(i))
            TextBox_def.Text = TextBox_def.Text & svar & "=" & VbCrLfMac    ' temp definitions
        End If
    Next
End Sub

Private Sub UserForm_Initialize()
    SetEscEvents Me.Controls
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    luk = True
End Sub

Sub opdaterLabels()
Dim fkt As String
Dim Pos As Integer
On Error Resume Next
    fkt = TextBox_funktion.Text
    Pos = InStr(fkt, "(")
    If Pos > 0 Then
        fkt = Left(fkt, Pos - 1)
    End If
    Label_diffy.Caption = fkt & "'(" & TextBox_startx.Text & ")="
    Label_y.Caption = fkt & "("
    Label_y2.Caption = fkt & "("
End Sub

Sub SetCaptions()
    Me.Caption = TT.A(745)
    Label1.Caption = TT.A(744) & ":"
    Label3.Caption = TT.A(746)
    Label2.Caption = TT.A(747)
    Label4.Caption = TT.A(748) & ":"
    Label8.Caption = TT.A(297)
    Label_temp.Caption = TT.A(764)
    Label_cancel.Caption = TT.Cancel
    Label_ok.Caption = TT.OK
    Label_solvenum.Caption = TT.A(887)
End Sub
Private Sub Label_solvenum_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_solvenum.BackColor = LBColorPress
End Sub
Private Sub Label_solvenum_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_solvenum.BackColor = LBColorHover
End Sub
Private Sub Label_cancel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorPress
End Sub
Private Sub Label_cancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorHover
End Sub
Private Sub Label_ok_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorPress
End Sub
Private Sub Label_ok_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorHover
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorInactive
    Label_cancel.BackColor = LBColorInactive
    Label_solvenum.BackColor = LBColorInactive
End Sub
