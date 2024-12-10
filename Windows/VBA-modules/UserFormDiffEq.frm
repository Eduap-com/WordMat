VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormDiffEq 
   Caption         =   "Løsning af differentialligning"
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

Public DefS As String
Public Vars As String
Public TempDefs As String
Public luk As Boolean
Private Svars As Variant ' array der holder variabelnavne som de skal returneres dvs. uden asciikonvertering

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
    TempDefs = TextBox_def.text
    TempDefs = Trim(TempDefs)
    If Len(TempDefs) > 2 Then
    TempDefs = Replace(TempDefs, ",", ".")
    Arr = Split(TempDefs, VbCrLfMac)
    TempDefs = ""
    For i = 0 To UBound(Arr)
        If Len(Arr(i)) > 2 And Not right(Arr(i), 1) = "=" Then
            If Split(Arr(i), "=")(0) <> TextBox_funktion.text Then ' kan ikke definere variabel der løses for
                TempDefs = TempDefs & omax.CodeForMaxima(Arr(i)) & ListSeparator
            Else
                MsgBox Sprog.A(252) & " " & TextBox_funktion.text & " " & Sprog.A(253), vbOKOnly, Sprog.Error
                Exit Sub
            End If
        End If
    Next
    If right(TempDefs, 1) = ListSeparator Then
        TempDefs = Left(TempDefs, Len(TempDefs) - 1)
    End If
    End If
    
    Me.hide
End Sub

Private Sub Label_solvenum_Click()
   Dim F As String, Arr() As String
   Arr = Split(Label_ligning.Caption, "=")
   If UBound(Arr) > 0 Then F = Trim(Arr(1))
   If Len(Trim(Arr(0))) > 2 Then
      MsgBox "Differentialligningen skal være på formen y'=...  for at den kan løses numerisk" & vbCrLf & "", vbOKOnly, "Fejl"
      Exit Sub
   End If
   luk = True
   Me.hide
   UserFormDeSolveNumeric.TextBox_varx.text = TextBox_variabel.text
   UserFormDeSolveNumeric.TextBox_var1.text = TextBox_funktion.text
   UserFormDeSolveNumeric.TextBox_eq1.text = F
   If TextBox_starty.text = vbNullString Then
      UserFormDeSolveNumeric.TextBox_init1.text = "1"
   Else
      UserFormDeSolveNumeric.TextBox_init1.text = TextBox_starty.text
   End If
   UserFormDeSolveNumeric.TextBox_xmin.text = TextBox_startx.text
   UserFormDeSolveNumeric.TextBox_xmax.text = TextBox_startx.text + 10
   UserFormDeSolveNumeric.TextBox_definitioner.text = TextBox_def.text
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
    Label_ligning.Caption = Replace(Label_ligning.Caption, ChrW(180), "'") ' converttoascii indfører 180. Der står det er nødvendigt, men ved ikke hvorfor denne ser bedre ud

    If InStr(Label_ligning.Caption, ChrW(180) & ChrW(180)) > 0 Then ' "´´" to accenter der vender opad
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
            TextBox_def.text = TextBox_def.text & svar & "=" & VbCrLfMac    ' midlertidige definitioner
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
    fkt = TextBox_funktion.text
    Pos = InStr(fkt, "(")
    If Pos > 0 Then
        fkt = Left(fkt, Pos - 1)
    End If
    Label_diffy.Caption = fkt & "'(" & TextBox_startx.text & ")="
    Label_y.Caption = fkt & "("
    Label_y2.Caption = fkt & "("
End Sub

Sub SetCaptions()
    Me.Caption = Sprog.SolveDE
    Label1.Caption = Sprog.DifferentialEquation & ":"
    Label3.Caption = Sprog.IndepVar
    Label2.Caption = Sprog.DepVar
    Label4.Caption = Sprog.StartCond & ":"
    Label8.Caption = Sprog.A(297)
    Label_temp.Caption = Sprog.TempDefs
    Label_cancel.Caption = Sprog.Cancel
    Label_ok.Caption = Sprog.OK
    If Sprog.SprogNr = 1 Then
        Label_solvenum.Caption = "Løs numerisk"
    Else
        Label_solvenum.Caption = "Solve numerical"
    End If
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
