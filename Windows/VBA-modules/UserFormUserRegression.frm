VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormUserRegression 
   Caption         =   "Regression"
   ClientHeight    =   4500
   ClientLeft      =   30
   ClientTop       =   165
   ClientWidth     =   9195.001
   OleObjectBlob   =   "UserFormUserRegression.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormUserRegression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private EventsOn As Boolean

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
Private Sub CommandButton_cancel_Click()
    TextBox_function.Text = ""
    Me.hide
End Sub

Private Sub CommandButton_ok_Click()
    Me.hide
End Sub

Private Sub TextBox_eps_Change()
    UpdateWarning
End Sub

Sub UpdateWarning()
    Dim n As Long
    On Error GoTo Fejl
    If Not EventsOn Then Exit Sub
    EventsOn = False
    n = val(TextBox_eps.Text)
    If n > 0 And n <= 7 Then
        Label_epswarning.Caption = ""
'        Label_epswarning.visible = False
    ElseIf n > 7 And n <= 8 Then
        Label_epswarning.Caption = TT.A(44)
        Label_eps.visible = True
    ElseIf n > 8 Then
        Label_epswarning.Caption = TT.A(43)
        Label_eps.visible = True
    ElseIf n < 0 Then
        Label_epswarning.Caption = TT.A(45)
        Label_eps.visible = True
    Else
'        Label_epswarning.visible = False
    End If

GoTo slut
Fejl:
        Label_epswarning.Caption = TT.A(46)
        Label_eps.visible = True
slut:
    EventsOn = True
End Sub
Private Sub UserForm_Activate()
    EventsOn = True
    SetCaptions
    UpdateWarning
'    TextBox_eps.Text = MaximaCifre
End Sub

Sub SetCaptions()
    Me.Caption = TT.A(39)
    Label_function.Caption = TT.A(40)
    Label_var.Caption = TT.A(746)
    Label_eps.Caption = TT.A(41)
    Label_guess.Caption = TT.A(42)
    CommandButton_ok.Caption = TT.OK
    CommandButton_cancel.Caption = TT.Cancel
    Label_epswarning.Caption = TT.A(658)
End Sub

Private Sub UserForm_Initialize()
    SetEscEvents Me.Controls
End Sub
