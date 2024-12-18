VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormUserRegression 
   Caption         =   "Regression"
   ClientHeight    =   3135
   ClientLeft      =   30
   ClientTop       =   165
   ClientWidth     =   9345.001
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
    TextBox_function.text = ""
    Me.hide
End Sub

Private Sub CommandButton_ok_Click()
    Me.hide
End Sub

Private Sub TextBox_eps_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim n As Long
    On Error GoTo Fejl
    If Not EventsOn Then Exit Sub
    EventsOn = False
    n = val(TextBox_eps.text)
    If n > 0 And n < 5 Then
        Label_epswarning.Caption = ""
'        Label_epswarning.visible = False
    ElseIf n > 4 And n < 7 Then
        Label_epswarning.Caption = Sprog.A(44)
        Label_eps.visible = True
    ElseIf n > 6 Then
        Label_epswarning.Caption = Sprog.A(43)
        Label_eps.visible = True
    ElseIf n < 0 Then
        Label_epswarning.Caption = Sprog.A(45)
        Label_eps.visible = True
    Else
'        Label_epswarning.visible = False
    End If

GoTo Slut
Fejl:
        Label_epswarning.Caption = Sprog.A(46)
        Label_eps.visible = True
Slut:
    EventsOn = True

End Sub

Private Sub UserForm_Activate()
EventsOn = True
    SetCaptions
End Sub

Sub SetCaptions()
    Me.Caption = Sprog.A(39)
    Label_function.Caption = Sprog.A(40)
    Label_var.Caption = Sprog.IndepVar
    Label_eps.Caption = Sprog.A(41)
    Label_guess.Caption = Sprog.A(42)
    CommandButton_ok.Caption = Sprog.OK
    CommandButton_cancel.Caption = Sprog.Cancel
    Label_epswarning.Caption = Sprog.A(658)
    
End Sub

Private Sub UserForm_Initialize()
    SetEscEvents Me.Controls
End Sub
