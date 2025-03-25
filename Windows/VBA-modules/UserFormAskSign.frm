VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormAskSign 
   Caption         =   "Fortegn?"
   ClientHeight    =   3540
   ClientLeft      =   -30
   ClientTop       =   75
   ClientWidth     =   5145
   OleObjectBlob   =   "UserFormAskSign.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormAskSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
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
Private Sub CommandButton_ok_Click()
    Me.hide
End Sub

Private Sub UserForm_Activate()
    Me.Caption = Sprog.A(818)
    Label1.Caption = Sprog.A(742)
    Label_statesign.Caption = Sprog.A(743)
    OptionButton_positiv.Caption = "+ " & Sprog.A(819)
    OptionButton_nul.Caption = "0 " & Sprog.A(821)
    OptionButton_negativ.Caption = "- " & Sprog.A(820)
End Sub

Private Sub UserForm_Initialize()
    SetEscEvents Me.Controls
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then 'esc
        Me.hide
    End If
End Sub
