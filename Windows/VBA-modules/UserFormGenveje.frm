VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormGenveje 
   Caption         =   "Genveje"
   ClientHeight    =   13185
   ClientLeft      =   -30
   ClientTop       =   75
   ClientWidth     =   10845
   OleObjectBlob   =   "UserFormGenveje.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormGenveje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub UserForm_Activate()
    SetCaptions
End Sub

Sub SetCaptions()
    Me.Caption = Sprog.Shortcuts
    Label1.Caption = Sprog.A(65)
    Label2.Caption = Sprog.A(66)
    Label3.Caption = Sprog.A(67)
'#If Mac Then
'    TextBox1.Text = Replace(Sprog.A(68), "Alt", "ctrl")
'#Else
    TextBox1.Text = Sprog.A(68)
'#End If
End Sub

Private Sub UserForm_Initialize()
    SetEscEvents Me.Controls
End Sub
