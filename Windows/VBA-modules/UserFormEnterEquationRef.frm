VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormEnterEquationRef 
   Caption         =   "Indtast navn på ligning"
   ClientHeight    =   5535
   ClientLeft      =   30
   ClientTop       =   165
   ClientWidth     =   6675
   OleObjectBlob   =   "UserFormEnterEquationRef.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormEnterEquationRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public EquationName As String

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
    EquationName = ""
    Me.hide
End Sub

Private Sub CommandButton_ok_Click()
Dim i As Integer
    EquationName = Trim(TextBox1.text)
    If InStr(EquationName, " ") > 0 Then
        EquationName = ""
        Label_error.visible = True
        Label_error.Caption = Sprog.A(13)
        TextBox1.SetFocus
        Exit Sub
    End If
For i = 1 To ActiveDocument.Bookmarks.Count
    If ActiveDocument.Bookmarks(i).Name = EquationName Then
        EquationName = ""
        Label_error.visible = True
        Label_error.Caption = Sprog.A(12)
        TextBox1.SetFocus
        Exit Sub
    End If
Next
    
    Me.hide
End Sub

Private Sub Label_ok_Click()
    CommandButton_ok_Click
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    TextBox1.text = ListBox1.text
    TextBox1.SetFocus
End Sub

Private Sub UserForm_Activate()
Dim i As Integer
    On Error GoTo Fejl
    SetCaptions
    EquationName = ""
    Label_error.visible = False

ListBox1.Clear
For i = 1 To ActiveDocument.Bookmarks.Count
    ListBox1.AddItem ActiveDocument.Bookmarks(i).Name
Next
TextBox1.SetFocus

Fejl:
End Sub

Private Sub UserForm_Initialize()
    SetEscEvents Me.Controls
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    EquationName = ""
End Sub

Sub SetCaptions()
    Me.Caption = Sprog.A(5)
    Label_name.Caption = Sprog.A(5)
    Label_cancel.Caption = Sprog.Cancel
    Label_ok.Caption = Sprog.OK
    Label_Ligninger.Caption = Sprog.A(10)
    Label_help.Caption = Sprog.A(11)
    Label_error.Caption = Sprog.A(12)
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
End Sub
