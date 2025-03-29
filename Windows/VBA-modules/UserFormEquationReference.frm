VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormEquationReference 
   Caption         =   "Ligningsreference"
   ClientHeight    =   6000
   ClientLeft      =   30
   ClientTop       =   165
   ClientWidth     =   4170
   OleObjectBlob   =   "UserFormEquationReference.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormEquationReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This form is used to insert a reference to a named equation

Public EqName As String
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
    EqName = ""
    Me.hide
End Sub

Private Sub Label_ok_Click()
    EqName = ListBox1.Text
    Me.hide
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    EqName = ListBox1.Text
    Me.hide
End Sub

Private Sub UserForm_Activate()
    Dim i As Integer
    SetCaptions
    EqName = ""
    ListBox1.Clear
    For i = 1 To ActiveDocument.Bookmarks.Count
        ListBox1.AddItem ActiveDocument.Bookmarks(i).Name
    Next
End Sub

Sub SetCaptions()
    Me.Caption = Sprog.A(14)
    Label_ok.Caption = Sprog.OK
    Label_cancel.Caption = Sprog.Cancel
    Label_Ligninger.Caption = Sprog.A(834)
End Sub

Private Sub Label_cancel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_cancel.BackColor = LBColorPress
End Sub
Private Sub Label_cancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_cancel.BackColor = LBColorHover
End Sub
Private Sub Label_ok_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_ok.BackColor = LBColorPress
End Sub
Private Sub Label_ok_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_ok.BackColor = LBColorHover
End Sub

Private Sub UserForm_Initialize()
    SetEscEvents Me.Controls
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then 'esc
        Me.hide
    End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_ok.BackColor = LBColorInactive
    Label_cancel.BackColor = LBColorInactive
End Sub
