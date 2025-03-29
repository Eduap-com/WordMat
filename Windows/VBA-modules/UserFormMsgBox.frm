VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormMsgBox 
   Caption         =   "MsgBox"
   ClientHeight    =   3750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7260
   OleObjectBlob   =   "UserFormMsgBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This is a replacement for msgbox. It has the same design as the other Forms, and the size can be set automatically to fit the text

Public MsgBoxStyle As VbMsgBoxStyle
Public MsgBoxResult As VbMsgBoxResult
Public Title As String
Public prompt As String

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
    Label_ok_Click
End Sub

Private Sub Label_cancel_Click()
    If MsgBoxStyle = vbOKOnly Or MsgBoxStyle = vbOKCancel Then
        MsgBoxResult = vbCancel
    Else
        MsgBoxResult = vbNo
    End If
    On Error Resume Next
    Me.hide
End Sub

Private Sub Label_ok_Click()
    If MsgBoxStyle = vbOKOnly Or MsgBoxStyle = vbOKCancel Then
        MsgBoxResult = vbOK
    Else
        MsgBoxResult = vbYes
    End If
    Me.hide
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

Private Sub Label_prompt_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetButtonsInactive
End Sub

Private Sub UserForm_Initialize()
    SetEscEvents Me.Controls
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetButtonsInactive
End Sub

Sub SetButtonsInactive()
    Label_ok.BackColor = LBColorInactive
    Label_cancel.BackColor = LBColorInactive
End Sub

Private Sub UserForm_Activate()
    Dim h As Integer, w As Integer, Arr() As String, i As Integer
    
#If Mac Then
    prompt = Replace(prompt, vbCrLf, vbLf)
    prompt = Replace(prompt, vbCr, vbLf)
#Else
#End If
'    prompt = Replace(prompt, vbCr, vbLf)
    Arr = Split(prompt, vbLf)
'    h = 120 + 16 * GetCountOfChar(prompt, VbCrLfMac)
    h = 120 + 16 * UBound(Arr)
    If h > 1000 Then h = 1000
    
    For i = 0 To UBound(Arr)
        If Len(Arr(i)) > w Then w = Len(Arr(i))
    Next
        
    w = 100 + 5 * w
    If w < 200 Then w = 200
    If w > 800 Then w = 800
    
    Me.Height = h
    Me.Width = w
    Label_prompt.Width = w - 30
    Label_prompt.Height = h - 80
    

    If MsgBoxStyle = vbOKOnly Then
        Label_cancel.visible = False
        Label_ok.Left = Me.Width / 2 - Label_ok.Width / 2
    ElseIf MsgBoxStyle = vbOKCancel Then
        Label_cancel.visible = True
        Label_ok.Left = Me.Width / 2 - Label_ok.Width - 2
        Label_cancel.Left = Me.Width / 2 + 2
    ElseIf MsgBoxStyle = vbYesNo Then
        Label_ok.Caption = "Yes"
        Label_cancel.Caption = "No"
        Label_ok.visible = True
        Label_cancel.visible = True
        Label_ok.Left = Me.Width / 2 - Label_ok.Width - 2
        Label_cancel.Left = Me.Width / 2 + 2
    End If
    
    Label_ok.Top = Me.Height - 60
    Label_cancel.Top = Me.Height - 60
    
    Me.Caption = Title
    Label_prompt.Caption = prompt
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Label_cancel_Click
End Sub
