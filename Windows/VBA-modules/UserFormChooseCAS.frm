VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormChooseCAS 
   Caption         =   "Choose CAS"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7515
   OleObjectBlob   =   "UserFormChooseCAS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormChooseCAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This form is used when an equation cannot be solved, and the user must choose an alternative CAS-engine
' It is somewhat deprecated, since version 1.33, when equation solving automatically goes to numerical solver
Public ChosenCAS As Integer
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
    ChosenCAS = -1
    Me.hide
End Sub

Private Sub Label_GeoGebraBrowser_Click()
    ChosenCAS = 5
    Me.hide
End Sub

Private Sub Label_GeoGebraGrafisk_Click()
    ChosenCAS = 6
    Me.hide
End Sub

Private Sub Label_GeoGebraNum_Click()
    ChosenCAS = 4
    Me.hide
End Sub

Private Sub Label_GeoGebraSym_Click()
    ChosenCAS = 3
    Me.hide
End Sub

Private Sub Label_maximaNum_Click()
    ChosenCAS = 2
    Me.hide
End Sub

Private Sub Label_MaximaSym_Click()
    ChosenCAS = 1
    Me.hide
End Sub

Private Sub UserForm_Activate()
Label_MaximaSym.Enabled = True
Label_GeoGebraSym.Enabled = True
Label_GeoGebraBrowser.Enabled = True
SetAllInactive
If CASengine = 0 Then
    Label_MaximaSym.Enabled = False
ElseIf CASengine = 1 Then
    Label_GeoGebraBrowser.Enabled = False
ElseIf CASengine = 2 Then
    Label_GeoGebraSym.Enabled = False
End If
    SetCaptions
End Sub

Sub SetCaptions()
    Label_MaximaSym.Caption = "Maxima" & VbCrLfMac & Sprog.A(710)
    Label_MaximaNum.Caption = "Maxima" & VbCrLfMac & Sprog.A(711)
    Label_GeoGebraSym.Caption = "GeoGebra" & VbCrLfMac & Sprog.A(710)
    Label_GeoGebraNum.Caption = "GeoGebra" & VbCrLfMac & Sprog.A(711)
    Label_GeoGebraBrowser.Caption = "GeoGebra browser" & VbCrLfMac & Sprog.A(710)
    Label_GeoGebraGrafisk.Caption = "GeoGebra" & VbCrLfMac & Sprog.A(241)
    Label_text.Caption = Sprog.A(685)
End Sub

Private Sub Label_cancel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorPress
End Sub
Private Sub Label_cancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetAllInactive
    Label_cancel.BackColor = LBColorHover
End Sub
Private Sub Label_MaximaSym_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_MaximaSym.BackColor = LBColorPress
End Sub
Private Sub Label_MaximaSym_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetAllInactive
    Label_MaximaSym.BackColor = LBColorHover
End Sub
Private Sub Label_MaximaNum_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetAllInactive
    Label_MaximaNum.BackColor = LBColorHover
End Sub
Private Sub Label_MaximaNum_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_MaximaNum.BackColor = LBColorPress
End Sub
Private Sub Label_GeoGebraSym_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetAllInactive
    Label_GeoGebraSym.BackColor = LBColorHover
End Sub
Private Sub Label_GeoGebraSym_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_GeoGebraSym.BackColor = LBColorPress
End Sub
Private Sub Label_GeoGebraNum_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetAllInactive
    Label_GeoGebraNum.BackColor = LBColorHover
End Sub
Private Sub Label_GeoGebraNum_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_GeoGebraNum.BackColor = LBColorPress
End Sub
Private Sub Label_GeoGebraBrowser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetAllInactive
    Label_GeoGebraBrowser.BackColor = LBColorHover
End Sub
Private Sub Label_GeoGebraBrowser_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_GeoGebraBrowser.BackColor = LBColorPress
End Sub
Private Sub Label_GeoGebraGrafisk_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetAllInactive
    Label_GeoGebraGrafisk.BackColor = LBColorHover
End Sub
Private Sub Label_GeoGebraGrafisk_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_GeoGebraGrafisk.BackColor = LBColorPress
End Sub

Private Sub UserForm_Initialize()
    SetEscEvents Me.Controls
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetAllInactive
End Sub

Sub SetAllInactive()
    Label_cancel.BackColor = LBColorInactive
    Label_MaximaSym.BackColor = LBColorInactive
    Label_MaximaNum.BackColor = LBColorInactive
    Label_GeoGebraSym.BackColor = LBColorInactive
    Label_GeoGebraNum.BackColor = LBColorInactive
    Label_GeoGebraBrowser.BackColor = LBColorInactive
    Label_GeoGebraGrafisk.BackColor = LBColorInactive
End Sub
