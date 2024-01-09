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

Public ChosenCAS As Integer

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
If CASengine = 0 Then
    Label_MaximaSym.Enabled = False
'    Label_GeoGebraSym.SetFocus
ElseIf CASengine = 1 Then
    Label_GeoGebraBrowser.Enabled = False
'    Label_MaximaSym.SetFocus
ElseIf CASengine = 2 Then
    Label_GeoGebraSym.Enabled = False
'    Label_MaximaSym.SetFocus
End If
    SetCaptions
End Sub

Sub SetCaptions()
    Label_MaximaSym.Caption = "Maxima" & VbCrLfMac & Sprog.Exact
    Label_MaximaNum.Caption = "Maxima" & VbCrLfMac & Sprog.Numeric
    Label_GeoGebraSym.Caption = "GeoGebra" & VbCrLfMac & Sprog.Exact
    Label_GeoGebraNum.Caption = "GeoGebra" & VbCrLfMac & Sprog.Numeric
    Label_GeoGebraBrowser.Caption = "GeoGebra browser" & VbCrLfMac & Sprog.Exact
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
