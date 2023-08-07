VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormChooseCAS 
   Caption         =   "Choose CAS"
   ClientHeight    =   2760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5100
   OleObjectBlob   =   "UserFormChooseCAS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormChooseCAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public ChosenCAS As Integer
Private Sub CommandButton_cancel_Click()
    ChosenCAS = -1
    Me.Hide
End Sub

Private Sub CommandButton_GeoGebraBrowser_Click()
    ChosenCAS = 5
    Me.Hide
End Sub

Private Sub CommandButton_GeoGebraGrafisk_Click()
    ChosenCAS = 6
    Me.Hide
End Sub

Private Sub CommandButton_GeoGebraNum_Click()
    ChosenCAS = 4
    Me.Hide
End Sub

Private Sub CommandButton_GeoGebraSym_Click()
    ChosenCAS = 3
    Me.Hide
End Sub

Private Sub CommandButton_maximaNum_Click()
    ChosenCAS = 2
    Me.Hide
End Sub

Private Sub CommandButton_MaximaSym_Click()
    ChosenCAS = 1
    Me.Hide
End Sub

Private Sub UserForm_Activate()
CommandButton_MaximaSym.Enabled = True
CommandButton_GeoGebraSym.Enabled = True
CommandButton_GeoGebraBrowser.Enabled = True
If CASengine = 0 Then
    CommandButton_MaximaSym.Enabled = False
    CommandButton_GeoGebraSym.SetFocus
ElseIf CASengine = 1 Then
    CommandButton_GeoGebraBrowser.Enabled = False
    CommandButton_MaximaSym.SetFocus
ElseIf CASengine = 2 Then
    CommandButton_GeoGebraSym.Enabled = False
    CommandButton_MaximaSym.SetFocus
End If
    SetCaptions
End Sub

Sub SetCaptions()
    CommandButton_MaximaSym.Caption = "Maxima" & VbCrLfMac & Sprog.Exact
    CommandButton_MaximaNum.Caption = "Maxima" & VbCrLfMac & Sprog.Numeric
    CommandButton_GeoGebraSym.Caption = "GeoGebra" & VbCrLfMac & Sprog.Exact
    CommandButton_GeoGebraNum.Caption = "GeoGebra" & VbCrLfMac & Sprog.Numeric
    CommandButton_GeoGebraBrowser.Caption = "GeoGebra browser" & VbCrLfMac & Sprog.Exact
    CommandButton_GeoGebraGrafisk.Caption = "GeoGebra" & VbCrLfMac & Sprog.A(241)
    Label_text.Caption = Sprog.A(685)
End Sub
