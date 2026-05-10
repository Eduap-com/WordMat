VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormWaitForMaxima 
   Caption         =   "Regner"
   ClientHeight    =   2610
   ClientLeft      =   -30
   ClientTop       =   675
   ClientWidth     =   3345
   OleObjectBlob   =   "UserFormWaitForMaxima.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "UserFormWaitForMaxima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StopNow As Boolean

Private Sub UserForm_Activate()
    StopNow = False
    SetCaptions
    DoEvents
End Sub

Private Sub SetCaptions()
    Me.Caption = TT.A(673)
    Label_tip.Caption = TT.A(674)
    Frame1.Caption = TT.A(827)
    Label_stop.Caption = TT.A(828)
    Label1.Caption = TT.A(826)
End Sub

Private Sub label_stop_Click()
On Error Resume Next
    omax.StopNow = True
    StopNow = True
    Me.hide
End Sub

Private Sub label_stop_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_stop.BackColor = LBColorPress
End Sub

Private Sub label_stop_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_stop.BackColor = LBColorHover
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_stop.BackColor = LBColorInactive
End Sub
