VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormBackup 
   Caption         =   "Backup?"
   ClientHeight    =   2020
   ClientLeft      =   60
   ClientTop       =   170
   ClientWidth     =   3840
   OleObjectBlob   =   "UserFormBackup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Backup As Boolean
Private Sub CommandButton_no_Click()
    Backup = False
    If CheckBox1.Value = True Then
        BackupType = 2
        MsgBox Sprog.A(207), vbOKOnly
    End If
    Unload Me
End Sub

Private Sub Label_cancel_Click()
    CommandButton_no_Click
End Sub

Private Sub Label_ok_Click()
    Backup = True
    If CheckBox1.Value = True Then
        BackupType = 1
        MsgBox Sprog.A(207), vbOKOnly
    End If
    Unload Me
End Sub

Private Sub Label_punkterhelp_Click()
    MsgBox Sprog.A(206), vbOKOnly, Sprog.A(170)
End Sub

Private Sub UserForm_Activate()
    SetCaptions
'    CommandButton_no.SetFocus
End Sub

Sub SetCaptions()
    Label_startbackup.Caption = Sprog.A(179)
    Label_cancel.Caption = Sprog.A(204)
    Label_ok.Caption = Sprog.A(203)
    Me.Caption = Sprog.A(170)
    CheckBox1.Caption = Sprog.A(205)
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
