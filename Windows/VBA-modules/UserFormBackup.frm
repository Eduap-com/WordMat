VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormBackup 
   Caption         =   "Backup?"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   170
   ClientWidth     =   3450
   OleObjectBlob   =   "UserFormBackup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False










Public Backup As Boolean

Private Sub CommandButton_no_Click()
    Backup = False
    If CheckBox1.Value = True Then
        BackupType = 2
        MsgBox Sprog.A(207), vbOKOnly
    End If
    Unload Me
End Sub

Private Sub CommandButton_yes_Click()
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
    CommandButton_no.SetFocus
End Sub

Sub SetCaptions()
    Label_startbackup.Caption = Sprog.A(179)
    CommandButton_no.Caption = Sprog.A(204)
    CommandButton_yes.Caption = Sprog.A(203)
    Me.Caption = Sprog.A(170)
    CheckBox1.Caption = Sprog.A(205)
End Sub

