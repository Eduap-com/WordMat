VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormError 
   Caption         =   "Fejl"
   ClientHeight    =   5865
   ClientLeft      =   -30
   ClientTop       =   80
   ClientWidth     =   7430
   OleObjectBlob   =   "UserFormError.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub CommandButton_ok_Click()
    Unload Me
End Sub

Private Sub CommandButton_restart_Click()
    RestartWordMat
    Unload Me
End Sub

Private Sub UserForm_Activate()
    SetCaptions
End Sub
Private Sub SetCaptions()
    Me.Caption = Sprog.Error
    MultiPage1.Pages(0).Caption = Sprog.Error
    MultiPage1.Pages(1).Caption = Sprog.MaximaError
    Label_definitioner.Caption = Sprog.Definitions
    CommandButton_restart.Caption = Sprog.RestartWordMat
'MultiPage1.Pages("Page1").Caption
End Sub
