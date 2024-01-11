VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormShowDef 
   Caption         =   "Definitions"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9330.001
   OleObjectBlob   =   "UserFormShowDef.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormShowDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_ok_Click()
    Me.hide
End Sub

Private Sub Label_ok_Click()
    Me.hide
End Sub

Private Sub UserForm_Activate()
    Me.Caption = Sprog.Definitions
    Label_text.Caption = Sprog.A(113)
    Label_defs.Caption = FormatDefinitions(omax.DefString)
End Sub

Private Sub Label_ok_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorPress
End Sub
Private Sub Label_ok_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorHover
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorInactive
End Sub

