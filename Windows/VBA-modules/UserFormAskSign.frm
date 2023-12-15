VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormAskSign 
   Caption         =   "Fortegn?"
   ClientHeight    =   3420
   ClientLeft      =   -30
   ClientTop       =   80
   ClientWidth     =   5100
   OleObjectBlob   =   "UserFormAskSign.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormAskSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit
Private Sub CommandButton_ok_Click()
    Me.hide
End Sub

Private Sub UserForm_Activate()
    Me.Caption = Sprog.Sign
    Label1.Caption = Sprog.SignDepends
    Label_statesign.Caption = Sprog.StateSign
    OptionButton_positiv.Caption = "+ " & Sprog.Positive
    OptionButton_nul.Caption = "0 " & Sprog.Zero
    OptionButton_negativ.Caption = "- " & Sprog.Negative
End Sub

