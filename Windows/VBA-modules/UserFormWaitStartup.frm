VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormWaitStartup 
   ClientHeight    =   2055
   ClientLeft      =   -30
   ClientTop       =   80
   ClientWidth     =   4310
   OleObjectBlob   =   "UserFormWaitStartup.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserFormWaitStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit
Private Sub UserForm_Initialize()
    Call RemoveCaption(Me)

    SetCaptions
    Label_tip.Caption = GetRandomTip()
End Sub

Private Sub SetCaptions()
    Label1.Caption = Sprog.WordMatStarting
    Label2.Caption = Sprog.PleaseWait
End Sub
