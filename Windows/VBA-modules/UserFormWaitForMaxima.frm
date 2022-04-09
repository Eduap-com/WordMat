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
'Public VarParam As String
'Public Param2 As String
Public StopNow As Boolean

Private Sub CommandButton_stop_Click()
On Error Resume Next
    omax.StopNow = True
    StopNow = True
'    omax.CloseCmd
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    StopNow = False
    SetCaptions
    DoEvents
End Sub

Private Sub UserForm_Initialize()
    Call RemoveCaption(Me)
#If Mac Then
    Me.Height = 180
#End If
End Sub

Private Sub SetCaptions()
    Me.Caption = Sprog.A(673)
    Label_tip.Caption = Sprog.A(674)
    Frame1.Caption = Sprog.Activity
    CommandButton_stop.Caption = Sprog.StopLabel
    Label1.Caption = Sprog.Wait
End Sub

