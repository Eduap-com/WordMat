VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormQuick 
   Caption         =   "Quick"
   ClientHeight    =   705
   ClientLeft      =   -15
   ClientTop       =   675
   ClientWidth     =   2370
   OleObjectBlob   =   "UserFormQuick.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "UserFormQuick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim start
Private Sub UserForm_Activate()
    start = Timer    ' Set start time.
    Do While Timer < start + 2
        DoEvents    ' Yield to other processes.
    Loop
    Me.Hide

End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    Call RemoveCaption(Me)

End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If (KeyCode.Value = 18 Or KeyCode.Value = 78) And Shift = 4 Then ' alt+n
On Error GoTo Slut
    If KeyCode.Value = 78 And Shift = 4 Then ' alt+n
    If MaximaExact = 0 Then
        Me.Label_text.Caption = Sprog.Exact ' "Eksakt"
        DoEvents
        MaximaExact = 1
        start = Timer    ' Set start time.
    ElseIf MaximaExact = 1 Then
        Me.Label_text.Caption = Sprog.Numeric ' "Num"
        DoEvents
        MaximaExact = 2
        start = Timer    ' Set start time.
    Else
        Me.Label_text.Caption = Sprog.Auto ' "Auto"
        DoEvents
        MaximaExact = 0
        start = Timer    ' Set start time.
    End If
    Else
        Me.Hide
    End If
    If Not (WoMatRibbon Is Nothing) Then
        WoMatRibbon.Invalidate
    End If
Slut:
End Sub
