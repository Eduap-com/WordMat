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
' Form used to show quick changes of settings. Units and num/exact

Private start As Single
Private Sub UserForm_Activate()
    start = timer    ' Set start time.
    Do While timer < start + 2
        DoEvents    ' Yield to other processes.
    Loop
    Me.hide
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.hide
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If (KeyCode.Value = 18 Or KeyCode.Value = 78) And Shift = 4 Then ' alt+n
On Error GoTo slut
    If KeyCode.Value = 78 And Shift = 4 Then ' alt+n
    If MaximaExact = 0 Then
        Me.Label_text.Caption = TT.A(710) ' "Eksakt"
        DoEvents
        MaximaExact = 1
        start = timer    ' Set start time.
    ElseIf MaximaExact = 1 Then
        Me.Label_text.Caption = TT.A(711) ' "Num"
        DoEvents
        MaximaExact = 2
        start = timer    ' Set start time.
    Else
        Me.Label_text.Caption = TT.A(712) ' "Auto"
        DoEvents
        MaximaExact = 0
        start = timer    ' Set start time.
    End If
    Else
        Me.hide
    End If
    If Not (WoMatRibbon Is Nothing) Then
        WoMatRibbon.Invalidate
    End If
slut:
End Sub
