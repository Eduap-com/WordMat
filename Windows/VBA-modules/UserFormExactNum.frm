VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormExactNum 
   ClientHeight    =   1215
   ClientLeft      =   30
   ClientTop       =   165
   ClientWidth     =   2040
   OleObjectBlob   =   "UserFormExactNum.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormExactNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim start
Private Sub UserForm_Activate()
    start = Timer    ' Set start time.
    Do While Timer < start + 1
        DoEvents    ' Yield to other processes.
    Loop
On Error Resume Next
    Me.Hide
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    Call RemoveCaption(Me)
    Me.Caption = ""
    Label_auto.Caption = Sprog.Auto
    Label_exact.Caption = Sprog.Exact
    Label_num.Caption = Sprog.Numeric
#If Mac Then
    Me.Height = 100
#End If

End Sub
Public Sub SetAuto()
    Label_auto.BorderStyle = fmBorderStyleSingle
    Label_exact.BorderStyle = fmBorderStyleNone
    Label_num.BorderStyle = fmBorderStyleNone
    Label_auto.Font.Bold = True
    Label_exact.Font.Bold = False
    Label_num.Font.Bold = False
End Sub
Public Sub SetExact()
    Label_auto.BorderStyle = fmBorderStyleNone
    Label_exact.BorderStyle = fmBorderStyleSingle
    Label_num.BorderStyle = fmBorderStyleNone
    Label_auto.Font.Bold = False
    Label_exact.Font.Bold = True
    Label_num.Font.Bold = False
End Sub
Public Sub SetNum()
    Label_auto.BorderStyle = fmBorderStyleNone
    Label_exact.BorderStyle = fmBorderStyleNone
    Label_num.BorderStyle = fmBorderStyleSingle
    Label_auto.Font.Bold = False
    Label_exact.Font.Bold = False
    Label_num.Font.Bold = True
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If (KeyCode.Value = 18 Or KeyCode.Value = 78) And Shift = 4 Then ' alt+n
On Error GoTo slut
    If KeyCode.Value = 78 And Shift = 4 Then ' alt+n
    If MaximaExact = 0 Then
        SetExact
        DoEvents
        MaximaExact = 1
        start = Timer    ' Set start time.
    ElseIf MaximaExact = 1 Then
        SetNum
        DoEvents
        MaximaExact = 2
        start = Timer    ' Set start time.
    Else
        SetAuto
        DoEvents
        MaximaExact = 0
        start = Timer    ' Set start time.
    End If
    Else
        Me.Hide
    End If
    WoMatRibbon.Invalidate
slut:
End Sub

