VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormError 
   Caption         =   "Fejl"
   ClientHeight    =   7335
   ClientLeft      =   -30
   ClientTop       =   75
   ClientWidth     =   9360.001
   OleObjectBlob   =   "UserFormError.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_ok_Click()
    Label_ok_Click
End Sub

Private Sub Label_ok_Click()
'    Unload Me
    Label_TAB2.visible = True
    Me.Hide
End Sub

Private Sub Label_restart_Click()
' Denne funktion crasher Word. Det sker n�r Cmaxima lukkes. Den m� ikke k�res fra formen er fra Cmaxima. Den k�res nu fra Checkforerror n�r form lukkes
    RestartWordMat
    Unload Me
End Sub

Private Sub UserForm_Activate()
    SetCaptions
    Label_TAB1_Click
End Sub

Sub SetErrorDefinition(ED As ErrorDefinition)
    Label_titel.Caption = ED.Title
    Label_fejltekst.Caption = ED.Description
    If ED.DefFejl Then
        Label_definitioner.visible = True
        TextBox_definitioner.visible = True
        TextBox_definitioner.Text = FormatDefinitions(omax.DefString)
    Else
        Label_definitioner.visible = False
        TextBox_definitioner.visible = False
    End If
    If ED.MaximaOutput = vbNullString Then
        Label_TAB2.visible = False
    Else
        Label_maximaoutput.Caption = ED.MaximaOutput
        Label_TAB2.visible = True
    End If
End Sub

Private Sub SetCaptions()
    Me.Caption = Sprog.Error
    MultiPage1.Pages(0).Caption = Sprog.Error
    MultiPage1.Pages(1).Caption = Sprog.MaximaError
    Label_TAB1.Caption = Sprog.Error
    Label_TAB2.Caption = Sprog.MaximaError
    Label_definitioner.Caption = Sprog.Definitions
    Label_restart.Caption = Sprog.RestartWordMat
'MultiPage1.Pages("Page1").Caption
End Sub
Private Sub Label_restart_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_restart.BackColor = LBColorPress
End Sub
Private Sub Label_restart_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_restart.BackColor = LBColorHover
End Sub
Private Sub Label_ok_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorPress
End Sub
Private Sub Label_ok_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorHover
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorInactive
    Label_restart.BackColor = LBColorInactive
End Sub

Private Sub Label_TAB1_Click()
    MultiPage1.Value = 0
    SetTabsInactive
    Label_TAB1.BackColor = LBColorTABPress
End Sub
Private Sub Label_TAB2_Click()
    MultiPage1.Value = 1
    SetTabsInactive
    Label_TAB2.BackColor = LBColorTABPress
End Sub
Private Sub Label_TAB1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_TAB1.BackColor = LBColorPress
End Sub
Private Sub Label_TAB1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetTabsInactive
    If MultiPage1.Value <> 0 Then Label_TAB1.BackColor = LBColorHover
End Sub
Private Sub Label_TAB2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_TAB2.BackColor = LBColorPress
End Sub
Private Sub Label_TAB2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetTabsInactive
    If MultiPage1.Value <> 1 Then Label_TAB2.BackColor = LBColorHover
End Sub

Sub SetTabsInactive()
    If MultiPage1.Value <> 0 Then Label_TAB1.BackColor = LBColorInactive
    If MultiPage1.Value <> 1 Then Label_TAB2.BackColor = LBColorInactive
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
        Label_TAB2.visible = True
End Sub

Private Sub UserForm_Terminate()
    Label_TAB2.visible = True
End Sub
