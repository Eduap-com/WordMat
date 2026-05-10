VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormGeoGebra 
   Caption         =   "GeoGebra"
   ClientHeight    =   3510
   ClientLeft      =   -15
   ClientTop       =   75
   ClientWidth     =   8775.001
   OleObjectBlob   =   "UserFormGeoGebra.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormGeoGebra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ReturnVal As Integer ' 1=Install, 2=browser
Private EventsCol As New Collection
Sub SetEscEvents(ControlColl As Controls)
' SetEscEvents Me.Controls     in Initialize
    Dim CE As CEvents, c As control, TN As String, F As MSForms.Frame
    On Error Resume Next
    For Each c In ControlColl ' Me.Controls
        TN = TypeName(c)
        If TN = "CheckBox" Then
            Set CE = New CEvents: Set CE.CheckBoxControl = c: EventsCol.Add CE
        ElseIf TN = "OptionButton" Then
            Set CE = New CEvents: Set CE.OptionButtonControl = c: EventsCol.Add CE
        ElseIf TN = "ComboBox" Then
            Set CE = New CEvents: Set CE.ComboBoxControl = c: EventsCol.Add CE
        ElseIf TN = "Label" Then
            Set CE = New CEvents: Set CE.LabelControl = c: EventsCol.Add CE
        ElseIf TN = "TextBox" Then
            Set CE = New CEvents: Set CE.TextBoxControl = c: EventsCol.Add CE
        ElseIf TN = "CommandButton" Then
            Set CE = New CEvents: Set CE.CommandButtonControl = c: EventsCol.Add CE
        ElseIf TN = "ListBox" Then
            Set CE = New CEvents: Set CE.ListBoxControl = c: EventsCol.Add CE
        ElseIf TN = "Frame" Then
            Set F = c
            SetEscEvents F.Controls
        End If
    Next
End Sub
Private Sub Label_Installer_Click()
    ReturnVal = 1
    Me.hide
End Sub

Private Sub CommandButton_webstart_Click()
    ReturnVal = 2
    Me.hide
End Sub


Private Sub Label_webstart_Click()
    CommandButton_webstart_Click
End Sub

Private Sub UserForm_Activate()
#If Mac Then
#Else
#End If
    
    Label_title.Caption = TT.A(292)
#If Mac Then
        Label2.Caption = TT.A(848)
#Else
        Label2.Caption = TT.A(293)
#End If
    
    Label3.Caption = TT.A(294)
    Label_webstart.Caption = TT.A(296)
    
#If Mac Then
    Label_Installer.Caption = TT.A(295) & " 5"
#Else
    Label_Installer.Caption = TT.A(295)
#End If
    
    ReturnVal = 0
End Sub

Private Sub Label_Installer_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_Installer.BackColor = LBColorPress
End Sub

Private Sub Label_Installer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_Installer.BackColor = LBColorHover
End Sub

Private Sub Label_webstart_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_webstart.BackColor = LBColorPress
End Sub

Private Sub Label_webstart_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_webstart.BackColor = LBColorHover
End Sub

Private Sub UserForm_Initialize()
    SetEscEvents Me.Controls
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_Installer.BackColor = LBColorInactive
    Label_webstart.BackColor = LBColorInactive
End Sub
