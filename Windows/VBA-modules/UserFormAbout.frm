VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormAbout 
   Caption         =   "About"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8355.001
   OleObjectBlob   =   "UserFormAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This form shows who made WordMat, version number, partnership info etc.

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
        ElseIf TN = "Image" Then
            Set CE = New CEvents: Set CE.ImageControl = c: EventsCol.Add CE
        ElseIf TN = "Frame" Then
            Set F = c
            SetEscEvents F.Controls
        ElseIf TN = "UserForm" Then
            Set CE = New CEvents: Set CE.ImageControl = c: EventsCol.Add CE
            Set F = c
            SetEscEvents F.Controls
        End If
    Next
End Sub

Private Sub Image1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    OpenLink "https://www.eduap.com"
End Sub

Private Sub Label_checkpartnerskab_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_checkpartnerskab.BackColor = LBColorHover
End Sub

Private Sub Label_license_Click()
    OpenLink "https://www.eduap.com/wordmath-license/"
End Sub

Private Sub Label_Licensetitle_Click()
    OpenLink "https://www.eduap.com/wordmath-license/"
End Sub

Private Sub Label_link_Click()
    OpenLink "https://www.eduap.com"
End Sub

Private Sub Label_logo_Click()
    OpenLink "https://www.eduap.com"
End Sub

Private Sub Label_ok_Click()
    Me.hide
End Sub

Private Sub Label_partnerskab_Click()
    OpenLink "https://www.eduap.com/da/partnerskab/"
End Sub

Private Sub Label_partnerskab2_Click()
    OpenLink "https://www.eduap.com/da/partnerskab/"
End Sub

Private Sub Label1_Click()
    OpenLink "https://www.eduap.com"
End Sub

Private Sub UserForm_Activate()
    SetCaptions
    
    Label_partnerskab2.Caption = TT.A(161)
    If QActivePartnership Then
        Label_partnerskab.Caption = TT.A(136) & " " & SkoleNavn
        Label_checkpartnerskab.visible = False
    Else
        Label_partnerskab.Caption = TT.A(140)
        Label_checkpartnerskab.visible = True
    End If
End Sub

Sub SetCaptions()
    Dim v As String

    v = AppVersion
    If PatchVersion <> "" Then
        v = v & PatchVersion
    End If
    If TT.LangNo = 1 Then
        Me.Caption = TT.A(812) & " WordMat"
        Label_title.Caption = "WordMat"
    Else
        Me.Caption = TT.A(812) & " WordMath"
        Label_title.Caption = "WordMath"
    End If
    Label_version.Caption = "Version: " & v
    Label_text.Caption = TT.A(20)

End Sub

Private Sub Label_ok_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorPress
End Sub
Private Sub Label_ok_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorHover
End Sub

Private Sub UserForm_Initialize()
    SetEscEvents Me.Controls
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then 'esc
        Me.hide
    End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorInactive
    Label_checkpartnerskab.BackColor = LBColorInactive
End Sub
Private Sub Label_checkpartnerskab_Click()
    If QActivePartnership(True) Then
        MsgBox TT.A(120), vbOKOnly, "Success!"
        UserForm_Activate
    Else
        MsgBox TT.A(121), vbOKOnly, TT.A(96)
    End If
End Sub
