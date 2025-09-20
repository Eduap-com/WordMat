VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormShortcuts 
   Caption         =   "Genveje"
   ClientHeight    =   9810.001
   ClientLeft      =   -30
   ClientTop       =   75
   ClientWidth     =   16380
   OleObjectBlob   =   "UserFormShortcuts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormShortcuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This form can be used to customize the keyboard shortcuts, and shows commonly used shortcuts for math symbols

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

Private Sub CommandButton_ok_Click()
    Label_ok_Click
End Sub

Private Sub Label_nulstil_Click()
    SettShortcutAltM = KeybShortcut.InsertNewEquation
    SettShortcutAltM2 = KeybShortcut.NoShortcut
    SettShortcutAltB = KeybShortcut.beregnudtryk
    SettShortcutAltL = KeybShortcut.SolveEquation
    SettShortcutAltP = KeybShortcut.ShowGraph
    SettShortcutAltD = KeybShortcut.Define
    SettShortcutAltS = KeybShortcut.sletdef
    SettShortcutAltF = KeybShortcut.Formelsamling
    SettShortcutAltO = KeybShortcut.OmskrivUdtryk
    SettShortcutAltR = KeybShortcut.PrevResult
    SettShortcutAltJ = KeybShortcut.SettingsForm
    SettShortcutAltN = KeybShortcut.NoShortcut
    SettShortcutAltE = KeybShortcut.NoShortcut
    SettShortcutAltT = KeybShortcut.ConvertEquationToLatex
    SettShortcutAltQ = KeybShortcut.GradTegn
    SettShortcutAltG = KeybShortcut.NoShortcut
    SettShortcutAltGr = KeybShortcut.NoShortcut
    
    SetComboIndexs
End Sub

Private Sub UserForm_Activate()
    SetCaptions
    SetComboIndexs

End Sub

Sub SetComboIndexs()
    On Error Resume Next
    ComboBox_AltM.ListIndex = SettShortcutAltM
    ComboBox_AltM2.ListIndex = SettShortcutAltM2
    ComboBox_AltB.ListIndex = SettShortcutAltB
    ComboBox_AltL.ListIndex = SettShortcutAltL
    ComboBox_AltD.ListIndex = SettShortcutAltD
    ComboBox_AltS.ListIndex = SettShortcutAltS
    ComboBox_AltP.ListIndex = SettShortcutAltP
    ComboBox_AltF.ListIndex = SettShortcutAltF
    ComboBox_AltO.ListIndex = SettShortcutAltO
    ComboBox_AltR.ListIndex = SettShortcutAltR
    ComboBox_AltJ.ListIndex = SettShortcutAltJ
    ComboBox_AltN.ListIndex = SettShortcutAltN
    ComboBox_AltE.ListIndex = SettShortcutAltE
    ComboBox_AltT.ListIndex = SettShortcutAltT
    ComboBox_AltQ.ListIndex = SettShortcutAltQ
    ComboBox_AltG.ListIndex = SettShortcutAltG
    ComboBox_AltGr.ListIndex = SettShortcutAltGr
    
End Sub

Sub SetCaptions()
    Me.Caption = TT.A(814)
    Label1.Caption = TT.A(65)
    Label2.Caption = TT.A(66)
    Label3.Caption = TT.A(67)
    Label_cancel.Caption = TT.Cancel
    Label_nulstil.Caption = TT.Reset
End Sub

Private Sub UserForm_Initialize()
    SetEscEvents Me.Controls

#If Mac Then
    ReplaceCaption Label_AltM
    ReplaceCaption Label_AltM2
    ReplaceCaption Label_AltB
    ReplaceCaption Label_AltL
    ReplaceCaption Label_AltD
    ReplaceCaption Label_AltS
    ReplaceCaption Label_AltP
    ReplaceCaption Label_AltF
    ReplaceCaption Label_AltO
    ReplaceCaption Label_AltR
    ReplaceCaption Label_AltJ
    ReplaceCaption Label_AltN
    ReplaceCaption Label_AltE
    ReplaceCaption Label_AltT
    ReplaceCaption Label_AltQ
    ReplaceCaption Label_AltG
    Label_AltGr.Caption = "Opt+enter"
#End If
    FillAllComboboxes
End Sub
'#If Mac Then
Function ReplaceCaption(l As Label)
    l.Caption = Replace(l.Caption, "Alt", "Opt")
End Function
'#End If
Sub FillAllComboboxes()
    FillComboBoxShortcuts ComboBox_AltM
    FillComboBoxShortcuts ComboBox_AltM2
    FillComboBoxShortcuts ComboBox_AltB
    FillComboBoxShortcuts ComboBox_AltL
    FillComboBoxShortcuts ComboBox_AltD
    FillComboBoxShortcuts ComboBox_AltS
    FillComboBoxShortcuts ComboBox_AltP
    FillComboBoxShortcuts ComboBox_AltF
    FillComboBoxShortcuts ComboBox_AltO
    FillComboBoxShortcuts ComboBox_AltR
    FillComboBoxShortcuts ComboBox_AltJ
    FillComboBoxShortcuts ComboBox_AltN
    FillComboBoxShortcuts ComboBox_AltE
    FillComboBoxShortcuts ComboBox_AltT
    FillComboBoxShortcuts ComboBox_AltQ
    FillComboBoxShortcuts ComboBox_AltG
    FillComboBoxShortcuts ComboBox_AltGr
End Sub

Sub FillComboBoxShortcuts(CB As ComboBox)

CB.Clear

CB.AddItem ""
CB.AddItem TT.A(701) 'New equation
CB.AddItem TT.A(1) 'New numbered equation
CB.AddItem TT.A(446)
CB.AddItem TT.A(760)
CB.AddItem TT.A(62) ' define
CB.AddItem TT.A(453)
CB.AddItem TT.A(461)
CB.AddItem TT.A(68)
CB.AddItem TT.A(456)
CB.AddItem TT.A(452)
CB.AddItem TT.A(505) ' Maxima command
CB.AddItem TT.A(702) 'prev resultat
CB.AddItem TT.A(443)
CB.AddItem TT.A(703) 'Toggle num/exact
CB.AddItem TT.A(262)
CB.AddItem TT.A(704) ' Convert latex
CB.AddItem TT.A(705) ' Latex pdf
CB.AddItem TT.A(607) ' reference to equation
CB.AddItem TT.A(706)
CB.AddItem TT.A(463)

End Sub
Private Sub Label_cancel_Click()
    Me.hide
End Sub
Private Sub Label_cancel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorPress
End Sub
Private Sub Label_cancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorHover
End Sub
Private Sub Label_ok_Click()
    SettShortcutAltM = ComboBox_AltM.ListIndex
    SettShortcutAltM2 = ComboBox_AltM2.ListIndex
    SettShortcutAltB = ComboBox_AltB.ListIndex
    SettShortcutAltL = ComboBox_AltL.ListIndex
    SettShortcutAltD = ComboBox_AltD.ListIndex
    SettShortcutAltS = ComboBox_AltS.ListIndex
    SettShortcutAltP = ComboBox_AltP.ListIndex
    SettShortcutAltF = ComboBox_AltF.ListIndex
    SettShortcutAltO = ComboBox_AltO.ListIndex
    SettShortcutAltR = ComboBox_AltR.ListIndex
    SettShortcutAltJ = ComboBox_AltJ.ListIndex
    SettShortcutAltN = ComboBox_AltN.ListIndex
    SettShortcutAltE = ComboBox_AltE.ListIndex
    SettShortcutAltT = ComboBox_AltT.ListIndex
    SettShortcutAltQ = ComboBox_AltQ.ListIndex
    SettShortcutAltG = ComboBox_AltG.ListIndex
    SettShortcutAltGr = ComboBox_AltGr.ListIndex
    
    Me.hide
End Sub
Private Sub Label_ok_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorPress
End Sub
Private Sub Label_ok_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorHover
End Sub
Private Sub Label_nulstil_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_nulstil.BackColor = LBColorPress
End Sub
Private Sub Label_nulstil_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_nulstil.BackColor = LBColorHover
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorInactive
    Label_cancel.BackColor = LBColorInactive
    Label_nulstil.BackColor = LBColorInactive
End Sub
