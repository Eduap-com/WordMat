VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' assign all controls on forms where esc should be able to close

Public WithEvents LabelControl As MSForms.Label
Attribute LabelControl.VB_VarHelpID = -1
Public WithEvents CheckBoxControl As MSForms.CheckBox
Attribute CheckBoxControl.VB_VarHelpID = -1
Public WithEvents OptionButtonControl As MSForms.OptionButton
Attribute OptionButtonControl.VB_VarHelpID = -1
Public WithEvents ComboBoxControl As MSForms.ComboBox
Attribute ComboBoxControl.VB_VarHelpID = -1
Public WithEvents CommandButtonControl As MSForms.CommandButton
Attribute CommandButtonControl.VB_VarHelpID = -1
Public WithEvents TextBoxControl As MSForms.TextBox
Attribute TextBoxControl.VB_VarHelpID = -1
Public WithEvents ListBoxControl As MSForms.ListBox
Attribute ListBoxControl.VB_VarHelpID = -1
Public WithEvents ImageControl As MSForms.Image
Attribute ImageControl.VB_VarHelpID = -1
Public WithEvents MultiPageControl As MSForms.MultiPage
Attribute MultiPageControl.VB_VarHelpID = -1

Private Sub CheckBoxControl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CloseForm CheckBoxControl, KeyCode
End Sub
Private Sub OptionButtonControl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CloseForm OptionButtonControl, KeyCode
End Sub
Private Sub ComboBoxControl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CloseForm ComboBoxControl, KeyCode
End Sub
Private Sub LabelControl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CloseForm LabelControl, KeyCode
End Sub
Private Sub TextBoxControl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CloseForm TextBoxControl, KeyCode
End Sub
Private Sub CommandButtonControl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CloseForm CommandButtonControl, KeyCode
End Sub
Private Sub ListBoxControl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CloseForm ListBoxControl, KeyCode
End Sub
Private Sub ImageControl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CloseForm ImageControl, KeyCode
End Sub
Private Sub MultiPageControl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CloseForm MultiPageControl, KeyCode
End Sub

Sub CloseForm(Ctrl As MSForms.control, ByVal KeyCode As Integer)
On Error GoTo slut
    If KeyCode = 27 Then
        If Left(TypeName(Ctrl.Parent), 8) = "UserForm" Then
            Ctrl.Parent.hide
        ElseIf Left(TypeName(Ctrl.Parent.Parent), 8) = "UserForm" Then
            Ctrl.Parent.Parent.hide
        ElseIf Left(TypeName(Ctrl.Parent.Parent.Parent), 8) = "UserForm" Then
            Ctrl.Parent.Parent.Parent.hide
        ElseIf Left(TypeName(Ctrl.Parent.Parent.Parent.Parent), 8) = "UserForm" Then
            Ctrl.Parent.Parent.Parent.Parent.hide
        ElseIf Left(TypeName(Ctrl.Parent.Parent.Parent.Parent.Parent), 8) = "UserForm" Then
            Ctrl.Parent.Parent.Parent.Parent.Parent.hide
        End If
    End If
slut:
End Sub

