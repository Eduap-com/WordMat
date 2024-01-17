VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormEquationReference 
   Caption         =   "Ligningsreference"
   ClientHeight    =   6000
   ClientLeft      =   30
   ClientTop       =   165
   ClientWidth     =   4170
   OleObjectBlob   =   "UserFormEquationReference.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormEquationReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public EqName As String
Private Sub Label_cancel_Click()
    EqName = ""
    Me.Hide
End Sub

Private Sub Label_ok_Click()
    EqName = ListBox1.Text
    Me.Hide
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    EqName = ListBox1.Text
    Me.Hide
End Sub

Private Sub UserForm_Activate()
Dim i As Integer

ListBox1.Clear
For i = 1 To ActiveDocument.Bookmarks.Count
    ListBox1.AddItem ActiveDocument.Bookmarks(i).Name
Next

End Sub

Sub SetCaptions()
    Me.Caption = Sprog.A(15)
    Label_ok.Caption = Sprog.OK
    Label_cancel.Caption = Sprog.Cancel
    Label_Ligninger.Caption = Sprog.Equations
    
End Sub

Private Sub Label_cancel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorPress
End Sub
Private Sub Label_cancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorHover
End Sub
Private Sub Label_ok_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorPress
End Sub
Private Sub Label_ok_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorHover
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorInactive
    Label_cancel.BackColor = LBColorInactive
End Sub
