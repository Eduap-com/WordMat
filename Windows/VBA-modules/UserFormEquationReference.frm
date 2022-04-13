VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormEquationReference 
   Caption         =   "Ligningsreference"
   ClientHeight    =   5505
   ClientLeft      =   30
   ClientTop       =   165
   ClientWidth     =   4890
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
Private Sub CommandButton_cancel_Click()
    EqName = ""
    Me.Hide
End Sub

Private Sub CommandButton_ok_Click()
    EqName = ListBox1.text
    Me.Hide
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    EqName = ListBox1.text
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
    CommandButton_ok.Caption = Sprog.OK
    CommandButton_cancel.Caption = Sprog.Cancel
    Label_ligninger.Caption = Sprog.Equations
    
End Sub
