VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormGenveje 
   Caption         =   "Genveje"
   ClientHeight    =   9490.001
   ClientLeft      =   -30
   ClientTop       =   75
   ClientWidth     =   9975.001
   OleObjectBlob   =   "UserFormGenveje.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormGenveje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub UserForm_Activate()
    SetCaptions
    GenerateKeyboardShortcuts
End Sub

Sub SetCaptions()
    Me.Caption = Sprog.Shortcuts
    Label1.Caption = Sprog.A(65)
    Label2.Caption = Sprog.A(66)
    Label3.Caption = Sprog.A(67)
'#If Mac Then
'    TextBox1.Text = Replace(Sprog.A(68), "Alt", "ctrl")
'#Else
    TextBox1.text = Sprog.A(68)
'#End If
End Sub
