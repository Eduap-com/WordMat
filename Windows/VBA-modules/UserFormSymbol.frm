VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormSymbol 
   Caption         =   "Symbol"
   ClientHeight    =   4875
   ClientLeft      =   -15
   ClientTop       =   30
   ClientWidth     =   4710
   OleObjectBlob   =   "UserFormSymbol.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormSymbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public tegn As String
Private Sub CommandButton1_Click()
On Error GoTo Fejl

If OptionButton1.Value = True Then
  tegn = VBA.ChrW(916) ' Delta
ElseIf OptionButton2.Value = True Then
  tegn = VBA.ChrW(937)
ElseIf OptionButton3.Value = True Then
  tegn = VBA.ChrW(945)
ElseIf OptionButton4.Value = True Then
  tegn = VBA.ChrW(946)
ElseIf OptionButton5.Value = True Then
  tegn = VBA.ChrW(947)
ElseIf OptionButton6.Value = True Then
  tegn = VBA.ChrW(948)
ElseIf OptionButton7.Value = True Then
  tegn = VBA.ChrW(949)
ElseIf OptionButton8.Value = True Then
  tegn = VBA.ChrW(951)
ElseIf OptionButton9.Value = True Then
  tegn = VBA.ChrW(952) ' theta
ElseIf OptionButton10.Value = True Then
  tegn = VBA.ChrW(955) ' lambda
ElseIf OptionButton11.Value = True Then
  tegn = VBA.ChrW(956) 'mu
ElseIf OptionButton12.Value = True Then
  tegn = VBA.ChrW(960) ' pi
ElseIf OptionButton13.Value = True Then
  tegn = VBA.ChrW(961) ' rho
ElseIf OptionButton14.Value = True Then
  tegn = VBA.ChrW(963) ' sigma
ElseIf OptionButton15.Value = True Then
  tegn = VBA.ChrW(966)
ElseIf OptionButton16.Value = True Then
  tegn = VBA.ChrW(969) ' omega
ElseIf OptionButton17.Value = True Then
  tegn = VBA.ChrW(185)
ElseIf OptionButton18.Value = True Then
  tegn = VBA.ChrW(178)
ElseIf OptionButton19.Value = True Then
  tegn = VBA.ChrW(179)
ElseIf OptionButton20.Value = True Then
  tegn = VBA.ChrW(183)
ElseIf OptionButton21.Value = True Then
  tegn = VBA.ChrW(176) ' grad tegn
ElseIf OptionButton22.Value = True Then
  tegn = ChrW(136)
End If

Me.Hide

GoTo slut
Fejl:
slut:
End Sub


Private Sub UserForm_Activate()
    SetCaptions
End Sub

Sub SetCaptions()
Me.Caption = Sprog.A(234)
Label1.Caption = Sprog.A(230)
Label2.Caption = Sprog.A(231)
Label3.Caption = Sprog.A(232)
CommandButton1.Caption = Sprog.A(233)

End Sub
