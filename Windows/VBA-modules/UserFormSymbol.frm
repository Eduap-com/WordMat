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
Option Explicit
' used by 2dGraph to insert symbols in labels

Public tegn As String
Private Sub CommandButton1_Click()
    On Error GoTo fejl

    If OptionButton1.value = True Then
        tegn = VBA.ChrW$(916) ' Delta
    ElseIf OptionButton2.value = True Then
        tegn = VBA.ChrW$(937)
    ElseIf OptionButton3.value = True Then
        tegn = VBA.ChrW$(945)
    ElseIf OptionButton4.value = True Then
        tegn = VBA.ChrW$(946)
    ElseIf OptionButton5.value = True Then
        tegn = VBA.ChrW$(947)
    ElseIf OptionButton6.value = True Then
        tegn = VBA.ChrW$(948)
    ElseIf OptionButton7.value = True Then
        tegn = VBA.ChrW$(949)
    ElseIf OptionButton8.value = True Then
        tegn = VBA.ChrW$(951)
    ElseIf OptionButton9.value = True Then
        tegn = VBA.ChrW$(952) ' theta
    ElseIf OptionButton10.value = True Then
        tegn = VBA.ChrW$(955) ' lambda
    ElseIf OptionButton11.value = True Then
        tegn = VBA.ChrW$(956) 'mu
    ElseIf OptionButton12.value = True Then
        tegn = VBA.ChrW$(960) ' pi
    ElseIf OptionButton13.value = True Then
        tegn = VBA.ChrW$(961) ' rho
    ElseIf OptionButton14.value = True Then
        tegn = VBA.ChrW$(963) ' sigma
    ElseIf OptionButton15.value = True Then
        tegn = VBA.ChrW$(966)
    ElseIf OptionButton16.value = True Then
        tegn = VBA.ChrW$(969) ' omega
    ElseIf OptionButton17.value = True Then
        tegn = VBA.ChrW$(185)
    ElseIf OptionButton18.value = True Then
        tegn = VBA.ChrW$(178)
    ElseIf OptionButton19.value = True Then
        tegn = VBA.ChrW$(179)
    ElseIf OptionButton20.value = True Then
        tegn = VBA.ChrW$(183)
    ElseIf OptionButton21.value = True Then
        tegn = VBA.ChrW$(176) ' grad tegn
    ElseIf OptionButton22.value = True Then
        tegn = ChrW$(136)
    End If

    Me.hide

    GoTo slut
fejl:
slut:
End Sub

Private Sub UserForm_Activate()
    SetCaptions
End Sub

Sub SetCaptions()
    Me.Caption = TT.A(234)
    Label1.Caption = TT.A(230)
    Label2.Caption = TT.A(231)
    Label3.Caption = TT.A(232)
    CommandButton1.Caption = TT.A(233)
End Sub
