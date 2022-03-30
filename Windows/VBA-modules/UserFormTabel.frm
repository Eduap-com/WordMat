VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormTabel 
   Caption         =   "Punkter"
   ClientHeight    =   4470
   ClientLeft      =   -30
   ClientTop       =   75
   ClientWidth     =   4575
   OleObjectBlob   =   "UserFormTabel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormTabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public raekker As Integer
Public kolonner As Integer

Private Sub CommandButton_ok_Click()
    Me.Hide
End Sub

Private Sub Labeladd_Click()
AddNewRow
Labeladd.top = 34 + 12 * raekker
MsgBox PunktText
End Sub

Sub AddNewRow()
    Dim i As Integer
    raekker = raekker + 1
    Call AddTextbox("X" & raekker, 34 + 12 * raekker, 6)
    Call AddLabel("-", 34 + 12 * raekker, 10 + 24 * kolonner)
    Call AddLabel("+", 34 + 12 * raekker, 10 + 24 * kolonner)
    If kolonner > 1 Then
        Call AddTextbox("Y" & raekker, 34 + 12 * raekker, 30)
    End If
    If kolonner > 2 Then
        Call AddTextbox("Z" & raekker, 34 + 12 * raekker, 54)
    End If
    For i = 3 To kolonner
        Call AddTextbox(i & raekker, 34 + 12 * raekker, 6 + 24 * i)
    Next
    
End Sub

Sub AddTextbox(n As String, t As Integer, l As Integer)
Dim tb As MSForms.TextBox

Set tb = Me.Controls.Add("Forms.textbox.1")
        With tb
            .Name = n
            .top = t
            .Left = l
            .Width = 24
            .Height = 12
            .Font.Size = 7
            .Font.Name = "Tahoma"
            .BorderStyle = fmBorderStyleSingle
            .SpecialEffect = fmSpecialEffectFlat
            .SelectionMargin = False
            .WordWrap = False
        End With
End Sub
Sub AddLabel(n As String, t As Integer, l As Integer)
Dim la As MSForms.Label

Set la = Me.Controls.Add("Forms.label.1")
        With la
            .Name = ""
            .Caption = n
            .top = t
            .Left = l
            .Width = 8
            .Height = 12
            .Font.Size = 8
            .Font.Name = "Tahoma"
            .BorderStyle = fmBorderStyleSingle
            .SpecialEffect = fmSpecialEffectFlat
'            .SelectionMargin = False
            .WordWrap = False
        End With
        
        
End Sub



Public Property Get PunktText() As Variant
Dim ctrlx As Object
Dim ctrly As Object
Dim text As String
Dim i As Integer

Do
Set ctrlx = Me.Controls("X" & i)
Set ctrly = Me.Controls("Y" & i)
text = text & ctrlx.text & ";" & ctrly.text & "$"
i = i + 1
Loop While i < raekker


PunktText = text
End Property

Public Property Let PunktText(ByVal vNewValue As Variant)

End Property
