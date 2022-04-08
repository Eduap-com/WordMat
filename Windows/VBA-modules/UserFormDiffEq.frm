VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormDiffEq 
   Caption         =   "Løsning af differentialligning"
   ClientHeight    =   4680
   ClientLeft      =   -30
   ClientTop       =   75
   ClientWidth     =   6225
   OleObjectBlob   =   "UserFormDiffEq.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormDiffEq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Public DefS As String
Public vars As String
Public TempDefs As String
Public luk As Boolean
Private Svars As Variant ' array der holder variabelnavne som de skal returneres dvs. uden asciikonvertering

Private Sub CommandButton_cancel_Click()
    luk = True
    Me.Hide
End Sub

Private Sub CommandButton_ok_Click()
Dim arr As Variant
Dim i As Integer
    
    
    TempDefs = TextBox_def.text
    TempDefs = Trim(TempDefs)
    If Len(TempDefs) > 2 Then
    TempDefs = Replace(TempDefs, ",", ".")
    arr = Split(TempDefs, VbCrLfMac)
    TempDefs = ""
    For i = 0 To UBound(arr)
        If Len(arr(i)) > 2 And Not right(arr(i), 1) = "=" Then
            If Split(arr(i), "=")(0) <> TextBox_funktion.text Then ' kan ikke definere variabel der l*oe*ses for
                TempDefs = TempDefs & omax.CodeForMaxima(arr(i)) & ListSeparator
            Else
                MsgBox Sprog.A(252) & " " & TextBox_funktion.text & " " & Sprog.A(253), vbOKOnly, Sprog.Error
                Exit Sub
            End If
        End If
    Next
    If right(TempDefs, 1) = ListSeparator Then
        TempDefs = Left(TempDefs, Len(TempDefs) - 1)
    End If
    End If
    
    Me.Hide
End Sub

Private Sub TextBox_funktion_Change()
    opdaterLabels
End Sub

Private Sub TextBox_startx_Change()
    opdaterLabels
End Sub

Private Sub UserForm_Activate()
Dim i As Integer
Dim svar As String
    SetCaptions

    If InStr(Label_ligning.Caption, "‚‚") > 0 Then
        Label_diffy.visible = True
        TextBox_starty2.visible = True
        Label_y2.visible = True
        TextBox_bcx.visible = True
        Label7.visible = True
        Label8.visible = True
        TextBox_bcy.visible = True
    Else
        Label_diffy.visible = False
        TextBox_starty2.visible = False
        Label_y2.visible = False
        TextBox_bcx.visible = False
        Label7.visible = False
        Label8.visible = False
        TextBox_bcy.visible = False
    End If

    Svars = Split(vars, ";")
    For i = 0 To UBound(Svars)
        If Svars(i) <> "" Then
            svar = omax.ConvertToWordSymbols(Svars(i))
            TextBox_def.text = TextBox_def.text & svar & "=" & VbCrLfMac    ' midlertidige definitioner
        End If
    Next

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    luk = True
End Sub

Sub opdaterLabels()
Dim fkt As String
Dim pos As Integer
On Error Resume Next
    fkt = TextBox_funktion.text
    pos = InStr(fkt, "(")
    If pos > 0 Then
        fkt = Left(fkt, pos - 1)
    End If
    Label_diffy.Caption = fkt & "'(" & TextBox_startx.text & ")="
    Label_y.Caption = fkt & "("
    Label_y2.Caption = fkt & "("
End Sub

Sub SetCaptions()
    Me.Caption = Sprog.SolveDE
    Label1.Caption = Sprog.DifferentialEquation
    Label3.Caption = Sprog.IndepVar
    Label2.Caption = Sprog.DepVar
    Label4.Caption = Sprog.StartCond
    Label8.Caption = Sprog.A(297)
    Label_temp.Caption = Sprog.TempDefs
    CommandButton_cancel.Caption = Sprog.Cancel
    CommandButton_ok.Caption = Sprog.OK
End Sub
