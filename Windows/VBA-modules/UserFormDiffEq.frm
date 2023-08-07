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
    
    
    TempDefs = TextBox_def.Text
    TempDefs = Trim(TempDefs)
    If Len(TempDefs) > 2 Then
    TempDefs = Replace(TempDefs, ",", ".")
    arr = Split(TempDefs, VbCrLfMac)
    TempDefs = ""
    For i = 0 To UBound(arr)
        If Len(arr(i)) > 2 And Not right(arr(i), 1) = "=" Then
            If Split(arr(i), "=")(0) <> TextBox_funktion.Text Then ' kan ikke definere variabel der løses for
                TempDefs = TempDefs & omax.CodeForMaxima(arr(i)) & ListSeparator
            Else
                MsgBox Sprog.A(252) & " " & TextBox_funktion.Text & " " & Sprog.A(253), vbOKOnly, Sprog.Error
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

Private Sub CommandButton_solvenum_Click()
   Dim F As String, arr() As String
   arr = Split(Label_ligning.Caption, "=")
   If UBound(arr) > 0 Then F = arr(1)
   If Len(arr(0)) > 2 Then
      MsgBox "Differentialligningen skal være på formen y'=...  for at den kan løses numerisk" & vbCrLf & "", vbOKOnly, "Fejl"
      Exit Sub
   End If
   luk = True
   Me.Hide
   UserFormDeSolveNumeric.TextBox_varx.Text = TextBox_variabel.Text
   UserFormDeSolveNumeric.TextBox_var1.Text = TextBox_funktion.Text
   UserFormDeSolveNumeric.TextBox_eq1.Text = F
   If TextBox_starty.Text = vbNullString Then
      UserFormDeSolveNumeric.TextBox_init1.Text = "1"
   Else
      UserFormDeSolveNumeric.TextBox_init1.Text = TextBox_starty.Text
   End If
   UserFormDeSolveNumeric.TextBox_xmin.Text = TextBox_startx.Text
   UserFormDeSolveNumeric.TextBox_xmax.Text = TextBox_startx.Text + 10
   UserFormDeSolveNumeric.TextBox_definitioner.Text = TextBox_def.Text
   UserFormDeSolveNumeric.Show
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

    If InStr(Label_ligning.Caption, ChrW(180) & ChrW(180)) > 0 Then ' "´´" to accenter der vender opad
        Label_diffy.visible = True
        TextBox_starty2.visible = True
        Label_y2.visible = True
        TextBox_bcx.visible = True
        Label7.visible = True
        Label8.visible = True
        TextBox_bcy.visible = True
        CommandButton_solvenum.visible = False
    Else
        Label_diffy.visible = False
        TextBox_starty2.visible = False
        Label_y2.visible = False
        TextBox_bcx.visible = False
        Label7.visible = False
        Label8.visible = False
        TextBox_bcy.visible = False
        CommandButton_solvenum.visible = True
    End If

    Svars = Split(vars, ";")
    For i = 0 To UBound(Svars)
        If Svars(i) <> "" Then
            svar = omax.ConvertToWordSymbols(Svars(i))
            TextBox_def.Text = TextBox_def.Text & svar & "=" & VbCrLfMac    ' midlertidige definitioner
        End If
    Next

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    luk = True
End Sub

Sub opdaterLabels()
Dim fkt As String
Dim Pos As Integer
On Error Resume Next
    fkt = TextBox_funktion.Text
    Pos = InStr(fkt, "(")
    If Pos > 0 Then
        fkt = Left(fkt, Pos - 1)
    End If
    Label_diffy.Caption = fkt & "'(" & TextBox_startx.Text & ")="
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
