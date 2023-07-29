VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormSelectVars 
   Caption         =   "Løs ligningssystem"
   ClientHeight    =   4470
   ClientLeft      =   -15
   ClientTop       =   75
   ClientWidth     =   7395
   OleObjectBlob   =   "UserFormSelectVars.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormSelectVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public vars As String
Public DefS As String
Public TempDefs As String
Public SelectedVar As String
Public NoEq As Integer ' no of equations to solve for
Public Eliminate As Boolean
Private Svars As Variant ' array der holder variabelnavne som de skal returneres dvs. uden asciikonvertering

Private Sub CommandButton_cancel_Click()
    UFSelectVars.Hide
    Application.ScreenUpdating = False
End Sub

Private Sub CommandButton_ok_Click()
On Error GoTo Fejl
    Dim i As Integer
    Dim c As Integer
    Dim Arr As Variant
    
    For i = 0 To ListBox_vars.ListCount - 1
        If ListBox_vars.Selected(i) Then
'            SelectedVar = SelectedVar & ListBox_vars.List(i) & ","
            SelectedVar = SelectedVar & Svars(i) & ","
            c = c + 1
        End If
    Next
    If Len(TextBox_variabel.Text) > 0 Then
    Arr = Split(TextBox_variabel.Text, ",")
    For i = 0 To UBound(Arr)
            SelectedVar = SelectedVar & Arr(i) & ","
            c = c + 1
    Next
    End If
    
    If SelectedVar <> "" Then
        SelectedVar = Left(SelectedVar, Len(SelectedVar) - 1)
    End If
    
        If Eliminate Then
            If c >= NoEq Then
                MsgBox Sprog.A(244) & " " & NoEq - 1 & " " & Sprog.A(245) & ".", vbOKOnly, Sprog.Error
                SelectedVar = ""
                Exit Sub
            End If
        Else
            If c <> NoEq Then
                MsgBox Sprog.A(246) & " " & NoEq & " " & Sprog.A(245) & ".", vbOKOnly, Sprog.Error
                SelectedVar = ""
                Exit Sub
            End If
        End If
    
    TempDefs = TextBox_def.Text
    TempDefs = Trim(TempDefs)
    If Len(TempDefs) > 2 Then
    TempDefs = Replace(TempDefs, ",", ".")
    Arr = Split(TempDefs, VbCrLfMac)
    TempDefs = ""
    For i = 0 To UBound(Arr)
        If Len(Arr(i)) > 2 And Not right(Arr(i), 1) = "=" Then
            TempDefs = TempDefs & omax.CodeForMaxima(Arr(i)) & ListSeparator
        End If
    Next
    If right(TempDefs, 1) = ListSeparator Then
        TempDefs = Left(TempDefs, Len(TempDefs) - 1)
    End If
    End If
    
    
    GoTo slut
Fejl:
    SelectedVar = ""
slut:
    UFSelectVars.Hide
    Application.ScreenUpdating = False
End Sub

Private Sub ListBox_vars_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    CommandButton_ok_Click
End Sub

Private Sub UserForm_Activate()
'Dim arr As Variant
Dim i As Integer, svar As String
On Error Resume Next
    SetCaptions
    SelectedVar = ""
    ListBox_vars.Clear
    TextBox_variabel.Text = ""

    If MaximaUnits Then
        Label_unitwarning.visible = True
    Else
        Label_unitwarning.visible = False
    End If


    Svars = Split(vars, ";")
    
    For i = 0 To UBound(Svars)
        If Svars(i) <> "" Then
            svar = omax.ConvertToWordSymbols(Svars(i))
            ListBox_vars.AddItem (svar)
            TextBox_def.Text = TextBox_def.Text & svar & "=" & VbCrLfMac
        End If
    Next
    
    
    ' definitioner vises
    If Len(DefS) > 3 Then
'    defs = Mid(defs, 2, Len(defs) - 3)
    DefS = omax.ConvertToAscii(DefS)
    DefS = Replace(DefS, "$", vbCrLf)
    DefS = Replace(DefS, ":=", vbTab & "= ")
    DefS = Replace(DefS, ":", vbTab & "= ")
    If DecSeparator = "," Then
        DefS = Replace(DefS, ",", ";")
        DefS = Replace(DefS, ".", ",")
    End If
    End If
    Label_def.Caption = DefS
    
    If Eliminate Then
        For i = 0 To NoEq - 2
            ListBox_vars.Selected(i) = True
        Next
        Label_choose.Caption = Sprog.A(247) & " " & NoEq - 1 & " " & Sprog.A(245)
        Label_tast.Caption = Sprog.A(248) & " " & NoEq - 1 & " " & Sprog.A(249)
    Else
        For i = 0 To NoEq - 1
            ListBox_vars.Selected(i) = True
        Next
        Label_choose.Caption = Sprog.A(250) & " " & NoEq & " " & Sprog.A(245)
        Label_tast.Caption = Sprog.A(251) & " " & NoEq & " " & Sprog.A(249)
    End If
    Application.ScreenUpdating = True
End Sub

Private Sub SetCaptions()
    Me.Caption = Sprog.SolveSystem
    CommandButton_ok.Caption = Sprog.OK
    CommandButton_cancel.Caption = Sprog.Cancel
    Label_unitwarning.Caption = Sprog.UnitWarning
    Label1.Caption = Sprog.PresentDefs
    Label3.Caption = Sprog.TempDefs
    Label_choose.Caption = Sprog.ChooseVariables
    Label_tast.Caption = ""
    
    
End Sub

