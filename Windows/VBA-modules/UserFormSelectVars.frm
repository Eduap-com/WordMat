VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormSelectVars 
   Caption         =   "Løs ligningssystem"
   ClientHeight    =   4460
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
Public defs As String
Public tempDefs As String
Public SelectedVar As String
Public NoEq As Integer ' no of equations to solve for
Public Eliminate As Boolean
Private Svars As Variant ' array der holder variabelnavne som de skal returneres dvs. uden asciikonvertering

Private Sub CommandButton_cancel_Click()
    UFSelectVars.hide
    Application.ScreenUpdating = False
End Sub

Private Sub CommandButton_ok_Click()
On Error GoTo Fejl
    Dim i As Integer
    Dim c As Integer
    Dim arr As Variant
    
    For i = 0 To ListBox_vars.ListCount - 1
        If ListBox_vars.Selected(i) Then
'            SelectedVar = SelectedVar & ListBox_vars.List(i) & ","
            SelectedVar = SelectedVar & Svars(i) & ","
            c = c + 1
        End If
    Next
    If Len(TextBox_variabel.text) > 0 Then
    arr = Split(TextBox_variabel.text, ",")
    For i = 0 To UBound(arr)
            SelectedVar = SelectedVar & arr(i) & ","
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
    
    tempDefs = TextBox_def.text
    tempDefs = Trim(tempDefs)
    If Len(tempDefs) > 2 Then
    tempDefs = Replace(tempDefs, ",", ".")
    arr = Split(tempDefs, VbCrLfMac)
    tempDefs = ""
    For i = 0 To UBound(arr)
        If Len(arr(i)) > 2 And Not right(arr(i), 1) = "=" Then
            tempDefs = tempDefs & omax.CodeForMaxima(arr(i)) & ListSeparator
        End If
    Next
    If right(tempDefs, 1) = ListSeparator Then
        tempDefs = Left(tempDefs, Len(tempDefs) - 1)
    End If
    End If
    
    
    GoTo Slut
Fejl:
    SelectedVar = ""
Slut:
    UFSelectVars.hide
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
    TextBox_variabel.text = ""

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
            TextBox_def.text = TextBox_def.text & svar & "=" & VbCrLfMac
        End If
    Next
    
    
    ' definitioner vises
    If Len(defs) > 3 Then
'    defs = Mid(defs, 2, Len(defs) - 3)
    defs = omax.ConvertToAscii(defs)
    defs = Replace(defs, "$", vbCrLf)
    defs = Replace(defs, ":=", vbTab & "= ")
    defs = Replace(defs, ":", vbTab & "= ")
    If DecSeparator = "," Then
        defs = Replace(defs, ",", ";")
        defs = Replace(defs, ".", ",")
    End If
    End If
    Label_def.Caption = defs
    
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
    Label3.Caption = Sprog.tempDefs
    Label_choose.Caption = Sprog.ChooseVariables
    Label_tast.Caption = ""
    
    
End Sub

