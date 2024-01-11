VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormSelectVar 
   Caption         =   "Løs ligning"
   ClientHeight    =   8655.001
   ClientLeft      =   -30
   ClientTop       =   75
   ClientWidth     =   16515
   OleObjectBlob   =   "UserFormSelectVar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormSelectVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vars As String
Public DefS As String
Public TempDefs As String
Public SelectedVar As String
Private Svars As Variant ' array der holder variabelnavne som de skal returneres dvs. uden asciikonvertering

Private Sub CommandButton_cancel_Click()
    UFSelectVar.hide
    Application.ScreenUpdating = False
End Sub

Private Sub CommandButton_ok_Click()
On Error GoTo fejl
Dim Arr As Variant
Dim i As Integer
    If OptionButton_numonly.Value = True Then
        MaximaExact = 2
    ElseIf OptionButton_exactonly.Value = True Then
        MaximaExact = 1
    Else
        MaximaExact = 0
    End If
    
    
    MaximaVidNotation = CheckBox_vidnotation.Value
    MaximaCifre = ComboBox_cifre.Value
    If MaximaUnits Then
        If OutUnits <> TextBox_outunits.Text Then
            OutUnits = TextBox_outunits.Text
'            omax.MaximaInputStreng = omax.MaximaInputStreng & "uforget(append(globalbaseunitlisting,globalderivedunitlisting))$"
'            If TextBox_outunits.text <> "" Then omax.MaximaInputStreng = omax.MaximaInputStreng & "setunits(" & omax.ConvertUnits(TextBox_outunits.text) & ")$"
        End If
    End If
    
    If TextBox_variabel.Text = "" Then
        SelectedVar = Svars(ListBox_vars.ListIndex)
'        SelectedVar = ListBox_vars.value
    Else
        SelectedVar = TextBox_variabel.Text
    End If
    
    TempDefs = TextBox_def.Text
    TempDefs = Trim(TempDefs)
    If Len(TempDefs) > 2 Then
    TempDefs = Replace(TempDefs, ",", ".")
    Arr = Split(TempDefs, VbCrLfMac)

    TempDefs = ""
    For i = 0 To UBound(Arr)
        If Len(Arr(i)) > 2 And Not right(Arr(i), 1) = "=" Then
            If Split(Arr(i), "=")(0) <> SelectedVar Then ' kan ikke definere variabel der løses for
                TempDefs = TempDefs & omax.CodeForMaxima(Arr(i)) & ListSeparator
            Else
                MsgBox Sprog.A(252) & " " & SelectedVar & " " & Sprog.A(253), vbOKOnly, Sprog.Error
                Exit Sub
            End If
        End If
    Next
    
    ' Hvis units er on så må man ikke løse for en enhed
    If MaximaUnits Then
        If InStr("A,C,F,H,J,K,L,N,S,T,V,W,m,g,u,s", SelectedVar) Then
            MsgBox Sprog.A(254), vbOKOnly, Sprog.Warning
        End If
    End If
    
    If right(TempDefs, 1) = ListSeparator Then
        TempDefs = Left(TempDefs, Len(TempDefs) - 1)
    End If
    End If
    
    If ComboBox_cas.ListIndex = 0 Then
        If CASengine <> 0 Then
            CASengine = 0
'            PrepareMaxima
        End If
    ElseIf ComboBox_cas.ListIndex = 1 Then
        CASengine = 2
    Else
        CASengine = 1
    End If
    
    GoTo slut
fejl:
    SelectedVar = ""
slut:
    UFSelectVar.hide
    Application.ScreenUpdating = False
End Sub

Private Sub ListBox_vars_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    CommandButton_ok_Click
End Sub


Private Sub TextBox_def_Enter()
    TextBox_def.WordWrap = True
    TextBox_def.WordWrap = False
    TextBox_def.WordWrap = True
End Sub

Private Sub UserForm_Activate()
'Dim arr As Variant

Dim i As Integer, svar As String
    SetCaptions
    TextBox_def.WordWrap = True
    TextBox_def.WordWrap = False
    TextBox_def.WordWrap = True
    Application.ScreenUpdating = True
    
    If MaximaUnits Then
        Label_unitwarning.visible = True
        Label_enheder.visible = True
        TextBox_outunits.visible = True
        TextBox_outunits.Text = OutUnits
    Else
        Label_unitwarning.visible = False
        Label_enheder.visible = False
        TextBox_outunits.visible = False
    End If
    
    If MaximaExact = 1 Then
        OptionButton_exactonly.Value = True
    ElseIf MaximaExact = 2 Then
        OptionButton_numonly.Value = True
    Else
        OptionButton_exactandnum.Value = True
    End If

    CheckBox_vidnotation.Value = MaximaVidNotation
    ComboBox_cifre.Value = MaximaCifre

    If CASengine = 0 Then
        ComboBox_cas.ListIndex = 0
    ElseIf CASengine = 1 Then
        ComboBox_cas.ListIndex = 2
    Else
        ComboBox_cas.ListIndex = 1
    End If

    SelectedVar = ""
    ListBox_vars.Clear
    TextBox_variabel.Text = ""
    Svars = Split(vars, ";")
    
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
    
     For i = 0 To UBound(Svars)
        If Svars(i) <> "" Then
            svar = omax.ConvertToWordSymbols(Svars(i))
            ListBox_vars.AddItem (svar)
            If UBound(Svars) > 0 Then TextBox_def.Text = TextBox_def.Text & svar & "=" & VbCrLfMac               ' midlertidige definitioner
        End If
    Next
    If ListBox_vars.ListCount > 0 Then
        ListBox_vars.ListIndex = 0
    End If
    
    ListBox_vars.SetFocus
End Sub
Sub FillComboBoxCifre()
Dim i As Integer
    For i = 2 To 16
        ComboBox_cifre.AddItem i
    Next
End Sub
Sub FillComboBoxCAS()
    ComboBox_cas.AddItem "Maxima"
    ComboBox_cas.AddItem "GeoGebra"
    ComboBox_cas.AddItem "GeoGebra Browser"
End Sub

Private Sub UserForm_Initialize()
    FillComboBoxCifre
    FillComboBoxCAS
    
'    ScaleForm 1.5
    
End Sub
Private Sub SetCaptions()
    Me.Caption = Sprog.SolveEquation
    Label1.Caption = Sprog.ChooseVariable
    Label2.Caption = Sprog.WriteVariable
    CommandButton_ok.Caption = Sprog.OK
    CommandButton_cancel.Caption = Sprog.Cancel
    Label4.Caption = Sprog.PresentDefs
    Label5.Caption = Sprog.TempDefs
    Label8.Caption = Sprog.RibSettings
    Frame5.Caption = Sprog.Exact & " ?"
    OptionButton_exactandnum.Caption = Sprog.Auto
    OptionButton_exactonly.Caption = Sprog.Exact
    OptionButton_numonly.Caption = Sprog.Numeric
    CheckBox_vidnotation.Caption = Sprog.ScientificNotation
    Label6.Caption = Sprog.SignificantFigures
    Label_enheder.Caption = Sprog.OutputUnits
    Label_unitwarning.Caption = Sprog.UnitWarning
End Sub

Sub ScaleForm(SF As Double)
' SF er scalefactor. Ændrer størrelsen på en formen og justerer font og position af alle elementer på formen
Dim c As control
    For Each c In Me.Controls
        c.Left = c.Left * SF
        c.Top = c.Top * SF
        c.Width = c.Width * SF
        c.Height = c.Height * SF
        c.Font.Size = c.Font.Size * SF
    Next
    Me.Width = Me.Width * SF
    Me.Height = Me.Height * SF
    Me.Font.Size = Me.Font.Size * SF
End Sub


