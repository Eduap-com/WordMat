VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormOmskriv 
   Caption         =   "Omskriv"
   ClientHeight    =   5595
   ClientLeft      =   -30
   ClientTop       =   75
   ClientWidth     =   10845
   OleObjectBlob   =   "UserFormOmskriv.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormOmskriv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public annuller As Boolean
Public TempDefs As String
Public vars As String
Public SammeLinje As Boolean

Private Sub CommandButton_annuller_Click()
    annuller = True
    Me.Hide
End Sub
Sub ExecuteOK()
    Dim Arr() As String, i As Integer
    SammeLinje = CheckBox_sammelinje.Value

    If OptionButton_numonly.Value = True Then
        MaximaExact = 2
    ElseIf OptionButton_exactonly.Value = True Then
        MaximaExact = 1
    Else
        MaximaExact = 0
    End If
'    MaximaVidNotation = CheckBox_vidnotation.value
    MaximaCifre = ComboBox_cifre.Value
    If MaximaUnits Then
        If OutUnits <> TextBox_outunits.text Then
            OutUnits = TextBox_outunits.text
            omax.MaximaInputStreng = omax.MaximaInputStreng & "uforget(append(globalbaseunitlisting,globalderivedunitlisting))$"
            If TextBox_outunits.text <> "" Then omax.MaximaInputStreng = omax.MaximaInputStreng & "setunits(" & omax.ConvertUnits(TextBox_outunits.text) & ")$"
        End If
    End If

    If OptionButton_logauto.Value = True Then
        MaximaLogOutput = 0
    ElseIf OptionButton_log10.Value = True Then
        MaximaLogOutput = 2
    Else
        MaximaLogOutput = 1
    End If


    TempDefs = TextBox_def.text
    TempDefs = Trim(TempDefs)
    If Len(TempDefs) > 2 Then
    TempDefs = Replace(TempDefs, ",", ".")
    Arr = Split(TempDefs, VbCrLfMac)
    TempDefs = ""
    For i = 0 To UBound(Arr)
        If Len(Arr(i)) > 2 And Not right(Arr(i), 1) = "=" Then
            TempDefs = TempDefs & Arr(i) & ListSeparator
        End If
    Next
    If right(TempDefs, 1) = ListSeparator Then
        TempDefs = Left(TempDefs, Len(TempDefs) - 1)
    End If
    End If


    Me.Hide
    Application.ScreenUpdating = False

End Sub
Private Sub CommandButton_ok_Click()
    ExecuteOK
End Sub

Private Sub UserForm_Activate()
Dim Arr() As String, i As Integer
    SetCaptions
    If MaximaUnits Then
        Label_enheder.visible = True
        TextBox_outunits.visible = True
        TextBox_outunits.text = OutUnits
    Else
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

    If MaximaLogOutput = 0 Then
        OptionButton_logauto.Value = True
    ElseIf MaximaLogOutput = 2 Then
        OptionButton_log10.Value = True
    Else
        OptionButton_ln.Value = True
    End If

    CheckBox_vidnotation.Value = MaximaVidNotation
    ComboBox_cifre.Value = MaximaCifre

    
    
    Arr = Split(vars, ";")
    ' definitioner vises
    
    For i = 0 To UBound(Arr)
        If Arr(i) <> "" Then
            TextBox_def.text = TextBox_def.text & Arr(i) & "=" & VbCrLfMac    ' midlertidige definitioner
        End If
    Next
    
    If CASengine = 0 Then
        CheckBox_rationaliser.visible = True
        Label_rationaliser.visible = True
    Else
        CheckBox_rationaliser.visible = False
        Label_rationaliser.visible = False
    End If
    
    Application.ScreenUpdating = True

End Sub
Sub FillComboBoxCifre()
Dim i As Integer
    For i = 2 To 16
        ComboBox_cifre.AddItem i
    Next

End Sub

Private Sub UserForm_Initialize()
    FillComboBoxCifre
End Sub

Private Sub SetCaptions()
    Me.Caption = Sprog.RibReduce
    Label5.Caption = Sprog.TempDefs
    Frame6.Caption = Sprog.Logarithm & " output"
    Label_cancel.Caption = Sprog.Cancel
    Label_enheder.Caption = Sprog.OutputUnits
    Label6.Caption = Sprog.SignificantFigures
    CheckBox_vidnotation.Caption = Sprog.ScientificNotation
    Frame5.Caption = Sprog.Exact & "?"
    OptionButton_exactonly.Caption = Sprog.Exact
    OptionButton_numonly.Caption = Sprog.Numeric
    Label8.Caption = Sprog.TempSettings
    CheckBox_auto.Caption = Sprog.AutoReduce
    Label_omskriv.Caption = Sprog.AutoReduceExplain
    CheckBox_factor.Caption = Sprog.Factor
    Label_factor.Caption = Sprog.FactorExplain
    CheckBox_expand.Caption = Sprog.Expand
    Label_expand.Caption = Sprog.ExpandExplain
    CheckBox_rationaliser.Caption = Sprog.Rationalize
    Label_rationaliser.Caption = Sprog.RationalizeExplain
    CheckBox_trigreduce.Caption = Sprog.TrigReduce
    Label_trigreduce.Caption = Sprog.TrigReduceExplain
    CommandButton_ok.Caption = Sprog.OK
    
    If Sprog.SprogNr = 1 Then
        CheckBox_sammelinje.Caption = "Inds�t samme linje"
    Else
        CheckBox_sammelinje.Caption = "Insert same line"
    End If
    
End Sub

Private Sub Label_ok_Click()
    CommandButton_ok_Click
End Sub

Private Sub Label_ok_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorPress
End Sub

Private Sub Label_ok_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorHover
End Sub
Private Sub Label_cancel_Click()
    annuller = True
    Me.Hide
End Sub

Private Sub Label_cancel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorPress
End Sub

Private Sub Label_cancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorHover
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorInactive
    Label_cancel.BackColor = LBColorInactive
End Sub
