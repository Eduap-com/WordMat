VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormReduce 
   Caption         =   "Omskriv"
   ClientHeight    =   6195
   ClientLeft      =   -30
   ClientTop       =   75
   ClientWidth     =   10845
   OleObjectBlob   =   "UserFormReduce.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormReduce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This form shows different methods to reduce a mathematical expression

Public annuller As Boolean
Public TempDefs As String
Public Vars As String
Public SammeLinje As Boolean

Private EventsCol As New Collection
Sub SetEscEvents(ControlColl As Controls)
' SetEscEvents Me.Controls     in Initialize
    Dim CE As CEvents, c As control, TN As String, F As MSForms.Frame
    On Error Resume Next
    For Each c In ControlColl ' Me.Controls
        TN = TypeName(c)
        If TN = "CheckBox" Then
            Set CE = New CEvents: Set CE.CheckBoxControl = c: EventsCol.Add CE
        ElseIf TN = "OptionButton" Then
            Set CE = New CEvents: Set CE.OptionButtonControl = c: EventsCol.Add CE
        ElseIf TN = "ComboBox" Then
            Set CE = New CEvents: Set CE.ComboBoxControl = c: EventsCol.Add CE
        ElseIf TN = "Label" Then
            Set CE = New CEvents: Set CE.LabelControl = c: EventsCol.Add CE
        ElseIf TN = "TextBox" Then
            Set CE = New CEvents: Set CE.TextBoxControl = c: EventsCol.Add CE
        ElseIf TN = "CommandButton" Then
            Set CE = New CEvents: Set CE.CommandButtonControl = c: EventsCol.Add CE
        ElseIf TN = "ListBox" Then
            Set CE = New CEvents: Set CE.ListBoxControl = c: EventsCol.Add CE
        ElseIf TN = "Frame" Then
            Set F = c
            SetEscEvents F.Controls
        End If
    Next
End Sub
Private Sub CommandButton_annuller_Click()
    annuller = True
    Me.hide
End Sub
Sub ExecuteOK()
    Dim arr() As String, i As Integer
    annuller = False
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

    MaximaDecOutType = ComboBox_DecType.ListIndex + 1

    TempDefs = TextBox_def.text
    TempDefs = Trim$(TempDefs)
    If Len(TempDefs) > 2 Then
    TempDefs = Replace(TempDefs, ",", ".")
    arr = Split(TempDefs, VbCrLfMac)
    TempDefs = ""
    For i = 0 To UBound(arr)
        If Len(arr(i)) > 2 And Not Right$(arr(i), 1) = "=" Then
            TempDefs = TempDefs & arr(i) & ListSeparator
        End If
    Next
    If Right$(TempDefs, 1) = ListSeparator Then
        TempDefs = Left$(TempDefs, Len(TempDefs) - 1)
    End If
    End If


    Me.hide
    Application.ScreenUpdating = False

End Sub
Private Sub CommandButton_ok_Click()
    ExecuteOK
End Sub

Private Sub UserForm_Activate()
Dim arr() As String, i As Integer
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

    ComboBox_cifre.Value = MaximaCifre
    ComboBox_DecType.ListIndex = MaximaDecOutType - 1
    
    
    arr = Split(Vars, ";")
    ' definitioner vises
    
    For i = 0 To UBound(arr)
        If arr(i) <> "" Then
            TextBox_def.text = TextBox_def.text & arr(i) & "=" & VbCrLfMac    ' Temp definitions
        End If
    Next
    
    If CASengine = 0 Then
        CheckBox_rationaliser.visible = True
        Label_rationaliser.visible = True
        CheckBox_completesquare.visible = True
        Label_completesquare.visible = True
    Else
        CheckBox_rationaliser.visible = False
        Label_rationaliser.visible = False
        
        CheckBox_completesquare.visible = False
        Label_completesquare.visible = False
    End If
    
    annuller = True
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
    FillComboBoxDecType
    SetEscEvents Me.Controls
End Sub

Private Sub SetCaptions()
    Me.Caption = TT.A(456)
    Label5.Caption = TT.A(764)
    Frame6.Caption = TT.A(816) & " output"
    Label_cancel.Caption = TT.Cancel
    Label_enheder.Caption = TT.A(168)
    Frame5.Caption = TT.A(710) & "?"
    OptionButton_exactonly.Caption = TT.A(710)
    OptionButton_numonly.Caption = TT.A(711)
    Label8.Caption = TT.A(789)
    CheckBox_auto.Caption = TT.A(790)
    Label_omskriv.Caption = TT.A(791)
    CheckBox_factor.Caption = TT.A(806)
    Label_factor.Caption = TT.A(792)
    CheckBox_expand.Caption = TT.A(807)
    Label_expand.Caption = TT.A(793)
    CheckBox_rationaliser.Caption = TT.A(794)
    Label_rationaliser.Caption = TT.A(795)
    CheckBox_trigreduce.Caption = TT.A(796)
    Label_trigreduce.Caption = TT.A(797)
    CommandButton_ok.Caption = TT.OK
    
    CheckBox_completesquare.Caption = TT.A(697)
    Label_completesquare.Caption = TT.A(698)
    
    CheckBox_sammelinje.Caption = TT.A(798)
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
    Me.hide
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
Sub FillComboBoxDecType()
    ComboBox_DecType.Clear
    ComboBox_DecType.AddItem "Decimaler"
    ComboBox_DecType.AddItem TT.A(687)
    ComboBox_DecType.AddItem TT.A(669)
End Sub

