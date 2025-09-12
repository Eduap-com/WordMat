VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormSelectVar 
   Caption         =   "Solve equation"
   ClientHeight    =   4695
   ClientLeft      =   -30
   ClientTop       =   75
   ClientWidth     =   10710
   OleObjectBlob   =   "UserFormSelectVar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormSelectVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Vars As String ' semicolon separated list of variables to choose from. Placed before show is called
Public DefS As String ' definitions that are in the document. Set before show is called
Public TempDefs As String ' Can be set both before show is called and on the form
Public SelectedVar As String ' The variable that has been selected
Public NoEq As Integer ' no of equations to solve for
Public Eliminate As Boolean
Private Svars As Variant ' array that holds variable names as they are to be returned, i.e. without ascii conversion

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

Private Sub CommandButton_ok_Click()
On Error GoTo fejl
Dim arr As Variant
Dim i As Integer, c As Integer
    If OptionButton_numonly.Value = True Then
        MaximaExact = 2
    ElseIf OptionButton_exactonly.Value = True Then
        MaximaExact = 1
    Else
        MaximaExact = 0
    End If
        
    MaximaCifre = ComboBox_cifre.Value
    If MaximaUnits Then
        If OutUnits <> TextBox_outunits.text Then
            OutUnits = TextBox_outunits.text
'            omax.MaximaInputStreng = omax.MaximaInputStreng & "uforget(append(globalbaseunitlisting,globalderivedunitlisting))$"
'            If TextBox_outunits.text <> "" Then omax.MaximaInputStreng = omax.MaximaInputStreng & "setunits(" & omax.ConvertUnits(TextBox_outunits.text) & ")$"
        End If
    End If
    
    If NoEq <= 1 Then
        ListBox_vars.MultiSelect = fmMultiSelectSingle
        If TextBox_variabel.text = "" Then
            SelectedVar = Svars(ListBox_vars.ListIndex)
    '        SelectedVar = ListBox_vars.value
        Else
            SelectedVar = TextBox_variabel.text
        End If
    Else
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
            SelectedVar = Left$(SelectedVar, Len(SelectedVar) - 1)
        End If
    End If
    
    TempDefs = TextBox_def.text
    TempDefs = Trim$(TempDefs)
    If Len(TempDefs) > 2 Then
    TempDefs = Replace(TempDefs, ",", ".")
    arr = Split(TempDefs, VbCrLfMac)

    TempDefs = vbNullString
    For i = 0 To UBound(arr)
        If Len(arr(i)) > 2 And Not right$(arr(i), 1) = "=" Then
            If Split(arr(i), "=")(0) <> SelectedVar Then ' may not define a variable which is solved for
                TempDefs = TempDefs & omax.CodeForMaxima(arr(i)) & ListSeparator
            Else
                MsgBox TT.A(252) & " " & SelectedVar & " " & TT.A(253), vbOKOnly, TT.Error
                Exit Sub
            End If
        End If
    Next
    
    ' If units are on, you may not solve for a unit
    If MaximaUnits Then
        If InStr("A,C,F,H,J,K,L,N,S,T,V,W,m,g,u,s", SelectedVar) Then
            MsgBox TT.A(254), vbOKOnly, TT.A(846)
        End If
    End If
    
    If right$(TempDefs, 1) = ListSeparator Then
        TempDefs = Left$(TempDefs, Len(TempDefs) - 1)
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
    
    MaximaDecOutType = ComboBox_DecType.ListIndex + 1
    
    
    GoTo slut
fejl:
    SelectedVar = vbNullString
slut:
    Me.hide
    Application.ScreenUpdating = False
    RefreshRibbon
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
    Dim i As Integer, svar As String
    On Error GoTo fejl
    SetCaptions
    TextBox_def.WordWrap = True
    TextBox_def.WordWrap = False
    TextBox_def.WordWrap = True
    Application.ScreenUpdating = True
    
    
    If MaximaUnits Then
        Label_unitwarning.visible = True
        Label_enheder.visible = True
        TextBox_outunits.visible = True
        TextBox_outunits.text = OutUnits
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

    ComboBox_cifre.Value = MaximaCifre

    If CASengine = 0 Then
        ComboBox_cas.ListIndex = 0
    ElseIf CASengine = 1 Then
        ComboBox_cas.ListIndex = 2
    Else
        ComboBox_cas.ListIndex = 1
    End If

    ComboBox_DecType.ListIndex = MaximaDecOutType - 1

    SelectedVar = ""
    ListBox_vars.Clear
    TextBox_variabel.text = ""
    Svars = Split(Vars, ";")
    
    ' definitioner vises
    If Len(DefS) > 3 Then
        DefS = FormatDefinitions(DefS)
    End If
    Label_def.Caption = DefS
    
    For i = 0 To UBound(Svars)
        If Svars(i) <> "" Then
            svar = omax.ConvertToWordSymbols(Svars(i))
            ListBox_vars.AddItem (svar)
            If UBound(Svars) > 0 Then TextBox_def.text = TextBox_def.text & svar & "=" & VbCrLfMac    ' Temp definitions
        End If
    Next
    If ListBox_vars.ListCount > 0 Then
        ListBox_vars.ListIndex = 0
    End If
    
    If Eliminate Then
        ListBox_vars.MultiSelect = fmMultiSelectMulti
        If ListBox_vars.ListCount >= NoEq Then
            For i = 0 To NoEq - 2
                ListBox_vars.Selected(i) = True
            Next
        Else
            For i = 0 To ListBox_vars.ListCount - 1
                ListBox_vars.Selected(i) = True
            Next
        End If
        Label_choose.Caption = TT.A(247) & " " & NoEq - 1 & " " & TT.A(245)
        Label_tast.Caption = TT.A(248) & " " & NoEq - 1 & " " & TT.A(249)
    ElseIf NoEq > 1 Then
        ListBox_vars.MultiSelect = fmMultiSelectMulti
        If ListBox_vars.ListCount >= NoEq Then
            For i = 0 To NoEq - 1
                ListBox_vars.Selected(i) = True
            Next
        Else
            For i = 0 To ListBox_vars.ListCount - 1
                ListBox_vars.Selected(i) = True
            Next
        End If
        Label_choose.Caption = TT.A(250) & " " & NoEq & " " & TT.A(245)
        Label_tast.Caption = TT.A(251) & " " & NoEq & " " & TT.A(249)
    Else
        ListBox_vars.MultiSelect = fmMultiSelectSingle
    End If

fejl:
    On Error Resume Next
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
    FillComboBoxDecType
    
    DoEvents
'    ScaleForm 1.5
    SetEscEvents Me.Controls
End Sub
Private Sub SetCaptions()
    If NoEq > 1 Then
        Me.Caption = TT.A(766)
        Label_choose.Caption = TT.A(767)
    Else
        Me.Caption = TT.A(760)
        Label_choose.Caption = TT.A(761)
    End If

    Label_tast.Caption = TT.A(762)
    Label_ok.Caption = TT.OK
    Label_cancel.Caption = TT.Cancel
    Label4.Caption = TT.A(763)
    Label5.Caption = TT.A(764)
    Label8.Caption = TT.A(443)
    Frame5.Caption = TT.A(710) & " ?"
    OptionButton_exactandnum.Caption = TT.A(712)
    OptionButton_exactonly.Caption = TT.A(710)
    OptionButton_numonly.Caption = TT.A(711)
    Label_enheder.Caption = TT.A(168) & ":"
    Label_unitwarning.Caption = TT.A(765)
End Sub

Sub ScaleForm(SF As Double)
' SF is a scalefactor. Changes the size of a form, and adjusts font and position of all elements
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

Private Sub Label_cancel_Click()
    Me.hide
    Application.ScreenUpdating = False
End Sub
Private Sub Label_ok_Click()
    CommandButton_ok_Click
End Sub

Private Sub Label_cancel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorPress
End Sub
Private Sub Label_cancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorHover
End Sub
Private Sub Label_ok_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorPress
End Sub
Private Sub Label_ok_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorHover
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorInactive
    Label_cancel.BackColor = LBColorInactive
End Sub
Sub FillComboBoxDecType()
    ComboBox_DecType.Clear
    ComboBox_DecType.AddItem TT.A(41)
    ComboBox_DecType.AddItem TT.A(687)
    ComboBox_DecType.AddItem TT.A(669)
End Sub
