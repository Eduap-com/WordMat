VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormSelectVar 
   Caption         =   "Løs ligning"
   ClientHeight    =   4695
   ClientLeft      =   -30
   ClientTop       =   75
   ClientWidth     =   10560
   OleObjectBlob   =   "UserFormSelectVar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormSelectVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vars As String ' semikolon separeret liste over de variable som der skal vælges mellem. Sættes inden show kaldes
Public DefS As String ' definitioner der er i dokumentet. Sættes inden show kaldes
Public TempDefs As String ' Kan både sættes før show kaldes og på formen
Public SelectedVar As String ' Den variabel der er blevet valgt
Public NoEq As Integer ' no of equations to solve for
Public Eliminate As Boolean
Private Svars As Variant ' array der holder variabelnavne.  som de skal returneres dvs. uden asciikonvertering

Private Sub CommandButton_ok_Click()
On Error GoTo fejl
Dim Arr As Variant
Dim i As Integer, c As Integer
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
    
    If NoEq <= 1 Then
        ListBox_vars.MultiSelect = fmMultiSelectSingle
        If TextBox_variabel.Text = "" Then
            SelectedVar = Svars(ListBox_vars.ListIndex)
    '        SelectedVar = ListBox_vars.value
        Else
            SelectedVar = TextBox_variabel.Text
        End If
    Else
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
    End If
    
    TempDefs = TextBox_def.Text
    TempDefs = Trim(TempDefs)
    If Len(TempDefs) > 2 Then
    TempDefs = Replace(TempDefs, ",", ".")
    Arr = Split(TempDefs, VbCrLfMac)

    TempDefs = vbNullString
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
    
    ' Hvis units er on, så må man ikke løse for en enhed
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
    SelectedVar = vbNullString
slut:
    Me.hide
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
        DefS = " " & omax.ConvertToAscii(DefS)
        DefS = Replace(DefS, "$", vbCrLf & " ")
        DefS = Replace(DefS, ":=", " = ")
        DefS = Replace(DefS, ":", " = ")
        
        If Not Radians Then DefS = Replace(DefS, "%pi/180*", "")
        
        DefS = Replace(DefS, "*", MaximaGangeTegn)
        
        DefS = Replace(DefS, "%pi", ChrW(&H3C0))
        DefS = Replace(DefS, "%i", "i")
        DefS = Replace(DefS, "log(", "ln(")
        DefS = Replace(DefS, "log10(", "log(")
        DefS = Replace(DefS, "^(x)", ChrW(&H2E3))
        DefS = Replace(DefS, "^(2)", ChrW(&HB2))
        DefS = Replace(DefS, "^(3)", ChrW(&HB3))
        DefS = Replace(DefS, "^(4)", ChrW(&H2074))
        DefS = Replace(DefS, "^(5)", ChrW(&H2075))
        DefS = Replace(DefS, "^(6)", ChrW(&H2076))
        DefS = Replace(DefS, "^(7)", ChrW(&H2077))
        DefS = Replace(DefS, "^(8)", ChrW(&H2078))
        DefS = Replace(DefS, "^(9)", ChrW(&H2079))
        DefS = Replace(DefS, "^(-1)", ChrW(&H207B) & ChrW(&HB9))
        DefS = Replace(DefS, "^(-2)", ChrW(&H207B) & ChrW(&HB2))
        DefS = Replace(DefS, "^(-3)", ChrW(&H207B) & ChrW(&HB3))
        
        DefS = Replace(DefS, "_0(", ChrW(&H2080) & "(")
        DefS = Replace(DefS, "_1(", ChrW(&H2081) & "(")
        DefS = Replace(DefS, "_2(", ChrW(&H2082) & "(")
        DefS = Replace(DefS, "_3(", ChrW(&H2083) & "(")
        DefS = Replace(DefS, "_4(", ChrW(&H2084) & "(")
        DefS = Replace(DefS, "_5(", ChrW(&H2085) & "(")
        DefS = Replace(DefS, "_6(", ChrW(&H2086) & "(")
        DefS = Replace(DefS, "_7(", ChrW(&H2087) & "(")
        DefS = Replace(DefS, "_8(", ChrW(&H2088) & "(")
        DefS = Replace(DefS, "_9(", ChrW(&H2089) & "(")
        DefS = Replace(DefS, "_a(", ChrW(&H2090) & "(")
        DefS = Replace(DefS, "_x(", ChrW(&H2093) & "(")
        DefS = Replace(DefS, "_n(", ChrW(&H2099) & "(")
        
        DefS = Replace(DefS, "minf", "-" & ChrW(&H221E))
        DefS = Replace(DefS, "inf", ChrW(&H221E))
        
        DefS = Replace(DefS, "sqrt(", ChrW(&H221A) & "(")
        DefS = Replace(DefS, "NIntegrate(", ChrW(&H222B) & "(")
        DefS = Replace(DefS, "Integrate(", ChrW(&H222B) & "(")
        DefS = Replace(DefS, "integrate(", ChrW(&H222B) & "(")
        DefS = Replace(DefS, "<=", VBA.ChrW(8804))
        DefS = Replace(DefS, ">=", VBA.ChrW(8805))
        DefS = Replace(DefS, "ae", "æ")
        DefS = Replace(DefS, "oe", "ø")
        DefS = Replace(DefS, "aa", "å")
        DefS = Replace(DefS, "AE", "Æ")
        DefS = Replace(DefS, "OE", "Ø")
        DefS = Replace(DefS, "AA", "Å")
        
        'græske bogstaver
        DefS = Replace(DefS, "gamma", VBA.ChrW(915))    ' stort gammategn
        DefS = Replace(DefS, "Delta", VBA.ChrW(916))
        DefS = Replace(DefS, "delta", VBA.ChrW(948))
        DefS = Replace(DefS, "alpha", VBA.ChrW(945))
        DefS = Replace(DefS, "beta", VBA.ChrW(946))
        DefS = Replace(DefS, "gammaLB", VBA.ChrW(947))
        DefS = Replace(DefS, "theta", VBA.ChrW(952))
        DefS = Replace(DefS, "Theta", VBA.ChrW(920))
        DefS = Replace(DefS, "lambda", VBA.ChrW(955))
        DefS = Replace(DefS, "Lambda", VBA.ChrW(923))
        DefS = Replace(DefS, "mu", VBA.ChrW(956))
        DefS = Replace(DefS, "rho", VBA.ChrW(961))
        DefS = Replace(DefS, "sigma", VBA.ChrW(963))
        DefS = Replace(DefS, "Sigma", VBA.ChrW(931))
        DefS = Replace(DefS, "varphi", VBA.ChrW(966))
        DefS = Replace(DefS, "phi", VBA.ChrW(981))
        DefS = Replace(DefS, "Phi", VBA.ChrW(934))
        DefS = Replace(DefS, "varepsilon", VBA.ChrW(949))
        DefS = Replace(DefS, "epsilon", VBA.ChrW(1013))
        DefS = Replace(DefS, "psi", VBA.ChrW(968))
        DefS = Replace(DefS, "Psi", VBA.ChrW(936))
        DefS = Replace(DefS, "Xi", VBA.ChrW(926))
        DefS = Replace(DefS, "xi", VBA.ChrW(958))
        DefS = Replace(DefS, "Chi", VBA.ChrW(935))
        DefS = Replace(DefS, "chi", VBA.ChrW(967))
        DefS = Replace(DefS, "Pi", VBA.ChrW(928))
        DefS = Replace(DefS, "tau", VBA.ChrW(964))
        DefS = Replace(DefS, "greek-nu", VBA.ChrW(957))
        DefS = Replace(DefS, "kappa", VBA.ChrW(954))
        DefS = Replace(DefS, "eta", VBA.ChrW(951))
        DefS = Replace(DefS, "zeta", VBA.ChrW(950))
        DefS = Replace(DefS, "omega", VBA.ChrW(969))    ' lille omega
    
        DefS = Replace(DefS, "((x))", "(x)")
    
        If DecSeparator = "," Then
            '        DefS = Replace(DefS, ",", ";")
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
    
    If Eliminate Then
        ListBox_vars.MultiSelect = fmMultiSelectMulti
        For i = 0 To NoEq - 2
            ListBox_vars.Selected(i) = True
        Next
        Label_choose.Caption = Sprog.A(247) & " " & NoEq - 1 & " " & Sprog.A(245)
        Label_tast.Caption = Sprog.A(248) & " " & NoEq - 1 & " " & Sprog.A(249)
    ElseIf NoEq > 1 Then
        ListBox_vars.MultiSelect = fmMultiSelectMulti
        For i = 0 To NoEq - 1
            ListBox_vars.Selected(i) = True
        Next
        Label_choose.Caption = Sprog.A(250) & " " & NoEq & " " & Sprog.A(245)
        Label_tast.Caption = Sprog.A(251) & " " & NoEq & " " & Sprog.A(249)
    Else
        ListBox_vars.MultiSelect = fmMultiSelectSingle
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
    If NoEq > 1 Then
        Me.Caption = Sprog.SolveSystem
        Label_choose.Caption = Sprog.ChooseVariables
    Else
        Me.Caption = Sprog.SolveEquation
        Label_choose.Caption = Sprog.ChooseVariable
    End If

    Label_tast.Caption = Sprog.WriteVariable
    Label_ok.Caption = Sprog.OK
    Label_cancel.Caption = Sprog.Cancel
    Label4.Caption = Sprog.PresentDefs
    Label5.Caption = Sprog.TempDefs
    Label8.Caption = Sprog.RibSettings
    Frame5.Caption = Sprog.Exact & " ?"
    OptionButton_exactandnum.Caption = Sprog.Auto
    OptionButton_exactonly.Caption = Sprog.Exact
    OptionButton_numonly.Caption = Sprog.Numeric
    CheckBox_vidnotation.Caption = Sprog.ScientificNotation
    Label6.Caption = Sprog.SignificantFigures
    Label_enheder.Caption = Sprog.OutputUnits & ":"
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

