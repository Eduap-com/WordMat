VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormTrekant 
   Caption         =   "Trekantsløser"
   ClientHeight    =   6580
   ClientLeft      =   -30
   ClientTop       =   75
   ClientWidth     =   11130
   OleObjectBlob   =   "UserFormTrekant.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormTrekant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim vA As Double
    Dim vB As Double
    Dim vC As Double
    Dim SA As Double
    Dim sb As Double
    Dim sc As Double
    Dim vA2 As Double
    Dim vB2 As Double
    Dim vC2 As Double
    Dim sa2 As Double
    Dim sb2 As Double
    Dim sc2 As Double
    Dim nv As Integer
    Dim ns As Integer
    Dim statustext As String
    Dim succes As Boolean
    Dim elabotext(10) As String
    Dim elabolign(10) As String
    Dim elaboindex As Integer
    Dim inputtext As String

Private Sub Label_nulstil_Click()
    TextBox_A.Text = ""
    TextBox_B.Text = ""
    TextBox_C.Text = ""
    TextBox_sa.Text = ""
    TextBox_sb.Text = ""
    TextBox_sc.Text = ""
    TextBox_captionA.Text = "A"
    TextBox_captionB.Text = "B"
    TextBox_captionC.Text = "C"
    TextBox_captionsa.Text = "a"
    TextBox_captionsb.Text = "b"
    TextBox_captionsc.Text = "c"
    
End Sub

Private Sub Label_ok_Click()

On Error GoTo Fejl

    Dim t As Table
    Dim r As Range
    Dim bc As Integer
    Dim i As Integer
    Dim gemsb As Integer, gemsa As Integer
    
    Application.ScreenUpdating = False
    
    FindSolutions True
    
    If Not succes Then Exit Sub
        
        
    '
    gemsb = Selection.ParagraphFormat.SpaceBefore
    gemsa = Selection.ParagraphFormat.SpaceAfter
            
    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
'        .LineUnitBefore = 0
'        .LineUnitAfter = 0
    End With

    
    ' indsæt i Word
#If Mac Then
#Else
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
#End If
    Selection.Collapse wdCollapseEnd
'    If MaximaForklaring Then
        Selection.TypeParagraph
        Selection.TypeText Sprog.TriangleSolverExplanation3 & inputtext
        Selection.TypeParagraph
'    End If
    Selection.TypeParagraph
    
    
    Set t = ActiveDocument.Tables.Add(Selection.Range, 1, 2)

    Set r = t.Cell(1, 1).Range
    t.Cell(1, 2).Select
    TypeLine TextBox_captionA.Text & " = " & ConvertNumberToStringBC(vA) & VBA.ChrW(176), Not (CBool(ConvertStringToNumber(TextBox_A.Text)))
    TypeLine TextBox_captionB.Text & " = " & ConvertNumberToStringBC(vB) & VBA.ChrW(176), Not (CBool(ConvertStringToNumber(TextBox_B.Text)))
    TypeLine TextBox_captionC.Text & " = " & ConvertNumberToStringBC(vC) & VBA.ChrW(176), Not (CBool(ConvertStringToNumber(TextBox_C.Text)))
    Selection.TypeParagraph
    TypeLine TextBox_captionsa.Text & " = " & ConvertNumberToStringBC(SA), Not (CBool(ConvertStringToNumber(TextBox_sa.Text)))
    TypeLine TextBox_captionsb.Text & " = " & ConvertNumberToStringBC(sb), Not (CBool(ConvertStringToNumber(TextBox_sb.Text)))
    TypeLine TextBox_captionsc.Text & " = " & ConvertNumberToStringBC(sc), Not (CBool(ConvertStringToNumber(TextBox_sc.Text)))
    

    If CheckBox_tal.Value Then
        bc = 3 ' antal betydende cifre på sidelængde på figur
        If Log10(SA) > bc Then bc = Int(Log10(SA)) + 1
        If Log10(sb) > bc Then bc = Int(Log10(sb)) + 1
        If Log10(sc) > bc Then bc = Int(Log10(sc)) + 1
        If bc > MaximaCifre Then bc = MaximaCifre
        InsertTriangle r, vA, sb, sc, ConvertNumberToStringBC(vA, 3) & VBA.ChrW(176), ConvertNumberToStringBC(vB, 3) & VBA.ChrW(176), ConvertNumberToStringBC(vC, 3) & VBA.ChrW(176), ConvertNumberToStringBC(SA, bc), ConvertNumberToStringBC(sb, bc), ConvertNumberToStringBC(sc, bc)
    Else
        InsertTriangle r, vA, sb, sc, TextBox_captionA.Text, TextBox_captionB.Text, TextBox_captionC.Text, TextBox_captionsa.Text, TextBox_captionsb.Text, TextBox_captionsc.Text
    End If
    
    t.Range.Select
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    
    'Hvis 2 løsninger
    If vA2 > 0 Then
    MsgBox Sprog.TS2Solutions, vbOKOnly, Sprog.TS2Solutions2
    Set t = ActiveDocument.Tables.Add(Selection.Range, 1, 2)
    
    Set r = t.Cell(1, 1).Range
    t.Cell(1, 2).Select
    TypeLine TextBox_captionA.Text & " = " & ConvertNumberToStringBC(vA2) & VBA.ChrW(176), Not (CBool(ConvertStringToNumber(TextBox_A.Text)))
    TypeLine TextBox_captionB.Text & " = " & ConvertNumberToStringBC(vB2) & VBA.ChrW(176), Not (CBool(ConvertStringToNumber(TextBox_B.Text)))
    TypeLine TextBox_captionC.Text & " = " & ConvertNumberToStringBC(vC2) & VBA.ChrW(176), Not (CBool(ConvertStringToNumber(TextBox_C.Text)))
    Selection.TypeParagraph
    TypeLine TextBox_captionsa.Text & " = " & ConvertNumberToStringBC(sa2), Not (CBool(ConvertStringToNumber(TextBox_sa.Text)))
    TypeLine TextBox_captionsb.Text & " = " & ConvertNumberToStringBC(sb2), Not (CBool(ConvertStringToNumber(TextBox_sb.Text)))
    TypeLine TextBox_captionsc.Text & " = " & ConvertNumberToStringBC(sc2), Not (CBool(ConvertStringToNumber(TextBox_sc.Text)))
        
    If CheckBox_tal.Value Then
        bc = 3 ' antal betydende cifre på sidelængde på figur
        If Log10(sa2) > bc Then bc = Int(Log10(sa2)) + 1
        If Log10(sb2) > bc Then bc = Int(Log10(sb2)) + 1
        If Log10(sc2) > bc Then bc = Int(Log10(sc2)) + 1
        If bc > MaximaCifre Then bc = MaximaCifre
        InsertTriangle r, vA2, sb2, sc2, ConvertNumberToStringBC(vA2, 3) & VBA.ChrW(176), ConvertNumberToStringBC(vB2, 3) & VBA.ChrW(176), ConvertNumberToStringBC(vC2, 3) & VBA.ChrW(176), ConvertNumberToStringBC(sa2, bc), ConvertNumberToStringBC(sb2, bc), ConvertNumberToStringBC(sc2, bc)
    Else
        InsertTriangle r, vA2, sb2, sc2, TextBox_captionA.Text, TextBox_captionB.Text, TextBox_captionC.Text, TextBox_captionsa.Text, TextBox_captionsb.Text, TextBox_captionsc.Text
    End If

    t.Range.Select
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    End If
    
    Dim mo As Range
    If CheckBox_forklaring Then
    For i = 0 To elaboindex - 1
    If Len(elabotext(i)) > 0 Then
        Selection.TypeText elabotext(i) & vbCrLf
    End If
    If Len(elabolign(i)) > 0 Then
        Set mo = Selection.OMaths.Add(Selection.Range)
        Selection.TypeText elabolign(i)
        mo.OMaths.BuildUp
        Selection.TypeParagraph
    End If
    Next
    End If
    
    With Selection.ParagraphFormat
        .SpaceBefore = gemsb
        .SpaceAfter = gemsa
    End With
    
#If Mac Then
#Else
        Oundo.EndCustomRecord
#End If
    
GoTo Slut
Fejl:
    MsgBox Sprog.TSNoSolution, vbOKOnly, Sprog.Error
    Exit Sub
Slut:
    SaveSettings
#If Mac Then
    Unload Me
#Else
    Me.Hide
#End If

End Sub

Sub TypeLine(Text As String, fed As Boolean)
    If fed Then
        Selection.Font.Bold = True
    Else
        Selection.Font.Bold = False
    End If
    Selection.TypeText Text
    Selection.Font.Bold = False
    Selection.TypeParagraph

End Sub

Sub FindSolutions(Optional advarsler As Boolean = False)

    Dim d As Double
    Dim san As String
    Dim sbn As String
    Dim scn As String
    Dim vAn As String
    Dim vBn As String
    Dim vCn As String
    
    On Error GoTo Fejl
    
    san = TextBox_captionsa.Text
    sbn = TextBox_captionsb.Text
    scn = TextBox_captionsc.Text
    vAn = TextBox_captionA.Text
    vBn = TextBox_captionB.Text
    vCn = TextBox_captionC.Text
    
    vA = ConvertStringToNumber(TextBox_A.Text)
    vB = ConvertStringToNumber(TextBox_B.Text)
    vC = ConvertStringToNumber(TextBox_C.Text)
    SA = ConvertStringToNumber(TextBox_sa.Text)
    sb = ConvertStringToNumber(TextBox_sb.Text)
    sc = ConvertStringToNumber(TextBox_sc.Text)
    nv = 0
    ns = 0
    elaboindex = 0
    succes = False
    inputtext = ""
    
    If vA > 0 Then
        nv = nv + 1
        inputtext = inputtext & TextBox_captionA.Text & " = " & TextBox_A.Text & VBA.ChrW(176) & " , "
    End If
    If vB > 0 Then
        nv = nv + 1
        inputtext = inputtext & TextBox_captionB.Text & " = " & TextBox_B.Text & VBA.ChrW(176) & " , "
    End If
    If vC > 0 Then
        nv = nv + 1
        inputtext = inputtext & TextBox_captionC.Text & " = " & TextBox_C.Text & VBA.ChrW(176) & " , "
    End If
    If SA > 0 Then
        ns = ns + 1
        inputtext = inputtext & TextBox_captionsa.Text & " = " & TextBox_sa.Text & " , "
    End If
    If sb > 0 Then
        ns = ns + 1
        inputtext = inputtext & TextBox_captionsb.Text & " = " & TextBox_sb.Text & " , "
    End If
    If sc > 0 Then
        ns = ns + 1
        inputtext = inputtext & TextBox_captionsc.Text & " = " & TextBox_sc.Text & " , "
    End If
    If Len(inputtext) > 1 Then inputtext = Left(inputtext, Len(inputtext) - 2)
        
    ' vinkelsum over 180
    If vA + vB + vC > 180 Then
        statustext = Sprog.A(209) ' "Vinkelsummen er over 180"
        If advarsler Then MsgBox Sprog.A(209), vbOKOnly, Sprog.Error
        Exit Sub
    End If
    
    If nv + ns < 3 Then
        statustext = Sprog.TSMissingInfo
        If advarsler Then MsgBox Sprog.TSMissingInfo & vbCrLf & Sprog.A(210), vbOKOnly, Sprog.Error
        Exit Sub
    ElseIf nv + ns > 3 Then
        If nv = 3 And ns = 1 Then
            If vA > 0 And vB > 0 And vC > 0 And vA + vB + vC <> 180 Then
                statustext = Sprog.A(211) ' "Vinkelsummen er ikke 180"
                If advarsler Then MsgBox Sprog.A(211), vbOKOnly, Sprog.Error
                Exit Sub
            End If
        Else
            statustext = Sprog.A(212) ' "Du har indtastet for mange sider/vinkler."
            If advarsler Then MsgBox Sprog.A(212) & vbCrLf & Sprog.A(213), vbOKOnly, Sprog.Error
            Exit Sub
        End If
    Else
        If nv = 3 And ns = 0 Then
        statustext = Sprog.A(214) ' "Mindst en side skal være kendt. 3 vinkler er ikke nok."
        If advarsler Then MsgBox Sprog.A(214) & vbCrLf & Sprog.A(213), vbOKOnly, Sprog.Error
        Exit Sub
        End If
    End If
    
    
    ' 3. vinkel beregnes hvis 2 kendes
    If nv = 2 Then
    If vA > 0 And vB > 0 And vC = 0 Then
        vC = 180 - vA - vB
        AddElaborate Sprog.A(215) & " " & vCn & " " & Sprog.A(216), vCn & "=180" & VBA.ChrW(176) & "-" & vAn & "-" & vBn & "=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vA) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vC) & VBA.ChrW(176)
    ElseIf vA > 0 And vB = 0 And vC > 0 Then
        vB = 180 - vA - vC
        AddElaborate Sprog.A(215) & " " & vBn & " " & Sprog.A(216), vBn & "=180" & VBA.ChrW(176) & "-" & vAn & "-" & vCn & "=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vA) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vC) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vB) & VBA.ChrW(176)
    ElseIf vA = 0 And vB > 0 And vC > 0 Then
        vA = 180 - vB - vC
        AddElaborate Sprog.A(215) & " " & vAn & " " & Sprog.A(216), vAn & "=180" & VBA.ChrW(176) & "-" & vBn & "-" & vCn & "=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vC) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vA) & VBA.ChrW(176)
    End If
    End If
    
    'retvinklede
    If vC = 90 Then
        If ns = 2 Then
            If SA > 0 And sb > 0 Then
                sc = Sqr(SA ^ 2 + sb ^ 2)
                AddElaborate Sprog.A(217) & " " & scn & " " & Sprog.A(218), scn & "=" & VBA.ChrW(8730) & "(" & san & "^2+" & sbn & "^2)=" & VBA.ChrW(8730) & "(" & ConvertNumberToStringBC(SA) & "^2+" & ConvertNumberToStringBC(sb) & "^2)=" & ConvertNumberToStringBC(sc)
                vA = Atn(SA / sb) * 180 / PI
                AddElaborate Sprog.A(215) & " " & vAn & " " & Sprog.A(220), vAn & "=tan^-1 (" & san & "/" & sbn & ")=tan^-1 (" & ConvertNumberToStringBC(SA) & "/" & ConvertNumberToStringBC(sb) & ")=" & ConvertNumberToStringBC(vA) & VBA.ChrW(176)
            ElseIf SA > 0 And sc > 0 Then
                sb = Sqr(sc ^ 2 - SA ^ 2)
                AddElaborate Sprog.A(217) & sbn & " " & Sprog.A(218), sbn & "=" & VBA.ChrW(8730) & "(" & scn & "^2-" & san & "^2)=" & VBA.ChrW(8730) & "(" & ConvertNumberToStringBC(sc) & "^2-" & ConvertNumberToStringBC(SA) & "^2)=" & ConvertNumberToStringBC(sb)
                vA = Arcsin(SA / sc) * 180 / PI
                AddElaborate Sprog.A(215) & " " & vAn & " " & Sprog.A(221), vAn & "=sin^-1 (" & san & "/" & scn & ")=sin^-1 (" & ConvertNumberToStringBC(SA) & "/" & ConvertNumberToStringBC(sc) & ")=" & ConvertNumberToStringBC(vA) & VBA.ChrW(176)
            ElseIf sb > 0 And sc > 0 Then
                SA = Sqr(sc ^ 2 - sb ^ 2)
                AddElaborate Sprog.A(217) & san & " " & Sprog.A(218), san & "=" & VBA.ChrW(8730) & "(" & scn & "^2-" & sbn & "^2)=" & VBA.ChrW(8730) & "(" & ConvertNumberToStringBC(sc) & "^2-" & ConvertNumberToStringBC(sb) & "^2)=" & ConvertNumberToStringBC(SA)
                vA = Arccos(sb / sc) * 180 / PI
                AddElaborate Sprog.A(215) & " " & vAn & " " & Sprog.A(222), vAn & "=cos^-1 (" & sbn & "/" & scn & ")=cos^-1 (" & ConvertNumberToStringBC(sb) & "/" & ConvertNumberToStringBC(sc) & ")=" & ConvertNumberToStringBC(vA) & VBA.ChrW(176)
            End If
            vB = 90 - vA
            AddElaborate Sprog.A(215) & " " & vBn & " " & Sprog.A(216), vBn & "=180" & VBA.ChrW(176) & "-" & vCn & "-" & vAn & "=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vC) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vA) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vB) & VBA.ChrW(176)
        ElseIf ns = 1 Then
            If SA > 0 Then
                sb = SA / Tan(vA * PI / 180)
                sc = SA / Sin(vA * PI / 180)
                AddElaborate Sprog.A(217) & " " & sbn & " " & Sprog.A(220), sbn & "=" & san & "/tan(" & vAn & ")=" & ConvertNumberToStringBC(SA) & "/tan(" & ConvertNumberToStringBC(vA) & ")=" & ConvertNumberToStringBC(sb)
                AddElaborate Sprog.A(217) & " " & scn & " " & Sprog.A(221), scn & "=" & san & "/sin(" & vAn & ")=" & ConvertNumberToStringBC(SA) & "/sin(" & ConvertNumberToStringBC(vA) & ")=" & ConvertNumberToStringBC(sc)
            ElseIf sb > 0 Then
                SA = sb * Tan(vA * PI / 180)
                sc = sb / Cos(vA * PI / 180)
                AddElaborate Sprog.A(217) & " " & san & " " & Sprog.A(220), san & "=" & sbn & VBA.ChrW(183) & "tan(" & vAn & ")=" & ConvertNumberToStringBC(sb) & VBA.ChrW(183) & "tan(" & ConvertNumberToStringBC(vA) & ")=" & ConvertNumberToStringBC(SA)
                AddElaborate Sprog.A(217) & " " & scn & " " & Sprog.A(222), scn & "=" & sbn & "/cos(" & vAn & ")=" & ConvertNumberToStringBC(sb) & "/cos(" & ConvertNumberToStringBC(vA) & ")=" & ConvertNumberToStringBC(sc)
            ElseIf sc > 0 Then
                SA = sc * Sin(vA * PI / 180)
                sb = sc * Cos(vA * PI / 180)
                AddElaborate Sprog.A(217) & " " & san & " " & Sprog.A(221), san & "=" & scn & VBA.ChrW(183) & "sin(" & vAn & ")=" & ConvertNumberToStringBC(sc) & VBA.ChrW(183) & "sin(" & ConvertNumberToStringBC(vA) & ")=" & ConvertNumberToStringBC(SA)
                AddElaborate Sprog.A(217) & " " & sbn & " " & Sprog.A(222), sbn & "=" & scn & VBA.ChrW(183) & "cos(" & vAn & ")=" & ConvertNumberToStringBC(sc) & VBA.ChrW(183) & "cos(" & ConvertNumberToStringBC(vA) & ")=" & ConvertNumberToStringBC(sb)
            End If
        End If
        GoTo Slut
    ElseIf vA = 90 Then
        If ns = 2 Then
            If SA > 0 And sb > 0 Then
                sc = Sqr(SA ^ 2 - sb ^ 2)
                AddElaborate Sprog.A(217) & " " & scn & " " & Sprog.A(218), scn & "=" & VBA.ChrW(8730) & "(" & san & "^2-" & sbn & "^2)=" & VBA.ChrW(8730) & "(" & ConvertNumberToStringBC(SA) & "^2-" & ConvertNumberToStringBC(sb) & "^2)=" & ConvertNumberToStringBC(sc)
                vC = Arccos(sb / SA) * 180 / PI
                AddElaborate Sprog.A(215) & " " & vCn & " " & Sprog.A(222), vCn & "=cos^-1 (" & sbn & "/" & san & ")=cos^-1 (" & ConvertNumberToStringBC(sb) & "/" & ConvertNumberToStringBC(SA) & ")=" & ConvertNumberToStringBC(vC) & VBA.ChrW(176)
            ElseIf SA > 0 And sc > 0 Then
                sb = Sqr(SA ^ 2 - sc ^ 2)
                AddElaborate Sprog.A(217) & " " & sbn & " " & Sprog.A(218), sbn & "=" & VBA.ChrW(8730) & "(" & san & "^2-" & scn & "^2)=" & VBA.ChrW(8730) & "(" & ConvertNumberToStringBC(SA) & "^2-" & ConvertNumberToStringBC(sc) & "^2)=" & ConvertNumberToStringBC(sb)
                vC = Arcsin(sc / SA) * 180 / PI
                AddElaborate Sprog.A(215) & " " & vCn & " " & Sprog.A(221), vCn & "=sin^-1 (" & scn & "/" & san & ")=sin^-1 (" & ConvertNumberToStringBC(sc) & "/" & ConvertNumberToStringBC(SA) & ")=" & ConvertNumberToStringBC(vC) & VBA.ChrW(176)
            ElseIf sb > 0 And sc > 0 Then
                SA = Sqr(sc ^ 2 + sb ^ 2)
                AddElaborate Sprog.A(217) & " " & san & " " & Sprog.A(218), san & "=" & VBA.ChrW(8730) & "(" & scn & "^2+" & sbn & "^2)=" & VBA.ChrW(8730) & "(" & ConvertNumberToStringBC(sc) & "^2+" & ConvertNumberToStringBC(sb) & "^2)=" & ConvertNumberToStringBC(SA)
                vC = Atn(sc / sb) * 180 / PI
                AddElaborate Sprog.A(215) & " " & vCn & " " & Sprog.A(220), vCn & "=tan^-1 (" & scn & "/" & sbn & ")=tan^-1 (" & ConvertNumberToStringBC(sc) & "/" & ConvertNumberToStringBC(sb) & ")=" & ConvertNumberToStringBC(vC) & VBA.ChrW(176)
            End If
            vB = 90 - vC
            AddElaborate Sprog.A(215) & " " & vBn & " " & Sprog.A(216), vBn & "=180" & VBA.ChrW(176) & "-" & vAn & "-" & vCn & "=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vA) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vC) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vB) & VBA.ChrW(176)
        ElseIf ns = 1 Then
            If sc > 0 Then
                SA = sc / Sin(vC * PI / 180)
                sb = sc / Tan(vC * PI / 180)
                AddElaborate Sprog.A(217) & " " & san & " " & Sprog.A(221), san & "=" & scn & "/sin(" & vCn & ")=" & ConvertNumberToStringBC(sc) & "/sin(" & ConvertNumberToStringBC(vC) & ")=" & ConvertNumberToStringBC(SA)
                AddElaborate Sprog.A(217) & " " & sbn & " " & Sprog.A(220), sbn & "=" & scn & "/tan(" & vCn & ")=" & ConvertNumberToStringBC(sc) & "/tan(" & ConvertNumberToStringBC(vC) & ")=" & ConvertNumberToStringBC(sb)
            ElseIf sb > 0 Then
                SA = sb / Cos(vC * PI / 180)
                sc = sb * Tan(vC * PI / 180)
                AddElaborate Sprog.A(217) & " " & san & " " & Sprog.A(222), san & "=" & sbn & "/cos(" & vCn & ")=" & ConvertNumberToStringBC(sb) & "/cos(" & ConvertNumberToStringBC(vC) & ")=" & ConvertNumberToStringBC(SA)
                AddElaborate Sprog.A(217) & " " & scn & " " & Sprog.A(220), scn & "=" & sbn & VBA.ChrW(183) & "tan(" & vCn & ")=" & ConvertNumberToStringBC(sb) & VBA.ChrW(183) & "tan(" & ConvertNumberToStringBC(vC) & ")=" & ConvertNumberToStringBC(sc)
            ElseIf SA > 0 Then
                sb = SA * Cos(vC * PI / 180)
                sc = SA * Sin(vC * PI / 180)
                AddElaborate Sprog.A(217) & " " & sbn & " " & Sprog.A(222), sbn & "=" & san & VBA.ChrW(183) & "cos(" & vCn & ")=" & ConvertNumberToStringBC(SA) & VBA.ChrW(183) & "cos(" & ConvertNumberToStringBC(vC) & ")=" & ConvertNumberToStringBC(sb)
                AddElaborate Sprog.A(217) & " " & scn & " " & Sprog.A(221), scn & "=" & san & VBA.ChrW(183) & "sin(" & vCn & ")=" & ConvertNumberToStringBC(SA) & VBA.ChrW(183) & "sin(" & ConvertNumberToStringBC(vC) & ")=" & ConvertNumberToStringBC(sc)
            End If
        End If
        GoTo Slut
    ElseIf vB = 90 Then
        If ns = 2 Then
            If SA > 0 And sb > 0 Then
                sc = Sqr(sb ^ 2 - SA ^ 2)
                AddElaborate Sprog.A(217) & " " & scn & " " & Sprog.A(218), scn & "=" & VBA.ChrW(8730) & "(" & sbn & "^2-" & san & "^2)=" & VBA.ChrW(8730) & "(" & ConvertNumberToStringBC(sb) & "^2-" & ConvertNumberToStringBC(SA) & "^2)=" & ConvertNumberToStringBC(sc)
                vA = Arcsin(SA / sb) * 180 / PI
                AddElaborate Sprog.A(215) & " " & vAn & " " & Sprog.A(221), vAn & "=sin^-1 (" & san & "/" & sbn & ")=sin^-1 (" & ConvertNumberToStringBC(SA) & "/" & ConvertNumberToStringBC(sb) & ")=" & ConvertNumberToStringBC(vA) & VBA.ChrW(176)
            ElseIf SA > 0 And sc > 0 Then
                sb = Sqr(sc ^ 2 + SA ^ 2)
                AddElaborate Sprog.A(217) & " " & sbn & " " & Sprog.A(218), sbn & "=" & VBA.ChrW(8730) & "(" & scn & "^2+" & san & "^2)=" & VBA.ChrW(8730) & "(" & ConvertNumberToStringBC(sc) & "^2+" & ConvertNumberToStringBC(SA) & "^2)=" & ConvertNumberToStringBC(sb)
                vA = Atn(SA / sc) * 180 / PI
                AddElaborate Sprog.A(215) & " " & vAn & " " & Sprog.A(220), vAn & "=tan^-1 (" & san & "/" & scn & ")=tan^-1 (" & ConvertNumberToStringBC(SA) & "/" & ConvertNumberToStringBC(sc) & ")=" & ConvertNumberToStringBC(vA) & VBA.ChrW(176)
            ElseIf sb > 0 And sc > 0 Then
                SA = Sqr(sb ^ 2 - sc ^ 2)
                AddElaborate Sprog.A(217) & " " & san & " " & Sprog.A(218), san & "=" & VBA.ChrW(8730) & "(" & sbn & "^2-" & scn & "^2)=" & VBA.ChrW(8730) & "(" & ConvertNumberToStringBC(sb) & "^2-" & ConvertNumberToStringBC(sc) & "^2)=" & ConvertNumberToStringBC(SA)
                vA = Arccos(sc / sb) * 180 / PI
                AddElaborate Sprog.A(215) & " " & vAn & " " & Sprog.A(222), vAn & "=cos^-1 (" & scn & "/" & sbn & ")=cos^-1 (" & ConvertNumberToStringBC(sc) & "/" & ConvertNumberToStringBC(sb) & ")=" & ConvertNumberToStringBC(vA) & VBA.ChrW(176)
            End If
            vC = 90 - vA
            AddElaborate Sprog.A(215) & " " & vCn & " " & Sprog.A(216), vCn & "=180" & VBA.ChrW(176) & "-" & vBn & "-" & vAn & "=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vA) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vC) & VBA.ChrW(176)
        ElseIf ns = 1 Then
            If SA > 0 Then
                sb = SA / Sin(vA * PI / 180)
                sc = SA / Tan(vA * PI / 180)
                AddElaborate Sprog.A(217) & " " & sbn & " " & Sprog.A(221), sbn & "=" & san & "/sin(" & vAn & ")=" & ConvertNumberToStringBC(SA) & "/sin(" & ConvertNumberToStringBC(vA) & ")=" & ConvertNumberToStringBC(sb)
                AddElaborate Sprog.A(217) & " " & scn & " " & Sprog.A(220), scn & "=" & san & "/tan(" & vAn & ")=" & ConvertNumberToStringBC(SA) & "/tan(" & ConvertNumberToStringBC(vA) & ")=" & ConvertNumberToStringBC(sc)
            ElseIf sc > 0 Then
                SA = sc * Tan(vA * PI / 180)
                sb = sc / Cos(vA * PI / 180)
                AddElaborate Sprog.A(217) & " " & san & " " & Sprog.A(220), san & "=" & scn & VBA.ChrW(183) & "tan(" & vAn & ")=" & ConvertNumberToStringBC(sc) & VBA.ChrW(183) & "tan(" & ConvertNumberToStringBC(vA) & ")=" & ConvertNumberToStringBC(SA)
                AddElaborate Sprog.A(217) & " " & sbn & " " & Sprog.A(222), sbn & "=" & scn & "/cos(" & vAn & ")=" & ConvertNumberToStringBC(sc) & "/cos(" & ConvertNumberToStringBC(vA) & ")=" & ConvertNumberToStringBC(sb)
            ElseIf sb > 0 Then
                SA = sb * Cos(vC * PI / 180)
                sc = sb * Sin(vC * PI / 180)
                AddElaborate Sprog.A(217) & " " & san & " " & Sprog.A(222), san & "=" & sbn & VBA.ChrW(183) & "cos(" & vCn & ")=" & ConvertNumberToStringBC(sb) & VBA.ChrW(183) & "cos(" & ConvertNumberToStringBC(vC) & ")=" & ConvertNumberToStringBC(SA)
                AddElaborate Sprog.A(217) & " " & scn & " " & Sprog.A(221), scn & "=" & sbn & VBA.ChrW(183) & "sin(" & vCn & ")=" & ConvertNumberToStringBC(sb) & VBA.ChrW(183) & "sin(" & ConvertNumberToStringBC(vC) & ")=" & ConvertNumberToStringBC(sc)
            End If
        End If
        GoTo Slut
    End If
    
    ' Vilkårlig trekant
    If ns = 3 Then
        vA = Arccos((sc ^ 2 + sb ^ 2 - SA ^ 2) / (2 * sc * sb)) * 180 / PI
        vB = Arccos((SA ^ 2 + sc ^ 2 - sb ^ 2) / (2 * SA * sc)) * 180 / PI
        vC = 180 - vB - vA
        AddElaborate Sprog.A(215) & " " & vAn & " og " & vBn & " " & Sprog.A(223), vAn & "=cos^(-1) ((" & scn & "^2 + " & sbn & "^2 - " & san & "^2)/(2" & VBA.ChrW(183) & sbn & VBA.ChrW(183) & scn & "))=cos^(-1) ((" & ConvertNumberToStringBC(sc) & "^2 + " & ConvertNumberToStringBC(sb) & "^2 - " & ConvertNumberToStringBC(SA) & "^2)/(2" & VBA.ChrW(183) & ConvertNumberToStringBC(sb) & VBA.ChrW(183) & ConvertNumberToStringBC(sc) & "))=" & ConvertNumberToStringBC(vA) & VBA.ChrW(176)
        AddElaborate "", vBn & "=cos^(-1) ((" & scn & "^2 + " & san & "^2 - " & sbn & "^2)/(2" & VBA.ChrW(183) & san & VBA.ChrW(183) & scn & "))=cos^(-1) ((" & ConvertNumberToStringBC(sc) & "^2 + " & ConvertNumberToStringBC(SA) & "^2 - " & ConvertNumberToStringBC(sb) & "^2)/(2" & VBA.ChrW(183) & ConvertNumberToStringBC(SA) & VBA.ChrW(183) & ConvertNumberToStringBC(sc) & "))=" & ConvertNumberToStringBC(vB) & VBA.ChrW(176)
        AddElaborate Sprog.A(215) & " " & vCn & " " & Sprog.A(216), vCn & "=180" & VBA.ChrW(176) & "-" & vAn & "-" & vBn & "=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vA) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vC) & VBA.ChrW(176)
    ElseIf ns = 1 Then
        If SA > 0 Then
            sb = SA * Sin(vB * PI / 180) / Sin(vA * PI / 180)
            sc = SA * Sin(vC * PI / 180) / Sin(vA * PI / 180)
            AddElaborate Sprog.A(219) & " " & sbn & " og " & scn & " " & Sprog.A(224), sbn & "=" & san & VBA.ChrW(183) & "sin(" & vBn & ")/sin(" & vAn & ")=" & ConvertNumberToStringBC(SA) & VBA.ChrW(183) & "sin(" & ConvertNumberToStringBC(vB) & VBA.ChrW(176) & ")/sin(" & ConvertNumberToStringBC(vA) & VBA.ChrW(176) & ")=" & ConvertNumberToStringBC(sb)
            AddElaborate "", scn & "=" & san & VBA.ChrW(183) & "sin(" & vCn & ")/sin(" & vAn & ")=" & ConvertNumberToStringBC(SA) & VBA.ChrW(183) & "sin(" & ConvertNumberToStringBC(vC) & VBA.ChrW(176) & ")/sin(" & ConvertNumberToStringBC(vA) & VBA.ChrW(176) & ")=" & ConvertNumberToStringBC(sc)
        ElseIf sb > 0 Then
            SA = sb * Sin(vA * PI / 180) / Sin(vB * PI / 180)
            sc = sb * Sin(vC * PI / 180) / Sin(vB * PI / 180)
            AddElaborate Sprog.A(219) & " " & san & " og " & scn & " " & Sprog.A(224), san & "=" & sbn & VBA.ChrW(183) & "sin(" & vAn & ")/sin(" & vBn & ")=" & ConvertNumberToStringBC(sb) & VBA.ChrW(183) & "sin(" & ConvertNumberToStringBC(vA) & VBA.ChrW(176) & ")/sin(" & ConvertNumberToStringBC(vB) & VBA.ChrW(176) & ")=" & ConvertNumberToStringBC(SA)
            AddElaborate "", scn & "=" & sbn & VBA.ChrW(183) & "sin(" & vCn & ")/sin(" & vBn & ")=" & ConvertNumberToStringBC(sb) & VBA.ChrW(183) & "sin(" & ConvertNumberToStringBC(vC) & VBA.ChrW(176) & ")/sin(" & ConvertNumberToStringBC(vB) & VBA.ChrW(176) & ")=" & ConvertNumberToStringBC(sc)
        Else ' sc>0
            SA = sc * Sin(vA * PI / 180) / Sin(vC * PI / 180)
            sb = sc * Sin(vB * PI / 180) / Sin(vC * PI / 180)
            AddElaborate Sprog.A(219) & " " & san & " og " & sbn & " " & Sprog.A(224), san & "=" & scn & VBA.ChrW(183) & "sin(" & vAn & ")/sin(" & vCn & ")=" & ConvertNumberToStringBC(sc) & VBA.ChrW(183) & "sin(" & ConvertNumberToStringBC(vA) & VBA.ChrW(176) & ")/sin(" & ConvertNumberToStringBC(vC) & VBA.ChrW(176) & ")=" & ConvertNumberToStringBC(SA)
            AddElaborate "", sbn & "=" & scn & VBA.ChrW(183) & "sin(" & vBn & ")/sin(" & vCn & ")=" & ConvertNumberToStringBC(sc) & VBA.ChrW(183) & "sin(" & ConvertNumberToStringBC(vB) & VBA.ChrW(176) & ")/sin(" & ConvertNumberToStringBC(vC) & VBA.ChrW(176) & ")=" & ConvertNumberToStringBC(sb)
        End If
    ElseIf ns = 2 Then
        If vA > 0 Then
            If sb > 0 And sc > 0 Then ' sider om vinkel
                SA = Sqr(sb ^ 2 + sc ^ 2 - 2 * sb * sc * Cos(vA * PI / 180))
                vB = Arccos((SA ^ 2 + sc ^ 2 - sb ^ 2) / (2 * SA * sc)) * 180 / PI
                vC = 180 - vB - vA
                AddElaborate Sprog.A(217) & " " & san & " " & Sprog.A(223), san & "=" & VBA.ChrW(8730) & "(" & sbn & "^2 + " & scn & "^2 - 2" & VBA.ChrW(183) & sbn & VBA.ChrW(183) & scn & VBA.ChrW(183) & "cos(" & vAn & "))=" & VBA.ChrW(8730) & "(" & ConvertNumberToStringBC(sb) & "^2 + " & ConvertNumberToStringBC(sc) & "^2 - 2" & VBA.ChrW(183) & ConvertNumberToStringBC(sb) & VBA.ChrW(183) & ConvertNumberToStringBC(sc) & VBA.ChrW(183) & "cos(" & ConvertNumberToStringBC(vA) & VBA.ChrW(176) & "))=" & ConvertNumberToStringBC(SA)
                AddElaborate Sprog.A(215) & " " & vBn & " " & Sprog.A(223), vBn & "=cos^(-1) ((" & san & "^2 + " & scn & "^2 - " & sbn & "^2)/(2" & VBA.ChrW(183) & san & VBA.ChrW(183) & scn & "))=cos^(-1) ((" & ConvertNumberToStringBC(SA) & "^2 + " & ConvertNumberToStringBC(sc) & "^2 - " & ConvertNumberToStringBC(sb) & "^2)/(2" & VBA.ChrW(183) & ConvertNumberToStringBC(SA) & VBA.ChrW(183) & ConvertNumberToStringBC(sc) & "))=" & ConvertNumberToStringBC(vB) & VBA.ChrW(176)
                AddElaborate Sprog.A(215) & " " & vCn & " " & Sprog.A(216), vCn & "=180" & VBA.ChrW(176) & "-" & vAn & "-" & vBn & "=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vA) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vC) & VBA.ChrW(176)
            ElseIf SA > 0 And sb > 0 Then ' sider ikke om vinkel
                d = SA ^ 2 - sb ^ 2 * Sin(vA * PI / 180) ^ 2
                If d < 0 Then ' ingen løsning
                    GoTo Fejl
                End If
                sc = sb * Cos(vA * PI / 180) + Sqr(d)
                sc2 = sb * Cos(vA * PI / 180) - Sqr(d)
                vB = Arccos((SA ^ 2 + sc ^ 2 - sb ^ 2) / (2 * SA * sc)) * 180 / PI
                vC = 180 - vB - vA
'                sc = sa * Sin(vC * PI / 180) / Sin(vA * PI / 180)
                AddElaborate Sprog.A(217) & " " & scn & " " & Sprog.A(223), san & "^2=" & sbn & "^2+" & scn & "^2-2" & sbn & VBA.ChrW(183) & scn & VBA.ChrW(183) & "cos(" & vAn & ")"
                AddElaborate Sprog.A(225) & " " & scn, scn & "=" & sbn & VBA.ChrW(183) & "cos(" & vAn & ")+" & VBA.ChrW(8730) & "(" & san & "^2-" & sbn & "^2" & VBA.ChrW(183) & "sin(" & vAn & ")^2)=" & ConvertNumberToStringBC(sc)
                If d > 0 Then AddElaborate Sprog.A(226), scn & "_2=" & sbn & VBA.ChrW(183) & "cos(" & vAn & ")-" & VBA.ChrW(8730) & "(" & san & "^2-" & sbn & "^2" & VBA.ChrW(183) & "sin(" & vAn & ")^2)=" & ConvertNumberToStringBC(sc2)
                If sc2 < 0 Then AddElaborate Sprog.A(227), ""
                AddElaborate Sprog.A(215) & " " & vBn & " " & Sprog.A(223), vBn & "=cos^(-1) ((" & san & "^2 + " & scn & "^2 - " & sbn & "^2)/(2" & VBA.ChrW(183) & san & VBA.ChrW(183) & scn & "))=cos^(-1) ((" & ConvertNumberToStringBC(SA) & "^2 + " & ConvertNumberToStringBC(sc) & "^2 - " & ConvertNumberToStringBC(sb) & "^2)/(2" & VBA.ChrW(183) & ConvertNumberToStringBC(SA) & VBA.ChrW(183) & ConvertNumberToStringBC(sc) & "))=" & ConvertNumberToStringBC(vB) & VBA.ChrW(176)
                AddElaborate Sprog.A(215) & " " & vCn & " " & Sprog.A(216), vCn & "=180" & VBA.ChrW(176) & "-" & vAn & "-" & vBn & "=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vA) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vC) & VBA.ChrW(176)
                If d > 0 And sc2 > 0.000000000000001 Then
                    vA2 = vA
                    sb2 = sb
                    sa2 = SA
                    vB2 = Arccos((sa2 ^ 2 + sc2 ^ 2 - sb2 ^ 2) / (2 * sa2 * sc2)) * 180 / PI
                    vC2 = 180 - vB2 - vA2
                    AddElaborate vbCrLf & Sprog.A(228) & " " & scn & " " & Sprog.A(229), ""
                    AddElaborate Sprog.A(215) & " " & vBn & VBA.ChrW(8322) & " findes vha. en cosinusrelation", vBn & "_2=cos^(-1) ((" & san & "^2 + " & scn & "_2^2 - " & sbn & "^2)/(2" & VBA.ChrW(183) & san & "" & VBA.ChrW(183) & scn & "_2))=cos^(-1) ((" & ConvertNumberToStringBC(sa2) & "^2 + " & ConvertNumberToStringBC(sc2) & "^2 - " & ConvertNumberToStringBC(sb2) & "^2)/(2" & VBA.ChrW(183) & ConvertNumberToStringBC(sa2) & VBA.ChrW(183) & ConvertNumberToStringBC(sc2) & "))=" & ConvertNumberToStringBC(vB2) & VBA.ChrW(176)
                    AddElaborate Sprog.A(215) & " " & vCn & VBA.ChrW(8322) & " " & Sprog.A(216), vCn & "_2=180" & VBA.ChrW(176) & "-" & vAn & "-" & vBn & "_2=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vA2) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB2) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vC2) & VBA.ChrW(176)
                End If
            ElseIf SA > 0 And sc > 0 Then ' sider ikke om vinkel
                d = SA ^ 2 - sc ^ 2 * Sin(vA * PI / 180) ^ 2
                If d < 0 Then ' ingen løsning
                    GoTo Fejl
                End If
                sb = sc * Cos(vA * PI / 180) + Sqr(d)
                sb2 = sc * Cos(vA * PI / 180) - Sqr(d)
                vB = Arccos((SA ^ 2 + sc ^ 2 - sb ^ 2) / (2 * SA * sc)) * 180 / PI
                vC = 180 - vB - vA
'                sc = sa * Sin(vC * PI / 180) / Sin(vA * PI / 180)
                AddElaborate Sprog.A(217) & " " & sbn & " " & Sprog.A(223), san & "^2=" & sbn & "^2+" & scn & "^2-2" & sbn & VBA.ChrW(183) & scn & VBA.ChrW(183) & "cos(" & vAn & ")"
                AddElaborate Sprog.A(225) & " " & sbn, sbn & "=" & scn & VBA.ChrW(183) & "cos(" & vAn & ")+" & VBA.ChrW(8730) & "(" & san & "^2-" & scn & "^2" & VBA.ChrW(183) & "sin(" & vAn & ")^2)=" & ConvertNumberToStringBC(sb)
                If d > 0 Then AddElaborate Sprog.A(226), sbn & "_2=" & scn & VBA.ChrW(183) & "cos(" & vAn & ")-" & VBA.ChrW(8730) & "(" & san & "^2-" & scn & "^2" & VBA.ChrW(183) & "sin(" & vAn & ")^2)=" & ConvertNumberToStringBC(sb2)
                If sb2 < 0 Then AddElaborate Sprog.A(227), ""
                AddElaborate Sprog.A(215) & " " & vBn & " " & Sprog.A(223), vBn & "=cos^(-1) ((" & san & "^2 + " & scn & "^2 - " & sbn & "^2)/(2" & VBA.ChrW(183) & san & VBA.ChrW(183) & scn & "))=cos^(-1) ((" & ConvertNumberToStringBC(SA) & "^2 + " & ConvertNumberToStringBC(sc) & "^2 - " & ConvertNumberToStringBC(sb) & "^2)/(2" & VBA.ChrW(183) & ConvertNumberToStringBC(SA) & VBA.ChrW(183) & ConvertNumberToStringBC(sc) & "))=" & ConvertNumberToStringBC(vB) & VBA.ChrW(176)
                AddElaborate Sprog.A(215) & " " & vCn & " " & Sprog.A(216), vCn & "=180" & VBA.ChrW(176) & "-" & vAn & "-" & vBn & "=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vA) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vC) & VBA.ChrW(176)
                If d > 0 And sb2 > 0.000000000000001 Then
                    vA2 = vA
                    sc2 = sc
                    sa2 = SA
                    vB2 = Arccos((sa2 ^ 2 + sc2 ^ 2 - sb2 ^ 2) / (2 * sa2 * sc2)) * 180 / PI
                    vC2 = 180 - vB2 - vA2
                    AddElaborate vbCrLf & Sprog.A(228) & " " & sbn & " " & Sprog.A(229), ""
                    AddElaborate Sprog.A(215) & " " & vBn & VBA.ChrW(8322) & " " & Sprog.A(223), vBn & "_2=cos^(-1) ((" & san & "^2 + " & scn & "^2 - " & sbn & "_2^2)/(2" & VBA.ChrW(183) & san & "" & VBA.ChrW(183) & scn & "))=cos^(-1) ((" & ConvertNumberToStringBC(sa2) & "^2 + " & ConvertNumberToStringBC(sc2) & "^2 - " & ConvertNumberToStringBC(sb2) & "^2)/(2" & VBA.ChrW(183) & ConvertNumberToStringBC(sa2) & VBA.ChrW(183) & ConvertNumberToStringBC(sc2) & "))=" & ConvertNumberToStringBC(vB2) & VBA.ChrW(176)
                    AddElaborate Sprog.A(215) & " " & vCn & VBA.ChrW(8322) & " " & Sprog.A(216), vCn & "_2=180" & VBA.ChrW(176) & "-" & vAn & "-" & vBn & "_2=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vA2) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB2) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vC2) & VBA.ChrW(176)
                End If
            End If
        ElseIf vB > 0 Then
            If SA > 0 And sc > 0 Then ' sider om vinkel
                sb = Sqr(SA ^ 2 + sc ^ 2 - 2 * SA * sc * Cos(vB * PI / 180))
                vA = Arccos((sb ^ 2 + sc ^ 2 - SA ^ 2) / (2 * sb * sc)) * 180 / PI
                vC = 180 - vB - vA
                AddElaborate Sprog.A(217) & " " & sbn & " " & Sprog.A(223), sbn & "=" & VBA.ChrW(8730) & "(" & san & "^2 + " & scn & "^2 - 2" & VBA.ChrW(183) & san & VBA.ChrW(183) & scn & VBA.ChrW(183) & "cos(" & vBn & "))=" & VBA.ChrW(8730) & "(" & ConvertNumberToStringBC(SA) & "^2 + " & ConvertNumberToStringBC(sc) & "^2 - 2" & VBA.ChrW(183) & ConvertNumberToStringBC(SA) & VBA.ChrW(183) & ConvertNumberToStringBC(sc) & VBA.ChrW(183) & "cos(" & ConvertNumberToStringBC(vB) & VBA.ChrW(176) & "))=" & ConvertNumberToStringBC(sb)
                AddElaborate Sprog.A(215) & " " & vAn & " " & Sprog.A(223), vAn & "=cos^(-1) ((" & sbn & "^2 + " & scn & "^2 - " & san & "^2)/(2" & VBA.ChrW(183) & sbn & VBA.ChrW(183) & scn & "))=cos^(-1) ((" & ConvertNumberToStringBC(sb) & "^2 + " & ConvertNumberToStringBC(sc) & "^2 - " & ConvertNumberToStringBC(SA) & "^2)/(2" & VBA.ChrW(183) & ConvertNumberToStringBC(sb) & VBA.ChrW(183) & ConvertNumberToStringBC(sc) & "))=" & ConvertNumberToStringBC(vA) & VBA.ChrW(176)
                AddElaborate Sprog.A(215) & " " & vCn & " " & Sprog.A(216), vCn & "=180" & VBA.ChrW(176) & "-" & vAn & "-" & vBn & "=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vA) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vC) & VBA.ChrW(176)
            ElseIf SA > 0 And sb > 0 Then ' sider ikke om vinkel
                d = sb ^ 2 - SA ^ 2 * Sin(vB * PI / 180) ^ 2
                If d < 0 Then ' ingen løsning
                    GoTo Fejl
                End If
                sc = SA * Cos(vB * PI / 180) + Sqr(d)
                sc2 = SA * Cos(vB * PI / 180) - Sqr(d)
                vA = Arccos((sb ^ 2 + sc ^ 2 - SA ^ 2) / (2 * sb * sc)) * 180 / PI
                vC = 180 - vB - vA
'                sc = sa * Sin(vC * PI / 180) / Sin(vA * PI / 180)
                AddElaborate Sprog.A(217) & " " & scn & " " & Sprog.A(223), sbn & "^2=" & san & "^2+" & scn & "^2-2" & san & VBA.ChrW(183) & scn & VBA.ChrW(183) & "cos(" & vAn & ")"
                AddElaborate Sprog.A(225) & " " & scn, scn & "=" & san & VBA.ChrW(183) & "cos(" & vBn & ")+" & VBA.ChrW(8730) & "(" & sbn & "^2-" & san & "^2" & VBA.ChrW(183) & "sin(" & vBn & ")^2)=" & ConvertNumberToStringBC(sc)
                If d > 0 Then AddElaborate Sprog.A(226), scn & "_2=" & san & VBA.ChrW(183) & "cos(" & vBn & ")-" & VBA.ChrW(8730) & "(" & sbn & "^2-" & san & "^2" & VBA.ChrW(183) & "sin(" & vBn & ")^2)=" & ConvertNumberToStringBC(sc2)
                If sc2 < 0 Then AddElaborate Sprog.A(227), ""
                AddElaborate Sprog.A(215) & " " & vAn & " " & Sprog.A(223), vAn & "=cos^(-1) ((" & sbn & "^2 + " & scn & "^2 - " & san & "^2)/(2" & VBA.ChrW(183) & sbn & VBA.ChrW(183) & scn & "))=cos^(-1) ((" & ConvertNumberToStringBC(sb) & "^2 + " & ConvertNumberToStringBC(sc) & "^2 - " & ConvertNumberToStringBC(SA) & "^2)/(2" & VBA.ChrW(183) & ConvertNumberToStringBC(sb) & VBA.ChrW(183) & ConvertNumberToStringBC(sc) & "))=" & ConvertNumberToStringBC(vA) & VBA.ChrW(176)
                AddElaborate Sprog.A(215) & " " & vCn & " " & Sprog.A(216), vCn & "=180" & VBA.ChrW(176) & "-" & vAn & "-" & vBn & "=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vA) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vC) & VBA.ChrW(176)
                If d > 0 And sc2 > 0.000000000000001 Then
                    vB2 = vB
                    sb2 = sb
                    sa2 = SA
                    vA2 = Arccos((sb2 ^ 2 + sc2 ^ 2 - sa2 ^ 2) / (2 * sb2 * sc2)) * 180 / PI
                    vC2 = 180 - vB2 - vA2
                    AddElaborate vbCrLf & Sprog.A(228) & " " & scn & " " & Sprog.A(229), ""
                    AddElaborate Sprog.A(215) & " " & vAn & VBA.ChrW(8322) & " " & Sprog.A(223), vAn & "_2=cos^(-1) ((" & sbn & "^2 + " & scn & "_2^2 - " & san & "^2)/(2" & VBA.ChrW(183) & sbn & "" & VBA.ChrW(183) & scn & "_2))=cos^(-1) ((" & ConvertNumberToStringBC(sb2) & "^2 + " & ConvertNumberToStringBC(sc2) & "^2 - " & ConvertNumberToStringBC(sa2) & "^2)/(2" & VBA.ChrW(183) & ConvertNumberToStringBC(sb2) & VBA.ChrW(183) & ConvertNumberToStringBC(sc2) & "))=" & ConvertNumberToStringBC(vA2) & VBA.ChrW(176)
                    AddElaborate Sprog.A(215) & " " & vCn & VBA.ChrW(8322) & " " & Sprog.A(216), vCn & "_2=180" & VBA.ChrW(176) & "-" & vAn & "-" & vBn & "_2=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vA2) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB2) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vC2) & VBA.ChrW(176)
                End If
            ElseIf sb > 0 And sc > 0 Then ' sider ikke om vinkel
                d = sb ^ 2 - sc ^ 2 * Sin(vB * PI / 180) ^ 2
                If d < 0 Then ' ingen løsning
                    GoTo Fejl
                End If
                SA = sc * Cos(vB * PI / 180) + Sqr(d)
                sa2 = sc * Cos(vB * PI / 180) - Sqr(d)
                vA = Arccos((sb ^ 2 + sc ^ 2 - SA ^ 2) / (2 * sb * sc)) * 180 / PI
                vC = 180 - vB - vA
'                sc = sa * Sin(vC * PI / 180) / Sin(vA * PI / 180)
                AddElaborate Sprog.A(217) & " " & san & " " & Sprog.A(223), sbn & "^2=" & san & "^2+" & scn & "^2-2" & san & VBA.ChrW(183) & scn & VBA.ChrW(183) & "cos(" & vBn & ")"
                AddElaborate Sprog.A(225) & " " & san, san & "=" & scn & VBA.ChrW(183) & "cos(" & vBn & ")+" & VBA.ChrW(8730) & "(" & sbn & "^2-" & scn & "^2" & VBA.ChrW(183) & "sin(" & vBn & ")^2)=" & ConvertNumberToStringBC(SA)
                If d > 0 Then AddElaborate Sprog.A(226), san & "_2=" & scn & VBA.ChrW(183) & "cos(" & vBn & ")-" & VBA.ChrW(8730) & "(" & sbn & "^2-" & scn & "^2" & VBA.ChrW(183) & "sin(" & vBn & ")^2)=" & ConvertNumberToStringBC(sa2)
                If sa2 < 0 Then AddElaborate Sprog.A(227), ""
                AddElaborate Sprog.A(215) & " " & vAn & " " & Sprog.A(223), vAn & "=cos^(-1) ((" & sbn & "^2 + " & scn & "^2 - " & san & "^2)/(2" & VBA.ChrW(183) & sbn & VBA.ChrW(183) & scn & "))=cos^(-1) ((" & ConvertNumberToStringBC(sb) & "^2 + " & ConvertNumberToStringBC(sc) & "^2 - " & ConvertNumberToStringBC(SA) & "^2)/(2" & VBA.ChrW(183) & ConvertNumberToStringBC(sb) & VBA.ChrW(183) & ConvertNumberToStringBC(sc) & "))=" & ConvertNumberToStringBC(vA) & VBA.ChrW(176)
                AddElaborate Sprog.A(215) & " " & vCn & " " & Sprog.A(216), vCn & "=180" & VBA.ChrW(176) & "-" & vAn & "-" & vBn & "=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vA) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vC) & VBA.ChrW(176)
                If d > 0 And sa2 > 0.000000000000001 Then
                    vB2 = vB
                    sc2 = sc
                    sb2 = sb
                    vA2 = Arccos((sb2 ^ 2 + sc2 ^ 2 - sa2 ^ 2) / (2 * sb2 * sc2)) * 180 / PI
                    vC2 = 180 - vB2 - vA2
                    AddElaborate vbCrLf & Sprog.A(228) & " " & san & " " & Sprog.A(229), ""
                    AddElaborate Sprog.A(215) & " " & vAn & VBA.ChrW(8322) & " " & Sprog.A(223), vAn & "_2=cos^(-1) ((" & sbn & "^2 + " & scn & "^2 - " & san & "_2^2)/(2" & VBA.ChrW(183) & sbn & "" & VBA.ChrW(183) & scn & "))=cos^(-1) ((" & ConvertNumberToStringBC(sb2) & "^2 + " & ConvertNumberToStringBC(sc2) & "^2 - " & ConvertNumberToStringBC(sa2) & "^2)/(2" & VBA.ChrW(183) & ConvertNumberToStringBC(sb2) & VBA.ChrW(183) & ConvertNumberToStringBC(sc2) & "))=" & ConvertNumberToStringBC(vA2) & VBA.ChrW(176)
                    AddElaborate Sprog.A(215) & " " & vCn & VBA.ChrW(8322) & " " & Sprog.A(216), vCn & "_2=180" & VBA.ChrW(176) & "-" & vAn & "_2-" & vBn & "=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vA2) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB2) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vC2) & VBA.ChrW(176)
                End If
            End If
        Else ' vc>0
            If sb > 0 And SA > 0 Then ' sider om vinkel
                sc = Sqr(sb ^ 2 + SA ^ 2 - 2 * sb * SA * Cos(vC * PI / 180))
                vB = Arccos((sc ^ 2 + SA ^ 2 - sb ^ 2) / (2 * SA * sc)) * 180 / PI
                vA = 180 - vB - vC
                AddElaborate Sprog.A(217) & " " & scn & " " & Sprog.A(223), scn & "=" & VBA.ChrW(8730) & "(" & sbn & "^2 + " & san & "^2 - 2" & VBA.ChrW(183) & sbn & VBA.ChrW(183) & san & VBA.ChrW(183) & "cos(" & vCn & "))=" & VBA.ChrW(8730) & "(" & ConvertNumberToStringBC(sb) & "^2 + " & ConvertNumberToStringBC(SA) & "^2 - 2" & VBA.ChrW(183) & ConvertNumberToStringBC(sb) & VBA.ChrW(183) & ConvertNumberToStringBC(SA) & VBA.ChrW(183) & "cos(" & ConvertNumberToStringBC(vC) & VBA.ChrW(176) & "))=" & ConvertNumberToStringBC(sc)
                AddElaborate Sprog.A(215) & " " & vBn & " " & Sprog.A(223), vBn & "=cos^(-1) ((" & scn & "^2 + " & san & "^2 - " & sbn & "^2)/(2" & VBA.ChrW(183) & scn & VBA.ChrW(183) & san & "))=cos^(-1) ((" & ConvertNumberToStringBC(sc) & "^2 + " & ConvertNumberToStringBC(SA) & "^2 - " & ConvertNumberToStringBC(sb) & "^2)/(2" & VBA.ChrW(183) & ConvertNumberToStringBC(sc) & VBA.ChrW(183) & ConvertNumberToStringBC(SA) & "))=" & ConvertNumberToStringBC(vB) & VBA.ChrW(176)
                AddElaborate Sprog.A(215) & " " & vAn & " " & Sprog.A(216), vAn & "=180" & VBA.ChrW(176) & "-" & vCn & "-" & vBn & "=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vC) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vA) & VBA.ChrW(176)
            ElseIf sc > 0 And sb > 0 Then ' sider ikke om vinkel
                d = sc ^ 2 - sb ^ 2 * Sin(vC * PI / 180) ^ 2
                If d < 0 Then ' ingen løsning
                    GoTo Fejl
                End If
                SA = sb * Cos(vC * PI / 180) + Sqr(d)
                sa2 = sb * Cos(vC * PI / 180) - Sqr(d)
                vB = Arccos((sc ^ 2 + SA ^ 2 - sb ^ 2) / (2 * SA * sc)) * 180 / PI
                vA = 180 - vB - vC
'                sc = sa * Sin(vC * PI / 180) / Sin(vA * PI / 180)
                AddElaborate Sprog.A(217) & " " & san & " " & Sprog.A(223), scn & "^2=" & sbn & "^2+" & san & "^2-2" & sbn & VBA.ChrW(183) & san & VBA.ChrW(183) & "cos(" & vCn & ")"
                AddElaborate Sprog.A(225) & " " & san, san & "=" & sbn & VBA.ChrW(183) & "cos(" & vCn & ")+" & VBA.ChrW(8730) & "(" & scn & "^2-" & sbn & "^2" & VBA.ChrW(183) & "sin(" & vCn & ")^2)=" & ConvertNumberToStringBC(SA)
                If d > 0 Then AddElaborate Sprog.A(226), san & "_2=" & sbn & VBA.ChrW(183) & "cos(" & vCn & ")-" & VBA.ChrW(8730) & "(" & scn & "^2-" & sbn & "^2" & VBA.ChrW(183) & "sin(" & vCn & ")^2)=" & ConvertNumberToStringBC(sa2)
                If sa2 < 0 Then AddElaborate Sprog.A(227), ""
                AddElaborate Sprog.A(215) & " " & vBn & " " & Sprog.A(223), vBn & "=cos^(-1) ((" & scn & "^2 + " & san & "^2 - " & sbn & "^2)/(2" & VBA.ChrW(183) & scn & VBA.ChrW(183) & san & "))=cos^(-1) ((" & ConvertNumberToStringBC(sc) & "^2 + " & ConvertNumberToStringBC(SA) & "^2 - " & ConvertNumberToStringBC(sb) & "^2)/(2" & VBA.ChrW(183) & ConvertNumberToStringBC(sc) & VBA.ChrW(183) & ConvertNumberToStringBC(SA) & "))=" & ConvertNumberToStringBC(vB) & VBA.ChrW(176)
                AddElaborate Sprog.A(215) & " " & vAn & " " & Sprog.A(216), vAn & "=180" & VBA.ChrW(176) & "-" & vCn & "-" & vBn & "=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vC) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vA) & VBA.ChrW(176)
                If d > 0 And sa2 > 0.000000000000001 Then
                    vC2 = vC
                    sb2 = sb
                    sc2 = sc
                    vB2 = Arccos((sc2 ^ 2 + sa2 ^ 2 - sb2 ^ 2) / (2 * sa2 * sc2)) * 180 / PI
                    vA2 = 180 - vB2 - vC2
                    AddElaborate vbCrLf & Sprog.A(228) & " " & san & " " & Sprog.A(229), ""
                    AddElaborate Sprog.A(215) & " " & vBn & VBA.ChrW(8322) & " " & Sprog.A(223), vBn & "_2=cos^(-1) ((" & scn & "^2 + " & san & "_2^2 - " & sbn & "^2)/(2" & VBA.ChrW(183) & scn & "" & VBA.ChrW(183) & san & "_2))=cos^(-1) ((" & ConvertNumberToStringBC(sc2) & "^2 + " & ConvertNumberToStringBC(sa2) & "^2 - " & ConvertNumberToStringBC(sb2) & "^2)/(2" & VBA.ChrW(183) & ConvertNumberToStringBC(sc2) & VBA.ChrW(183) & ConvertNumberToStringBC(sa2) & "))=" & ConvertNumberToStringBC(vB2) & VBA.ChrW(176)
                    AddElaborate Sprog.A(215) & " " & vAn & VBA.ChrW(8322) & " " & Sprog.A(216), vAn & "_2=180" & VBA.ChrW(176) & "-" & vCn & "-" & vBn & "_2=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vC2) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB2) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vA2) & VBA.ChrW(176)
                End If
            ElseIf SA > 0 And sc > 0 Then ' sider ikke om vinkel
                d = sc ^ 2 - SA ^ 2 * Sin(vC * PI / 180) ^ 2
                If d < 0 Then ' ingen løsning
                    GoTo Fejl
                End If
                sb = SA * Cos(vC * PI / 180) + Sqr(d)
                sb2 = SA * Cos(vC * PI / 180) - Sqr(d)
                vB = Arccos((sc ^ 2 + SA ^ 2 - sb ^ 2) / (2 * SA * sc)) * 180 / PI
                vA = 180 - vB - vC
'                sc = sa * Sin(vC * PI / 180) / Sin(vA * PI / 180)
                AddElaborate Sprog.A(217) & " " & sbn & " " & Sprog.A(223), scn & "^2=" & sbn & "^2+" & san & "^2-2" & sbn & VBA.ChrW(183) & san & VBA.ChrW(183) & "cos(" & vCn & ")"
                AddElaborate Sprog.A(225) & " " & sbn, sbn & "=" & san & VBA.ChrW(183) & "cos(" & vCn & ")+" & VBA.ChrW(8730) & "(" & scn & "^2-" & san & "^2" & VBA.ChrW(183) & "sin(" & vCn & ")^2)=" & ConvertNumberToStringBC(sb)
                If d > 0 Then AddElaborate Sprog.A(226), sbn & "_2=" & san & VBA.ChrW(183) & "cos(" & vCn & ")-" & VBA.ChrW(8730) & "(" & scn & "^2-" & san & "^2" & VBA.ChrW(183) & "sin(" & vCn & ")^2)=" & ConvertNumberToStringBC(sb2)
                If sb2 < 0 Then AddElaborate Sprog.A(227), ""
                AddElaborate Sprog.A(215) & " " & vBn & " " & Sprog.A(223), vBn & "=cos^(-1) ((" & scn & "^2 + " & san & "^2 - " & sbn & "^2)/(2" & VBA.ChrW(183) & scn & VBA.ChrW(183) & san & "))=cos^(-1) ((" & ConvertNumberToStringBC(sc) & "^2 + " & ConvertNumberToStringBC(SA) & "^2 - " & ConvertNumberToStringBC(sb) & "^2)/(2" & VBA.ChrW(183) & ConvertNumberToStringBC(sc) & VBA.ChrW(183) & ConvertNumberToStringBC(SA) & "))=" & ConvertNumberToStringBC(vB) & VBA.ChrW(176)
                AddElaborate Sprog.A(215) & " " & vAn & " " & Sprog.A(216), vAn & "=180" & VBA.ChrW(176) & "-" & vCn & "-" & vBn & "=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vC) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vA) & VBA.ChrW(176)
                If d > 0 And sb2 > 0.000000000000001 Then
                    vC2 = vC
                    sa2 = SA
                    sc2 = sc
                    vB2 = Arccos((sc2 ^ 2 + sa2 ^ 2 - sb2 ^ 2) / (2 * sa2 * sc2)) * 180 / PI
                    vA2 = 180 - vB2 - vC2
                    AddElaborate vbCrLf & Sprog.A(228) & " " & sbn & " " & Sprog.A(229), ""
                    AddElaborate Sprog.A(215) & " " & vBn & VBA.ChrW(8322) & " " & Sprog.A(223), vBn & "_2=cos^(-1) ((" & scn & "^2 + " & san & "^2 - " & sbn & "_2^2)/(2" & VBA.ChrW(183) & scn & "" & VBA.ChrW(183) & san & "))=cos^(-1) ((" & ConvertNumberToStringBC(sc2) & "^2 + " & ConvertNumberToStringBC(sa2) & "^2 - " & ConvertNumberToStringBC(sb2) & "^2)/(2" & VBA.ChrW(183) & ConvertNumberToStringBC(sc2) & VBA.ChrW(183) & ConvertNumberToStringBC(sa2) & "))=" & ConvertNumberToStringBC(vB2) & VBA.ChrW(176)
                    AddElaborate Sprog.A(215) & " " & vAn & VBA.ChrW(8322) & " " & Sprog.A(216), vAn & "_2=180" & VBA.ChrW(176) & "-" & vCn & "-" & vBn & "_2=180" & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vC2) & VBA.ChrW(176) & "-" & ConvertNumberToStringBC(vB2) & VBA.ChrW(176) & "=" & ConvertNumberToStringBC(vA2) & VBA.ChrW(176)
                End If
            End If
        End If
    End If
GoTo Slut
Fejl:
    statustext = Sprog.TSMissingInfo
    If advarsler Then MsgBox statustext, vbOKOnly, Sprog.Error
    Exit Sub
Slut:
    If SA <= 0 Or sb <= 0 Or sc <= 0 Or vA <= 0 Or vB <= 0 Or vC <= 0 Then
        GoTo Fejl
    Else
        succes = True
        statustext = Sprog.TSInfoOK
    End If
    If vA2 > 0 Then statustext = statustext & vbCrLf & "(" & Sprog.TS2Solutions2 & ")."

End Sub

'#If Mac Then
'Sub InsertTriangle(r As Range, ByVal vA As Double, ByVal sb As Double, ByVal sc As Double, NameA As String, NameB As String, NameC As String, Namesa As String, Namesb As String, Namesc As String, Anch As Range)
'#Else
Sub InsertTriangle(r As Range, ByVal vA As Double, ByVal sb As Double, ByVal sc As Double, NameA As String, NameB As String, NameC As String, Namesa As String, Namesb As String, Namesc As String)
'#End If

' givet vinkel A og siderne b og c tegner trekanten skaleret
Dim maxs As Double
Dim xmin As Double
Dim nsa As Double
Dim nsb As Double
Dim nsc As Double
Dim xa As Double
Dim ya As Double
Dim xb As Double
Dim yb As Double
Dim xc As Double
Dim yc As Double
Dim F As Double
Dim SA As Double



F = 200

SA = Sqr(sb ^ 2 + sc ^ 2 - 2 * sb * sc * Cos(vA * PI / 180))

If SA <= 0 Or sb <= 0 Or sc <= 0 Then
    MsgBox "Der er sider der er 0", vbOKOnly, Sprog.Error
    GoTo Slut
End If

If SA > sb Then maxs = SA Else maxs = sb
If sc > maxs Then maxs = sc

nsa = SA / maxs * F
nsb = sb / maxs * F
nsc = sc / maxs * F

xb = nsc * Cos(vA * PI / 180)
yb = 0
xa = 0
ya = nsc * Sin(vA * PI / 180)
xc = nsb
yc = ya

If xb < xa Then xmin = -xb
xa = xa + xmin + 10
xb = xb + xmin + 10
xc = xc + xmin + 10

ya = ya + 15
yb = yb + 15
yc = yc + 15


    Dim cv As Shape
    Set cv = ActiveDocument.Shapes.AddCanvas(0, 0, CSng(Maks(xb, xc) + 30), CSng(yc + 30), r)
    cv.WrapFormat.Type = wdWrapInline

    AddLabel NameA, xa - 10, yc, cv  ' yc-5 fjernet for at ikke skal stå oveni figur
    AddLabel NameB, xb - 4, 0, cv
    AddLabel NameC, xc + 5, yc, cv
    
    AddLabel Namesa, (xc + xb) / 2 + 7, yc / 2 - 4, cv
    AddLabel Namesb, (xc + xa) / 2 - 3, yc, cv
    AddLabel Namesc, (xb + xa) / 2 - 10, yc / 2 - 4, cv

    If val(Application.Version) >= 14 Then
        On Error GoTo v12
        cv.CanvasItems.AddConnector msoConnectorStraight, CSng(xa), CSng(ya), CSng(xc), CSng(yc)
        cv.CanvasItems.AddConnector msoConnectorStraight, CSng(xa), CSng(ya), CSng(xb), CSng(yb)
        cv.CanvasItems.AddConnector msoConnectorStraight, CSng(xc), CSng(yc), CSng(xb), CSng(yb)
        cv.Select
        Selection.Cut
        r.Paste
        ClearClipBoard
    Else
v12:
        On Error GoTo Slut
        cv.CanvasItems.AddConnector msoConnectorStraight, CSng(xa), CSng(ya), CSng(xc - xa), 0
        cv.CanvasItems.AddConnector msoConnectorStraight, CSng(xa), CSng(ya), CSng(xb - xa), CSng(yb - ya)
        cv.CanvasItems.AddConnector msoConnectorStraight, CSng(xc), CSng(yc), CSng(xb - xc), CSng(yb - yc)
    End If
Slut:
End Sub

Function AddLabel(Text As String, X As Double, Y As Double, s As Shape) As Shape
    Dim lbl As Shape
    Set lbl = s.CanvasItems.AddLabel(msoTextOrientationHorizontal, CSng(X), CSng(Y), 8, 14)
    lbl.TextFrame.AutoSize = msoTrue
    lbl.TextFrame.WordWrap = False
    lbl.TextFrame.TextRange.Text = Text
    lbl.TextFrame.TextRange.Font.Size = 10
    lbl.TextFrame.MarginBottom = 0
    lbl.TextFrame.MarginTop = 0
    lbl.TextFrame.MarginLeft = 0
    lbl.TextFrame.MarginRight = 0
    lbl.Line.visible = msoFalse
'    lbl.Select
'    Selection.ShapeRange.Fill.Transparency = 0#
    Set AddLabel = lbl
End Function


Private Sub OptionButton_navngivstorlille_Change()
OpdaterNavngivning
End Sub
Private Sub OptionButton_reth_Click()
Dim FN As String
On Error Resume Next
TextBox_C.Text = 90
If CSng(TextBox_A.Text) >= 90 Then TextBox_A.Text = ""
TextBox_C.Enabled = False
TextBox_A.Enabled = True
#If Mac Then
#Else
    FN = GetProgramFilesDir & "\WordMat\Images\trekantreth.emf"
    If Dir(FN) = vbNullString Then FN = Environ("AppData") & "\WordMat\Images\trekantreth.emf"
    If Dir(FN) <> vbNullString Then ImageTrekant.Picture = LoadPicture(FN)
#End If
TextBox_A.Left = 32
TextBox_A.Top = 186
TextBox_B.Left = 318
TextBox_B.Top = 24
TextBox_C.Left = 318
TextBox_C.Top = 174
TextBox_captionA.Left = 48
TextBox_captionA.Top = 174
TextBox_captionB.Left = 300
TextBox_captionB.Top = 24
TextBox_captionC.Left = 300
TextBox_captionC.Top = 174

TextBox_sa.Left = 320
TextBox_sa.Top = 90
TextBox_sb.Left = 151
TextBox_sb.Top = 192
TextBox_sc.Left = 120
TextBox_sc.Top = 90
TextBox_captionsa.Left = 305
TextBox_captionsa.Top = 90
TextBox_captionsb.Left = 162
TextBox_captionsb.Top = 180
TextBox_captionsc.Left = 160
TextBox_captionsc.Top = 90
Me.Repaint

End Sub

Private Sub OptionButton_retv_Click()
Dim FN As String
On Error Resume Next
TextBox_A.Text = 90
If CSng(TextBox_C.Text) >= 90 Then TextBox_C.Text = ""
TextBox_A.Enabled = False
TextBox_C.Enabled = True
#If Mac Then
#Else
    FN = GetProgramFilesDir & "\WordMat\Images\trekantretv.emf"
    If Dir(FN) = vbNullString Then FN = Environ("AppData") & "\WordMat\Images\trekantretv.emf"
    If Dir(FN) <> vbNullString Then ImageTrekant.Picture = LoadPicture(FN)
#End If

TextBox_A.Left = 32
TextBox_A.Top = 186
TextBox_B.Left = 10
TextBox_B.Top = 24
TextBox_C.Left = 318
TextBox_C.Top = 174
TextBox_captionA.Left = 48
TextBox_captionA.Top = 174
TextBox_captionB.Left = 48
TextBox_captionB.Top = 24
TextBox_captionC.Left = 300
TextBox_captionC.Top = 174

TextBox_sa.Left = 195
TextBox_sa.Top = 90
TextBox_sb.Left = 151
TextBox_sb.Top = 192
TextBox_sc.Left = 5
TextBox_sc.Top = 88
TextBox_captionsa.Left = 180
TextBox_captionsa.Top = 90
TextBox_captionsb.Left = 162
TextBox_captionsb.Top = 180
TextBox_captionsc.Left = 41
TextBox_captionsc.Top = 90
Me.Repaint

End Sub

Private Sub OptionButton_vilk_Click()
Dim FN As String
On Error Resume Next
TextBox_A.Enabled = True
TextBox_C.Enabled = True
#If Mac Then
#Else
    FN = GetProgramFilesDir & "\WordMat\Images\trekantvilk.emf"
    If Dir(FN) = vbNullString Then FN = Environ("AppData") & "\WordMat\Images\trekantvilk.emf"
    If Dir(FN) <> vbNullString Then ImageTrekant.Picture = LoadPicture(FN)
#End If

TextBox_A.Left = 32
TextBox_A.Top = 186
TextBox_B.Left = 115
TextBox_B.Top = 12
TextBox_C.Left = 318
TextBox_C.Top = 174
TextBox_captionA.Left = 48
TextBox_captionA.Top = 174
TextBox_captionB.Left = 126
TextBox_captionB.Top = 24
TextBox_captionC.Left = 300
TextBox_captionC.Top = 174

TextBox_sa.Left = 234
TextBox_sa.Top = 84
TextBox_sb.Left = 151
TextBox_sb.Top = 192
TextBox_sc.Left = 38
TextBox_sc.Top = 90
TextBox_captionsa.Left = 216
TextBox_captionsa.Top = 84
TextBox_captionsb.Left = 162
TextBox_captionsb.Top = 180
TextBox_captionsc.Left = 78
TextBox_captionsc.Top = 90
Me.Repaint

End Sub
Static Function Log10(X)
    Log10 = Log(X) / Log(10#)
End Function
Function Arcsin(X As Double)
'Arcsin(X) = Atn(X / Sqr(-X * X + 1))
    If X = 1 Then
        Arcsin = PI / 2
    ElseIf X = -1 Then
        Arcsin = 3 / 2 * PI
    Else
        Arcsin = Atn(X / Sqr(-X * X + 1))
    End If
End Function
Function Arccos(X As Double)
'Arccos(X) = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
    If X = 1 Then
        Arccos = 0
    ElseIf X = -1 Then
        Arccos = PI
    Else
        Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
    End If
End Function

Function Maks(A As Double, b As Double)
    If A < b Then
        Maks = b
    Else
        Maks = A
    End If
End Function
Sub UpdateSolution()
    FindSolutions
    If succes Then
        Label_status.ForeColor = RGB(0, 255, 0)
    Else
        Label_status.ForeColor = RGB(255, 0, 0)
    End If
    Label_status.Caption = statustext
End Sub
Private Sub TextBox_A_Change()
    UpdateSolution
End Sub
Private Sub TextBox_B_Change()
    UpdateSolution
End Sub
Private Sub TextBox_C_Change()
    UpdateSolution
End Sub
Private Sub OptionButton_navngivsiderAB_Change()
OpdaterNavngivning
End Sub
Sub OpdaterNavngivning()
If OptionButton_navngivstorlille.Value = True Then
    TextBox_captionA.Text = VBA.UCase(TextBox_captionA.Text)
    TextBox_captionsa.Text = VBA.LCase(TextBox_captionA.Text)
    TextBox_captionB.Text = VBA.UCase(TextBox_captionB.Text)
    TextBox_captionsb.Text = VBA.LCase(TextBox_captionB.Text)
    TextBox_captionC.Text = VBA.UCase(TextBox_captionC.Text)
    TextBox_captionsc.Text = VBA.LCase(TextBox_captionC.Text)
ElseIf OptionButton_navngivsiderAB.Value = True Then
    TextBox_captionsa.Text = TextBox_captionB.Text & TextBox_captionC.Text
    TextBox_captionsb.Text = TextBox_captionA.Text & TextBox_captionC.Text
    TextBox_captionsc.Text = TextBox_captionA.Text & TextBox_captionB.Text
End If
OptionButton_retv.Caption = TextBox_captionA.Text & " " & Sprog.right
OptionButton_reth.Caption = TextBox_captionC.Text & " " & Sprog.right
End Sub
Private Sub TextBox_captionA_Change()
If OptionButton_navngivstorlille.Value = True Then
    TextBox_captionA.Text = VBA.UCase(TextBox_captionA.Text)
    TextBox_captionsa.Text = VBA.LCase(TextBox_captionA.Text)
ElseIf OptionButton_navngivsiderAB.Value = True Then
    OpdaterNavngivning
End If
OptionButton_retv.Caption = TextBox_captionA.Text & " " & Sprog.right
End Sub

Private Sub TextBox_captionB_Change()
If OptionButton_navngivstorlille.Value = True Then
    TextBox_captionB.Text = VBA.UCase(TextBox_captionB.Text)
    TextBox_captionsb.Text = VBA.LCase(TextBox_captionB.Text)
ElseIf OptionButton_navngivsiderAB.Value = True Then
    OpdaterNavngivning
End If
End Sub

Private Sub TextBox_captionC_Change()
If OptionButton_navngivstorlille.Value = True Then
    TextBox_captionC.Text = VBA.UCase(TextBox_captionC.Text)
    TextBox_captionsc.Text = VBA.LCase(TextBox_captionC.Text)
ElseIf OptionButton_navngivsiderAB.Value = True Then
    OpdaterNavngivning
End If
OptionButton_reth.Caption = TextBox_captionC.Text & " " & Sprog.right
End Sub

Private Sub TextBox_captionsa_Change()
If OptionButton_navngivstorlille.Value = True Then
    TextBox_captionsa.Text = VBA.LCase(TextBox_captionsa.Text)
    TextBox_captionA.Text = VBA.UCase(TextBox_captionsa.Text)
End If
OpdaterNavngivning
End Sub

Private Sub TextBox_captionsb_Change()
If OptionButton_navngivstorlille.Value = True Then
    TextBox_captionsb.Text = VBA.LCase(TextBox_captionsb.Text)
    TextBox_captionB.Text = VBA.UCase(TextBox_captionsb.Text)
End If
OpdaterNavngivning
End Sub

Private Sub TextBox_captionsc_Change()
If OptionButton_navngivstorlille.Value = True Then
    TextBox_captionsc.Text = VBA.LCase(TextBox_captionsc.Text)
    TextBox_captionC.Text = VBA.UCase(TextBox_captionsc.Text)
End If
OpdaterNavngivning
End Sub

Private Sub TextBox_sa_Change()
    UpdateSolution
End Sub
Private Sub TextBox_sb_Change()
    UpdateSolution
End Sub
Private Sub TextBox_sc_Change()
    UpdateSolution
End Sub

Sub AddElaborate(Text As String, lign As String)

    elabotext(elaboindex) = Text
    elabolign(elaboindex) = lign
    
    elaboindex = elaboindex + 1
End Sub

Private Sub UserForm_Activate()
    SaveBackup
    SetCaptions
#If Mac Then
    Frame1.visible = False
#End If
    TextBox_A.Text = TriangleAV
    TextBox_B.Text = TriangleBV
    TextBox_C.Text = TriangleCV
    TextBox_sa.Text = TriangleAS
    TextBox_sb.Text = TriangleBS
    TextBox_sc.Text = TriangleCS
    
    If TriangleSett1 = 1 Then
        OptionButton_retv.Value = True
    ElseIf TriangleSett1 = 2 Then
        OptionButton_reth.Value = True
    Else
        OptionButton_vilk.Value = True
    End If
    If TriangleSett2 = 1 Then
        OptionButton_navngivmanuel.Value = True
    ElseIf TriangleSett2 = 2 Then
        OptionButton_navngivstorlille.Value = True
    Else
        OptionButton_navngivsiderAB.Value = True
    End If
    

    If TriangleNAS = "" And TriangleNBS = "" And TriangleNCS = "" And TriangleNAV = "" And TriangleNBV = "" And TriangleNCV = "" Then
     TriangleNAS = "A"
     TriangleNBS = "B"
    TriangleNCS = "C"
    TriangleNAV = "a"
    TriangleNBV = "b"
    TriangleNCV = "c"
    TriangleSett1 = 3
    TriangleSett2 = 2
    TriangleSett3 = False
    TriangleSett4 = False
    End If
    TextBox_captionA.Text = TriangleNAV
    TextBox_captionB.Text = TriangleNBV
    TextBox_captionC.Text = TriangleNCV
    TextBox_captionsa.Text = TriangleNAS
    TextBox_captionsb.Text = TriangleNBS
    TextBox_captionsc.Text = TriangleNCS
    CheckBox_tal.Value = TriangleSett3
    CheckBox_forklaring.Value = TriangleSett4
    
    OpdaterNavngivning
End Sub

Private Sub SaveSettings()
    TriangleAV = TextBox_A.Text
    TriangleBV = TextBox_B.Text
    TriangleCV = TextBox_C.Text
    TriangleAS = TextBox_sa.Text
    TriangleBS = TextBox_sb.Text
    TriangleCS = TextBox_sc.Text
    TriangleNAV = TextBox_captionA.Text
    TriangleNBV = TextBox_captionB.Text
    TriangleNCV = TextBox_captionC.Text
    TriangleNAS = TextBox_captionsa.Text
    TriangleNBS = TextBox_captionsb.Text
    TriangleNCS = TextBox_captionsc.Text
    TriangleSett3 = CheckBox_tal.Value
    TriangleSett4 = CheckBox_forklaring.Value
    If OptionButton_retv.Value Then
        TriangleSett1 = 1
    ElseIf OptionButton_reth.Value = True Then
        TriangleSett1 = 2
    Else
        TriangleSett1 = 3
    End If
    If OptionButton_navngivmanuel.Value = True Then
        TriangleSett2 = 1
    ElseIf OptionButton_navngivstorlille.Value = True Then
        TriangleSett2 = 2
    Else
        TriangleSett2 = 3
    End If

End Sub


Private Sub SetCaptions()
    Me.Caption = Sprog.TriangleSolver
    Label_ok.Caption = Sprog.OK
    Frame1.Caption = Sprog.RightAngled & "?"
    Frame2.Caption = Sprog.Naming
    OptionButton_navngivmanuel.Caption = Sprog.Manuel
    OptionButton_navngivstorlille.Caption = Sprog.AngleNaming1
    OptionButton_navngivsiderAB.Caption = Sprog.AngleNaming2
    CheckBox_tal.Caption = Sprog.InsertNumbers
    CheckBox_forklaring.Caption = Sprog.ShowCalculations
    Label1.Caption = Sprog.TriangleSolverExplanation1
    Label2.Caption = Sprog.TriangleSolverExplanation2
    OptionButton_vilk.Caption = Sprog.AnyTriangle
    Label_nulstil.Caption = Sprog.Clear
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    SaveSettings
End Sub

Private Sub Label_ok_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorPress
End Sub

Private Sub Label_ok_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorHover
End Sub
Private Sub Label_cancel_Click()
    Me.Hide
    Application.ScreenUpdating = False
End Sub

Private Sub Label_cancel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorPress
End Sub

Private Sub Label_cancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_cancel.BackColor = LBColorHover
End Sub
Private Sub Label_nulstil_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_nulstil.BackColor = LBColorPress
End Sub

Private Sub Label_nulstil_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_nulstil.BackColor = LBColorHover
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label_ok.BackColor = LBColorInactive
    Label_cancel.BackColor = LBColorInactive
    Label_nulstil.BackColor = LBColorInactive
End Sub


