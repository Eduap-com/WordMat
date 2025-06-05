Attribute VB_Name = "Grafer3D"
Option Explicit
Sub OmdrejningsLegeme()
    Dim Kommando As String
    Dim fktnavn As String, Udtryk As String, LHS As String, RHS As String, varnavn As String, fktudtryk As String
    Dim Arr As Variant
    Dim i As Integer, UrlLink As String, cmd As String, j As Integer
    Dim DefList As String

    Dim ea As New ExpressionAnalyser
    
    ea.SetNormalBrackets

    'On Error GoTo fejl

#If Mac Then
    UrlLink = "file://" & GetGeoGebraMathAppsFolder() & "GeoGebra3dApplet.html"
#Else
    UrlLink = "file://" & GetGeoGebraMathAppsFolder() & "GeoGebra3dApplet.html"
#End If
    UrlLink = UrlLink & "?command="
    PrepareMaxima
    omax.ConvertLnLog = False
    omax.ReadSelection
    
    
    ' Insert selected functions
    For i = 0 To omax.KommandoArrayLength
        Udtryk = omax.KommandoArray(i)
        Udtryk = Replace(Udtryk, "definer:", "")
        Udtryk = Replace(Udtryk, "Definer:", "")
        Udtryk = Replace(Udtryk, "define:", "")
        Udtryk = Replace(Udtryk, "Define:", "")
        Udtryk = Replace(Udtryk, VBA.ChrW(8788), "=") ' :=
        Udtryk = Replace(Udtryk, VBA.ChrW(8797), "=") ' triple =
        Udtryk = Replace(Udtryk, VBA.ChrW(8801), "=") ' def =
        Udtryk = Trim(Udtryk)
        If Len(Udtryk) > 0 Then
            If InStr(Udtryk, "matrix") < 1 Then
                If InStr(Udtryk, "=") > 0 Then
                    Arr = Split(Udtryk, "=")
                    LHS = Arr(0)
                    RHS = Arr(1)
                    ea.Text = LHS
                    fktnavn = ea.GetNextVar(1)
                    varnavn = ea.GetNextBracketContent(1)
                    
                    If LHS = fktnavn & "(" & varnavn & ")" Then
                        ea.Text = RHS
                        ea.Pos = 1
                        ea.ReplaceVar varnavn, "x"
                        fktudtryk = ea.Text
                        DefinerKonstanter fktudtryk, DefList, Nothing, UrlLink
                        
                        cmd = "surface(" & Replace(ConvertToGeogebraSyntax(fktudtryk), "+", "%2B") & ",2*pi);"
                        '                        cmd = "z^2=(" & Replace(ConvertToGeogebraSyntax(fktudtryk), "+", "%2B") & ")^2-y^2" & ";"
                        UrlLink = UrlLink & cmd

                    Else
                        fktudtryk = ReplaceIndepvarX(RHS)
                        DefinerKonstanter fktudtryk, DefList, Nothing, UrlLink
                        cmd = "surface(" & Replace(ConvertToGeogebraSyntax(fktudtryk), "+", "%2B") & ",2*pi);"
                        '                        cmd = "z^2=(" & Replace(ConvertToGeogebraSyntax(fktudtryk), "+", "%2B") & ")^2-y^2" & ";"
                        UrlLink = UrlLink & cmd
                        j = j + 1
                    End If
                ElseIf InStr(Udtryk, ">") > 0 Or InStr(Udtryk, "<") > 0 Or InStr(Udtryk, VBA.ChrW(8804)) > 0 Or InStr(Udtryk, VBA.ChrW(8805)) > 0 Then
                    ' can only be used with GeoGebra 4.0
                    DefinerKonstanter Udtryk, DefList, Nothing, UrlLink
                    cmd = Replace(ConvertToGeogebraSyntax(cmd), "+", "%2B") & ";"
                    cmd = "z^2=(" & Replace(ConvertToGeogebraSyntax(Udtryk), "+", "%2B") & ")^2-y^2" & ";"
                    UrlLink = UrlLink & cmd
                    '                    geogebrafil.CreateFunction "u" & j, udtryk, True
                Else
                    Udtryk = ReplaceIndepvarX(Udtryk)
                    Udtryk = Replace(Udtryk, vbCrLf, "")
                    Udtryk = Replace(Udtryk, vbCr, "")
                    Udtryk = Replace(Udtryk, vbLf, "")
                    DefinerKonstanter Udtryk, DefList, Nothing, UrlLink
                    cmd = "surface(" & Replace(ConvertToGeogebraSyntax(Udtryk), "+", "%2B") & ",2*pi);"
                    UrlLink = UrlLink & cmd
                    j = j + 1
                End If
            End If
        End If
    Next
    
    '    UrlLink = UrlLink & "z^2=(" & Replace(geogebrafil.ConvertToGeoGebraSyntax(omax.Kommando), "+", "%2B") & ")^2-y^2"
    'omax.CodeForMaxima(omax.Kommando)
    
    OpenLink UrlLink, True

    Exit Sub '******************************************

    PrepareMaxima
    omax.ReadSelection
    i = 0
    Do While i < omax.KommandoArrayLength + 1
        Kommando = omax.KommandoArray(i)
        Arr = Split(Kommando, "=")
        If Len(Kommando) > 0 Then Kommando = Arr(UBound(Arr))
        
        Kommando = Replace(Kommando, vbLf, "")
        Kommando = Replace(Kommando, vbCrLf, "")
        Kommando = Replace(Kommando, vbCr, "")
        Kommando = Replace(Kommando, " ", "")
        Kommando = omax.ConvertToWordSymbols(Kommando)
        Kommando = Replace(Kommando, ";", ".")
        If Len(Kommando) > 0 And i = 0 Then
            UserFormSolidOfRevolution.TextBox_forskrift.Text = Kommando
        ElseIf Len(Kommando) > 0 And i = 1 Then
            UserFormSolidOfRevolution.TextBox_forskrift2.Text = Kommando
        End If
        i = i + 1
    Loop
    
    Application.ScreenUpdating = True
    
    UserFormSolidOfRevolution.Show
    
Fejl:
slut:
End Sub

Sub Plot3DGraph()
    Dim forskrifter As String
    Dim Arr As Variant
    Dim i As Integer
    On Error GoTo Fejl
    
    PrepareMaxima
    omax.ReadSelection
    
    forskrifter = omax.FindDefinitions
    If Len(forskrifter) > 3 Then
'        forskrifter = Mid(forskrifter, 2, Len(forskrifter) - 3) 'removed 1.33
        Arr = Split(forskrifter, ListSeparator)
        forskrifter = ""
    
        For i = 0 To UBound(Arr)
            If InStr(Arr(i), "):") > 0 Then
                Arr(i) = Replace(Arr(i), ":=", "=")
                forskrifter = forskrifter & omax.ConvertToWordSymbols(Arr(i)) & "#$"
            End If
        Next
    End If
    
    If forskrifter <> "" Then
        forskrifter = Left(forskrifter, Len(forskrifter) - 2)
    End If
    
    For i = 0 To omax.KommandoArrayLength
        forskrifter = Trim(LCase(omax.KommandoArray(i))) & "#$" & forskrifter 'omax.KommandoerStreng
    Next
    
    
    
    If Len(forskrifter) > 1 Then
        Arr = Split(forskrifter, "#$")
        For i = 0 To UBound(Arr)
            Arr(i) = Replace(Arr(i), " ", "")
            If Arr(i) <> "" Then
'                If MsgBox2(TT.A(374) & ": " & Arr(i) & " ?", vbYesNo, TT.A(375) & "?") = vbYes Then
                    Insert3DEquation (Arr(i))
'                End If
            End If
        Next
    End If
    
    UserForm3DGraph.Show
    GoTo slut
Fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
End Sub

Sub Insert3DEquation(Equation As String)
    Dim LHS As String, RHS As String, Arr() As String, p As Integer
    Dim ea As New ExpressionAnalyser, s As String
    Dim tbx As TextBox, tby As TextBox, tbz As TextBox, tbtmin As TextBox, tbtmax As TextBox, tbsmin As TextBox, tbsmax As TextBox
    Dim px As String, py As String, pz As String
    ea.SetNormalBrackets
    If Equation = vbNullString Then Exit Sub
    
    Equation = Replace(Equation, ChrW(9632), ChrW(9608)) ' two different symbols used for vectors. Otherwise syntax is the same
    Arr = Split(Equation, "=")
    LHS = LCase(Replace(Replace(Trim(Arr(0)), " ", ""), ";", ","))
    If UBound(Arr) > 0 Then RHS = Arr(1)

    p = InStr(LHS, ChrW(9608)) ' vector symbol

    If p > 0 Then  ' vector or parametric plot
        If RHS <> vbNullString Then ' if RHS exist only RHS is used
            LHS = RHS
        End If
        If InStr(LHS, ChrW(166)) > 0 Then ' vector input by template stacks. cannot be used for 3d plot, but can be combined with normal for problematic input
            ea.Text = LHS
            s = ea.GetNextBracketContent
            If InStr(s, "¦") > 0 Then
                Arr = Split(s, "¦")
                If UBound(Arr) = 2 Then
                    px = px & Arr(0)
                    py = py & Arr(1)
                    pz = py & Arr(2)
                End If
            End If
        Else ' Normal vector input
            ea.Text = LHS
            s = ea.GetNextBracketContent(p)
            Arr = Split(s, "@")
            If UBound(Arr) = 2 Then
                px = Arr(0)
                py = Arr(1)
                pz = Arr(2)
            End If
        End If
        If InStr(LHS, "t") > 0 Then ' if t in expression it is probably parametric plot
            If UserForm3DGraph.TextBox_parametric1x = vbNullString Then
                Set tbx = UserForm3DGraph.TextBox_parametric1x
                Set tby = UserForm3DGraph.TextBox_parametric1y
                Set tbz = UserForm3DGraph.TextBox_parametric1z
                Set tbtmin = UserForm3DGraph.TextBox_tmin1
                Set tbtmax = UserForm3DGraph.TextBox_tmax1
                Set tbsmin = UserForm3DGraph.TextBox_smin1
                Set tbsmax = UserForm3DGraph.TextBox_smax1
            ElseIf UserForm3DGraph.TextBox_parametric2x = vbNullString Then
                Set tbx = UserForm3DGraph.TextBox_parametric2x
                Set tby = UserForm3DGraph.TextBox_parametric2y
                Set tbz = UserForm3DGraph.TextBox_parametric2z
                Set tbtmin = UserForm3DGraph.TextBox_tmin2
                Set tbtmax = UserForm3DGraph.TextBox_tmax2
                Set tbsmin = UserForm3DGraph.TextBox_smin2
                Set tbsmax = UserForm3DGraph.TextBox_smax2
            ElseIf UserForm3DGraph.TextBox_parametric3x = vbNullString Then
                Set tbx = UserForm3DGraph.TextBox_parametric3x
                Set tby = UserForm3DGraph.TextBox_parametric3y
                Set tbz = UserForm3DGraph.TextBox_parametric3z
                Set tbtmin = UserForm3DGraph.TextBox_tmin3
                Set tbtmax = UserForm3DGraph.TextBox_tmax3
                Set tbsmin = UserForm3DGraph.TextBox_smin3
                Set tbsmax = UserForm3DGraph.TextBox_smax3
            End If
            If Not tbx Is Nothing Then
                tbx.Text = px
                tby.Text = py
                tbz.Text = pz
            End If
            If tbtmin.Text = vbNullString Then
                If InStr(px, "t") > 0 Or InStr(py, "t") > 0 Or InStr(pz, "t") > 0 Then
                    tbtmin.Text = "0"
                    tbtmax.Text = "1"
                End If
            End If
        Else ' vector
            Equation = "(0" & ListSeparator & "0" & ListSeparator & "0)(" & px & ListSeparator & " " & py & ListSeparator & " " & pz & ")"
            If UserForm3DGraph.TextBox_vektorer.Text <> "" Then
                If Right(UserForm3DGraph.TextBox_vektorer.Text, 1) = ")" Then
                    UserForm3DGraph.TextBox_vektorer.Text = UserForm3DGraph.TextBox_vektorer.Text & vbCr
                End If
            End If
            UserForm3DGraph.TextBox_vektorer.Text = UserForm3DGraph.TextBox_vektorer.Text & Equation
        End If
    ElseIf InStr(Equation, "=") > 0 And LHS <> "z" And LHS <> "f(x,y)" Then
        If UserForm3DGraph.TextBox_ligning1.Text = Equation Then Exit Sub
        If UserForm3DGraph.TextBox_ligning2.Text = Equation Then Exit Sub
        If UserForm3DGraph.TextBox_ligning3.Text = Equation Then Exit Sub
        If UserForm3DGraph.TextBox_ligning1.Text = "" Then
            UserForm3DGraph.TextBox_ligning1.Text = Equation
        ElseIf UserForm3DGraph.TextBox_ligning2.Text = "" Then
            UserForm3DGraph.TextBox_ligning2.Text = Equation
        ElseIf UserForm3DGraph.TextBox_ligning3.Text = "" Then
            UserForm3DGraph.TextBox_ligning3.Text = Equation
        End If
    Else
        If UserForm3DGraph.TextBox_forskrift1.Text = RHS Then Exit Sub
        If UserForm3DGraph.TextBox_forskrift2.Text = RHS Then Exit Sub
        If UserForm3DGraph.TextBox_forskrift3.Text = RHS Then Exit Sub
        If UserForm3DGraph.TextBox_forskrift1.Text = "" Then
            UserForm3DGraph.TextBox_forskrift1.Text = RHS
        ElseIf UserForm3DGraph.TextBox_forskrift2.Text = "" Then
            UserForm3DGraph.TextBox_forskrift2.Text = RHS
        ElseIf UserForm3DGraph.TextBox_forskrift3.Text = "" Then
            UserForm3DGraph.TextBox_forskrift3.Text = RHS
        End If
    End If

End Sub

Function GetNextColor() As String
colindex = colindex + 1
If colindex = 1 Then
    GetNextColor = "black"
ElseIf colindex = 2 Then
    GetNextColor = "green"
ElseIf colindex = 3 Then
    GetNextColor = "red"
ElseIf colindex = 4 Then
    GetNextColor = "blue"
ElseIf colindex = 5 Then
    GetNextColor = "cyan"
ElseIf colindex = 6 Then
    GetNextColor = "magenta"
Else
    GetNextColor = "black"
    colindex = 1
End If

End Function

