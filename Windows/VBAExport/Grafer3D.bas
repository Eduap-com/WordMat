Attribute VB_Name = "Grafer3D"
Option Explicit
Sub OmdrejningsLegeme()
Dim Kommando As String
    Dim fktnavn As String, Udtryk As String, LHS As String, RHS As String, varnavn As String, fktudtryk As String
Dim Arr As Variant
Dim i As Integer, UrlLink As String, cmd As String, j As Integer
    Dim DefList As String

    Dim ea As New ExpressionAnalyser
'    Dim ea2 As New ExpressionAnalyser
    
    ea.SetNormalBrackets
'    ea2.SetNormalBrackets

'On Error GoTo fejl

#If Mac Then
'    UrlLink = "file:///Library/Application%20Support/Microsoft/Office365/User%20Content.localized/Add-Ins.localized/WordMat/geogebra-math-apps/GeoGebra/HTML5/5.0/GeoGebra.html"
    UrlLink = "file://" & GetGeoGebraMathAppsFolder() & "GeoGebra3dApplet.html"
#Else
'    UrlLink = "https://geogebra.org/calculator"
'    UrlLink = "file:///C:/Program%20Files%20(x86)/WordMat/geogebra-math-apps/GeoGebraApplet.html"
    UrlLink = "file://" & GetGeoGebraMathAppsFolder() & "GeoGebra3dApplet.html"
#End If
    UrlLink = UrlLink & "?command="
    PrepareMaxima
    omax.ConvertLnLog = False
    omax.ReadSelection
    
    
    
        ' inds�t de markerede funktioner
    For i = 0 To omax.KommandoArrayLength
        Udtryk = omax.KommandoArray(i)
        Udtryk = Replace(Udtryk, "definer:", "")
        Udtryk = Replace(Udtryk, "Definer:", "")
        Udtryk = Replace(Udtryk, "define:", "")
        Udtryk = Replace(Udtryk, "Define:", "")
        Udtryk = Replace(Udtryk, VBA.ChrW(8788), "=") ' :=
        Udtryk = Replace(Udtryk, VBA.ChrW(8797), "=") ' tripel =
        Udtryk = Replace(Udtryk, VBA.ChrW(8801), "=") ' def =
        Udtryk = Trim(Udtryk)
        If Len(Udtryk) > 0 Then
            If InStr(Udtryk, "matrix") < 1 Then ' matricer og vektorer er ikke implementeret endnu
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
                ' kan f�rst bruges med GeoGebra 4.0
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
'                    If Trim(Udtryk) = "x" Then 'line�re funktioner kan plottes implicit og bliver meget p�nere
'                        cmd = "z^2=(" & Replace(ConvertToGeogebraSyntax(Udtryk), "+", "%2B") & ")^2-y^2" & ";"
'                        UrlLink = UrlLink & cmd
'                    Else
'                        cmd = "z=sqrt((" & Replace(ConvertToGeogebraSyntax(Udtryk), "+", "%2B") & ")^2-y^2)" & ";"
'                        UrlLink = UrlLink & cmd
'                        cmd = "z=-sqrt((" & Replace(ConvertToGeogebraSyntax(Udtryk), "+", "%2B") & ")^2-y^2)" & ";"
'                        UrlLink = UrlLink & cmd
'                    End If

'                    geogebrafil.CreateFunction "f" & j, udtryk, False
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
            UserFormOmdrejninglegeme.TextBox_forskrift.Text = Kommando
        ElseIf Len(Kommando) > 0 And i = 1 Then
            UserFormOmdrejninglegeme.TextBox_forskrift2.Text = Kommando
        End If
        i = i + 1
    Loop
    
    Application.ScreenUpdating = True
    
    UserFormOmdrejninglegeme.Show
    
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
    
    '   Set UF2Dgraph = New UserForm2DGraph
    forskrifter = omax.FindDefinitions
    If Len(forskrifter) > 3 Then
'        forskrifter = Mid(forskrifter, 2, Len(forskrifter) - 3) 'fjernet 1.33
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
    forskrifter = omax.KommandoerStreng & "#$" & forskrifter
    
    If Len(forskrifter) > 1 Then
        Arr = Split(forskrifter, "#$")
        For i = 0 To UBound(Arr)
            Arr(i) = Replace(Arr(i), " ", "")
            If Arr(i) <> "" Then
                If MsgBox2(Sprog.A(374) & ": " & Arr(i) & " ?", vbYesNo, Sprog.A(375) & "?") = vbYes Then
                    Insert3DEquation (Arr(i))
                End If
            End If
        Next
    End If
    
    UserForm3DGraph.Show
    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub Insert3DEquation(Equation As String)
    Dim LHS As String, RHS As String, Arr() As String
    
    If Equation = vbNullString Then Exit Sub
    Arr = Split(Equation, "=")
    LHS = LCase(Replace(Replace(Trim(Arr(0)), " ", ""), ";", ","))
    If UBound(Arr) > 0 Then RHS = Arr(1)

If InStr(Equation, "=") > 0 And LHS <> "z" And LHS <> "f(x,y)" Then
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
ElseIf InStr(Equation, VBA.ChrW(9632)) Then
    Equation = Replace(Equation, VBA.ChrW(9632), "")
    Equation = Replace(Equation, "@", ",")
    Equation = Replace(Equation, "((", "(")
    Equation = Replace(Equation, "))", ")")
    Equation = "(0,0,0)-" & Equation
    If UserForm3DGraph.TextBox_vektorer.Text <> "" Then
        If right(UserForm3DGraph.TextBox_vektorer.Text, 1) = ")" Then
            UserForm3DGraph.TextBox_vektorer.Text = UserForm3DGraph.TextBox_vektorer.Text & vbCr
        End If
    End If
    UserForm3DGraph.TextBox_vektorer.Text = UserForm3DGraph.TextBox_vektorer.Text & Equation
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

