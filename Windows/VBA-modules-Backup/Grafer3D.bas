Attribute VB_Name = "Grafer3D"
Option Explicit
Sub OmdrejningsLegeme()
    Dim geogebrafil As New CGeoGebraFile
Dim Kommando As String
    Dim fktnavn As String, udtryk As String, lhs As String, rhs As String, varnavn As String, fktudtryk As String
Dim Arr As Variant
Dim i As Integer, UrlLink As String, Cmd As String, j As Integer
    Dim var As String, DefList As String

    Dim ea As New ExpressionAnalyser
'    Dim ea2 As New ExpressionAnalyser
    
    ea.SetNormalBrackets
'    ea2.SetNormalBrackets

'On Error GoTo fejl

#If Mac Then
'    UrlLink = "file:///Library/Application%20Support/Microsoft/Office365/User%20Content.localized/Add-Ins.localized/WordMat/geogebra-math-apps/GeoGebra/HTML5/5.0/GeoGebra.html"
    UrlLink = "file:///Library/Application%20Support/Microsoft/Office365/User%20Content.localized/Add-Ins.localized/WordMat/geogebra-math-apps/GeoGebra3dApplet.html"
#Else
'    UrlLink = "https://geogebra.org/calculator"
'    UrlLink = "file:///C:/Program%20Files%20(x86)/WordMat/geogebra-math-apps/GeoGebraApplet.html"
    UrlLink = "file://" & GetProgramFilesDir & "/WordMat/geogebra-math-apps/GeoGebra3dApplet.html"
#End If
    UrlLink = UrlLink & "?command="
    PrepareMaxima
    omax.ConvertLnLog = False
    omax.ReadSelection
    
    
    
        ' indsæt de markerede funktioner
    For i = 0 To omax.KommandoArrayLength
        udtryk = omax.KommandoArray(i)
        udtryk = Replace(udtryk, "definer:", "")
        udtryk = Replace(udtryk, "Definer:", "")
        udtryk = Replace(udtryk, "define:", "")
        udtryk = Replace(udtryk, "Define:", "")
        udtryk = Replace(udtryk, VBA.ChrW(8788), "=") ' :=
        udtryk = Replace(udtryk, VBA.ChrW(8797), "=") ' tripel =
        udtryk = Replace(udtryk, VBA.ChrW(8801), "=") ' def =
        udtryk = Trim(udtryk)
        If Len(udtryk) > 0 Then
            If InStr(udtryk, "matrix") < 1 Then ' matricer og vektorer er ikke implementeret endnu
                If InStr(udtryk, "=") > 0 Then
                    Arr = Split(udtryk, "=")
                    lhs = Arr(0)
                    rhs = Arr(1)
                    ea.text = lhs
                    fktnavn = ea.GetNextVar(1)
                    varnavn = ea.GetNextBracketContent(1)
                    
                    If lhs = fktnavn & "(" & varnavn & ")" Then
                        ea.text = rhs
                        ea.pos = 1
                        ea.ReplaceVar varnavn, "x"
                        fktudtryk = ea.text
                        DefinerKonstanter fktudtryk, DefList, Nothing, UrlLink
                        
                        Cmd = "z^2=(" & Replace(ConvertToGeogebraSyntax(fktudtryk), "+", "%2B") & ")^2-y^2" & ";"
                        UrlLink = UrlLink & Cmd

                    Else
                        fktudtryk = ReplaceIndepvarX(rhs)
                        DefinerKonstanter udtryk, DefList, Nothing, UrlLink
                        Cmd = "z^2=(" & Replace(ConvertToGeogebraSyntax(fktudtryk), "+", "%2B") & ")^2-y^2" & ";"
                        UrlLink = UrlLink & Cmd
                        j = j + 1
                    End If
                ElseIf InStr(udtryk, ">") > 0 Or InStr(udtryk, "<") > 0 Or InStr(udtryk, VBA.ChrW(8804)) > 0 Or InStr(udtryk, VBA.ChrW(8805)) > 0 Then
                ' kan først bruges med GeoGebra 4.0
                    DefinerKonstanter udtryk, DefList, Nothing, UrlLink
                    Cmd = Replace(ConvertToGeogebraSyntax(Cmd), "+", "%2B") & ";"
                    Cmd = "z^2=(" & Replace(ConvertToGeogebraSyntax(udtryk), "+", "%2B") & ")^2-y^2" & ";"
                    UrlLink = UrlLink & Cmd
'                    geogebrafil.CreateFunction "u" & j, udtryk, True
                Else
                    udtryk = ReplaceIndepvarX(udtryk)
                    udtryk = Replace(udtryk, vbCrLf, "")
                    udtryk = Replace(udtryk, vbCr, "")
                    udtryk = Replace(udtryk, vbLf, "")
                    DefinerKonstanter udtryk, DefList, Nothing, UrlLink
                    If Trim(udtryk) = "x" Then 'lineære funktioner kan plottes implicit og bliver meget pænere
                        Cmd = "z^2=(" & Replace(ConvertToGeogebraSyntax(udtryk), "+", "%2B") & ")^2-y^2" & ";"
                        UrlLink = UrlLink & Cmd
                    Else
                        Cmd = "z=sqrt((" & Replace(ConvertToGeogebraSyntax(udtryk), "+", "%2B") & ")^2-y^2)" & ";"
                        UrlLink = UrlLink & Cmd
                        Cmd = "z=-sqrt((" & Replace(ConvertToGeogebraSyntax(udtryk), "+", "%2B") & ")^2-y^2)" & ";"
                        UrlLink = UrlLink & Cmd
                    End If

'                    geogebrafil.CreateFunction "f" & j, udtryk, False
                    j = j + 1
                End If
            End If
        End If
    Next
    
'    UrlLink = UrlLink & "z^2=(" & Replace(geogebrafil.ConvertToGeoGebraSyntax(omax.Kommando), "+", "%2B") & ")^2-y^2"
    'omax.CodeForMaxima(omax.Kommando)
    
    OpenLink UrlLink, True

Exit Sub

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
            UserFormOmdrejninglegeme.TextBox_forskrift.text = Kommando
        ElseIf Len(Kommando) > 0 And i = 1 Then
            UserFormOmdrejninglegeme.TextBox_forskrift2.text = Kommando
        End If
        i = i + 1
    Loop
    
    Application.ScreenUpdating = True
    
    UserFormOmdrejninglegeme.Show
    
fejl:
slut:
End Sub

Sub Plot3DGraph()
    Dim forskrifter As String
    Dim Result As Variant
    Dim Arr As Variant
    Dim i As Integer
    On Error GoTo fejl
    
    PrepareMaxima
    omax.ReadSelection
    
'   Set UF2Dgraph = New UserForm2DGraph
    forskrifter = omax.FindDefinitions
    If Len(forskrifter) > 3 Then
    forskrifter = Mid(forskrifter, 2, Len(forskrifter) - 3)
    Arr = Split(forskrifter, ListSeparator)
    forskrifter = ""
    
    For i = 0 To UBound(Arr)
        If InStr(Arr(i), "):") > 0 Then
            forskrifter = forskrifter & omax.ConvertToWordSymbols(Arr(i)) & ListSeparator
        End If
    Next
    End If
    
    If forskrifter <> "" Then
        forskrifter = Left(forskrifter, Len(forskrifter) - 1)
    End If
    forskrifter = omax.KommandoerStreng & ListSeparator & forskrifter
    
    If Len(forskrifter) > 1 Then
    Arr = Split(forskrifter, ListSeparator)
    For i = 0 To UBound(Arr)
        Arr(i) = Replace(Arr(i), " ", "")
        If Arr(i) <> "" Then
            If MsgBox(Sprog.A(374) & ": " & Arr(i) & " ?", vbYesNo, Sprog.A(375) & "?") = vbYes Then
                Insert3DEquation (Arr(i))
            End If
        End If
    Next
    End If
    
    UserForm3DGraph.Show
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub Insert3DEquation(Equation As String)

If InStr(Equation, "=") > 0 Then
    If UserForm3DGraph.TextBox_ligning1.text = Equation Then Exit Sub
    If UserForm3DGraph.TextBox_ligning2.text = Equation Then Exit Sub
    If UserForm3DGraph.TextBox_ligning3.text = Equation Then Exit Sub
    If UserForm3DGraph.TextBox_ligning1.text = "" Then
        UserForm3DGraph.TextBox_ligning1.text = Equation
    ElseIf UserForm3DGraph.TextBox_ligning2.text = "" Then
        UserForm3DGraph.TextBox_ligning2.text = Equation
    ElseIf UserForm3DGraph.TextBox_ligning3.text = "" Then
        UserForm3DGraph.TextBox_ligning3.text = Equation
    End If
ElseIf InStr(Equation, VBA.ChrW(9632)) Then
    Equation = Replace(Equation, VBA.ChrW(9632), "")
    Equation = Replace(Equation, "@", ",")
    Equation = Replace(Equation, "((", "(")
    Equation = Replace(Equation, "))", ")")
    Equation = "(0,0,0)-" & Equation
    If UserForm3DGraph.TextBox_vektorer.text <> "" Then
        If right(UserForm3DGraph.TextBox_vektorer.text, 1) = ")" Then
            UserForm3DGraph.TextBox_vektorer.text = UserForm3DGraph.TextBox_vektorer.text & vbCr
        End If
    End If
    UserForm3DGraph.TextBox_vektorer.text = UserForm3DGraph.TextBox_vektorer.text & Equation
Else
    If UserForm3DGraph.TextBox_forskrift1.text = Equation Then Exit Sub
    If UserForm3DGraph.TextBox_forskrift2.text = Equation Then Exit Sub
    If UserForm3DGraph.TextBox_forskrift3.text = Equation Then Exit Sub
    If UserForm3DGraph.TextBox_forskrift1.text = "" Then
         UserForm3DGraph.TextBox_forskrift1.text = Equation
    ElseIf UserForm3DGraph.TextBox_forskrift2.text = "" Then
         UserForm3DGraph.TextBox_forskrift2.text = Equation
    ElseIf UserForm3DGraph.TextBox_forskrift3.text = "" Then
         UserForm3DGraph.TextBox_forskrift3.text = Equation
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

