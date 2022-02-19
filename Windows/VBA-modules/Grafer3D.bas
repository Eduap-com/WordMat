Attribute VB_Name = "Grafer3D"
Option Explicit
Sub OmdrejningsLegeme()
Dim Kommando As String
Dim arr As Variant
Dim i As Integer
On Error GoTo fejl
    PrepareMaxima
    omax.ReadSelection
    i = 0
    Do While i < omax.KommandoArrayLength + 1
        Kommando = omax.KommandoArray(i)
        arr = Split(Kommando, "=")
        If Len(Kommando) > 0 Then Kommando = arr(UBound(arr))
        
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
    
fejl:
    Application.ScreenUpdating = True
    UserFormOmdrejninglegeme.Show

End Sub

Sub Plot3DGraph()
    Dim forskrifter As String
    Dim Result As Variant
    Dim arr As Variant
    Dim i As Integer
    On Error GoTo fejl
    
    PrepareMaxima
    omax.ReadSelection
    
'   Set UF2Dgraph = New UserForm2DGraph
    forskrifter = omax.FindDefinitions
    If Len(forskrifter) > 3 Then
    forskrifter = Mid(forskrifter, 2, Len(forskrifter) - 3)
    arr = Split(forskrifter, ListSeparator)
    forskrifter = ""
    
    For i = 0 To UBound(arr)
        If InStr(arr(i), "):") > 0 Then
            forskrifter = forskrifter & omax.ConvertToWordSymbols(arr(i)) & ListSeparator
        End If
    Next
    End If
    
    If forskrifter <> "" Then
        forskrifter = Left(forskrifter, Len(forskrifter) - 1)
    End If
    forskrifter = omax.KommandoerStreng & ListSeparator & forskrifter
    
    If Len(forskrifter) > 1 Then
    arr = Split(forskrifter, ListSeparator)
    For i = 0 To UBound(arr)
        arr(i) = Replace(arr(i), " ", "")
        If arr(i) <> "" Then
            If MsgBox(Sprog.A(374) & ": " & arr(i) & " ?", vbYesNo, Sprog.A(375) & "?") = vbYes Then
                Insert3DEquation (arr(i))
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

