Attribute VB_Name = "Statistik"
Option Explicit

Sub Chi2Fordeling()
    Dim k As Integer
    Dim g As Variant
    Dim ut As String
    Dim kstring As String
    
    kstring = InputBox(Sprog.A(398), Sprog.A(360), "n")
    k = val(kstring)
    If (k > 0 And k <= 20) Or (k > 20 And Int(k / 2) = k / 2) Then
        g = Gamma(k / 2)
        ut = ("f(x)" & VBA.ChrW(8797) & ConvertNumberToString(1 / (2 ^ (k / 2) * g)) & ChrW(183) & "x^(" & k / 2 - 1 & ")" & ChrW(183) & "e^(-x/2)")
        ut = Replace(ut, ",", ".")
    Else
        ut = "f(x)" & VBA.ChrW(8801) & "1/(2^(n/2)" & ChrW(183) & VBA.ChrW(915) & "(n/2) )" & ChrW(183) & "x^(n/2-1)" & ChrW(183) & "e^(-x/2)"
    End If
    Selection.InsertAfter (VBA.ChrW(&H3C7) & VBA.ChrW(&HB2) & " - " & Sprog.A(399) & " " & kstring & " " & Sprog.A(360))
    Selection.Collapse (wdCollapseEnd)
    Selection.TypeParagraph
    
    Selection.InsertAfter ut
'    Selection.InsertAfter ("p(x)=1/(2^(" & k & "/2)á" & G & ")áx^(" & k & "/2-1)áe^(-x/2)")
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).BuildUp
    Selection.MoveRight unit:=wdCharacter, Count:=2

End Sub

Function Gamma(z As Variant) As Variant
' begrænset implementation af gammafunktion
    If z = Int(z) Then ' hvis z er heltal er det bare fakultet
        Gamma = Factorial(z - 1)
    ElseIf z = 0.5 Then ' for specielle halvtallige er løsningen kendt
        Gamma = 1.77245384774943
    ElseIf z = 1.5 Then
        Gamma = 0.88622692387471
    ElseIf z = 2.5 Then
        Gamma = 1.32934038581207
    ElseIf z = 3.5 Then
        Gamma = 3.32335102350265
    ElseIf z = 4.5 Then
        Gamma = 11.6317284109978
    ElseIf z = 5.5 Then
        Gamma = 52.3427781617139
    ElseIf z = 6.5 Then
        Gamma = 287.885281385281
    ElseIf z = 7.5 Then
        Gamma = 1871.25431034483
    ElseIf z = 8.5 Then
        Gamma = 14034.4074074074
    ElseIf z = 9.5 Then
        Gamma = 119292.461538462
    Else ' Andre kan jeg ikke
        Gamma = 0
    End If
    
End Function

Function Factorial(n)
    If n <= 1 Then    ' Reached end of recursive calls.
        Factorial = 1    ' (N = 0) so climb back out of calls.
    Else    ' Call Factorial again if N > 0.
        Factorial = Factorial(n - 1) * n
    End If
End Function

Function Lgamma(z As Variant) As Variant
' begrænset implementation af Lower gammafunktion
    If z = Int(z) Then ' hvis z er heltal er det bare faktor
        Lgamma = Factorial(z)
    ElseIf z = 0.5 Then ' for specielle halvtallige er løsningen kendt
        Lgamma = 1.77245384774943
    ElseIf z = 1.5 Then
        Lgamma = 0.88622692387471
    ElseIf z = 2.5 Then
        Lgamma = 1.32934038581207
    ElseIf z = 3.5 Then
        Lgamma = 3.32335102350265
    ElseIf z = 4.5 Then
        Lgamma = 11.6317284109978
    ElseIf z = 5.5 Then
        Lgamma = 52.3427781617139
    ElseIf z = 6.5 Then
        Lgamma = 287.885281385281
    ElseIf z = 7.5 Then
        Lgamma = 1871.25431034483
    ElseIf z = 8.5 Then
        Lgamma = 14034.4074074074
    ElseIf z = 9.5 Then
        Lgamma = 119292.461538462
    Else ' Andre kan jeg ikke
        Lgamma = 0
    End If
    
End Function


