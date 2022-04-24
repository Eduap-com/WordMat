Attribute VB_Name = "TestModul"
Option Explicit
    Private visok As Boolean
    Private UfWait2 As UserFormWaitForMaxima
    Private ErrCount As Long
    Private TestCount As Long

Sub TestTabel()
    UserFormTest.Show
End Sub
Sub RunTestSequence()
' runs a series of test calculations. The expressions are typed into Word and the test math action is performed. Result is shown
'    Application.ScreenUpdating = False
    Dim scrollpos As Double
    
    On Error GoTo fejl
    ErrCount = 0
    TestCount = 0
    visok = True
    
    If MsgBox("Are you sure want to conduct a test. The document will be filled with calculations. It can take some time.", vbOKCancel, "Confirm") = vbCancel Then Exit Sub
    
#If Mac Then
#Else
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
#End If
    
    Set UfWait2 = New UserFormWaitForMaxima
    UfWait2.Label1.Font.Size = 12
    UfWait2.Label_tip.Font.Size = 10
    UfWait2.Show vbModeless
    
    CASengine = 0
    MaximaExact = 1 ' 0 - auto ' 1 - exact ' 2 - num
    MaximaCifre = 7
    Radians = False
    
    PrepareMaxima
    
    MoveCursorToEndOfCalculation
    
    Selection.TypeParagraph
    Selection.Font.Bold = True
    Selection.Font.Size = 14
    Selection.TypeText "Test results:"
    Selection.Font.Bold = False
    Selection.Font.Size = 11
    ShowSettings
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.GoToPrevious (wdGoToLine)
    Selection.GoToPrevious (wdGoToLine)


    ' brug UnicodeValsToString for at finde streng fra et matematikfelt. Både til kommando og resultat
    ' Insert equation, solve it. Select all an execute 'UnicodeValsToString'. Then copy the string for command and result
    ' the result may have to be adjusted. Often brackets of exponents are different
'    DoEvents ' virker ikke
'    Application.ScreenRefresh
    TestBeregn "2+3", "=5"
    If StopNow Then GoTo slut
    TestBeregn "1/5 2", "=2/5"
    If StopNow Then GoTo slut
    TestBeregn "a b/c+a b+f_a (x)", "=f_a (x)+(a" & VBA.ChrW(183) & "b)/c+a" & VBA.ChrW(183) & "b"
    If StopNow Then GoTo slut
    TestBeregn "f_a (x)+a (b+d)/c+f_c+2+a^x (2)", "=f_a (x)+f_c+(a" & VBA.ChrW(183) & "(d+b))/c+2" & VBA.ChrW(183) & "a^(x)+2"
    If StopNow Then GoTo slut
    TestBeregn "2+3^3,4/log" & VBA.ChrW(8289) & "(889) -sin" & VBA.ChrW(8289) & "(34)", "=(-sin" & VBA.ChrW(8289) & "((17" & VBA.ChrW(183) & "" & VBA.ChrW(960) & ")/90) )+(41,89983049571472" & VBA.ChrW(183) & "ln" & VBA.ChrW(8289) & "(10))/ln" & VBA.ChrW(8289) & "(889) +2"
    If StopNow Then GoTo slut
    TestSolve "40=72" & VBA.ChrW(183) & "e^((0,619/0,22" & VBA.ChrW(8729) & "(e^22-e^0,22t )) )", "t", "t=e^22" & VBA.ChrW(183) & "0,8025188+ln" & VBA.ChrW(8289) & "(5/9)" & VBA.ChrW(183) & "-0,2852248"
    If StopNow Then GoTo slut
    TestSolve "x^2=9", "x", "x=-3    " & VBA.ChrW(8744) & "    x=3"
    If StopNow Then GoTo slut
    TestSolve VBA.ChrW(8747) & "_0^a" & VBA.ChrW(9618) & "x^2 dx=4", "a", "a=12^(1/3)"
    If StopNow Then GoTo slut
    TestSolve VBA.ChrW(8747) & "_0^2" & VBA.ChrW(9618) & "" & VBA.ChrW(12310) & "a·x·(x-2)^2 " & VBA.ChrW(12311) & " dx=20", "a", "a=15"
    If StopNow Then GoTo slut
    TestSolve "92,3*(e^(0,21278*x)-1)/(e^(0,21278*x)+1)=0.5", "x", "x=0,05091805"
    If StopNow Then GoTo slut
    TestSolve "x^3-x^2+x-2=0", "x", "x=1,35321"
    If StopNow Then GoTo slut
    TestSolve "sin" & VBA.ChrW(8289) & "(2x)=1/2", "x", "x=15"
    If StopNow Then GoTo slut
    TestSolve "340=34" & VBA.ChrW(183) & "x^2,9", "x", "x=2,212216"
    If StopNow Then GoTo slut
    TestSolve "sin" & VBA.ChrW(8289) & "(x)=2", "x", "x" & VBA.ChrW(8712) & "" & VBA.ChrW(8709)
    If StopNow Then GoTo slut
    TestSolve "x^(1/3)=V", "x", "x=V^3"
    If StopNow Then GoTo slut
    TestSolve "5/(x-1)=7/x", "x", "x=3,5"
    If StopNow Then GoTo slut
    TestSolve "x-x=2", "x", "x" & VBA.ChrW(8712) & "" & VBA.ChrW(8709)
    If StopNow Then GoTo slut
    TestSolve "x" & VBA.ChrW(183) & "e^2x+e^2x=0", "x", "x=-1    " & VBA.ChrW(8744) & "    x=0"
    If StopNow Then GoTo slut
    TestSolve "x^2" & VBA.ChrW(183) & "(x^2-1)=9" & VBA.ChrW(183) & "(x^2-1)", "x", "x=-3    " & VBA.ChrW(8744) & "    x=-1    " & VBA.ChrW(8744) & "    x=1    " & VBA.ChrW(8744) & "    x=3"
    If StopNow Then GoTo slut
    TestSolve "x^4-x^3+x^2-x-1=0", "x", "x=-0,5187901    " & VBA.ChrW(8744) & "    x=1,290649"
    If StopNow Then GoTo slut
    InsertTestMath "T_2" & VBA.ChrW(8788) & "8"
    Selection.TypeParagraph
    TestSolve "T_2=log" & VBA.ChrW(8289) & "(2)/log" & VBA.ChrW(8289) & "(a) ", "a", "a=2^(1/8)"
    If StopNow Then GoTo slut
    InsertSletDef
'    TestSolve "", "x", "x="
'    If StopNow Then GoTo slut
'    TestSolve "", "x", "x="
    
    ' kan kun løses med geogebra og nsolve maxima, så kræver bruger input ved UI
    TestSolve "x^2=" & VBA.ChrW(12310) & "0,7" & VBA.ChrW(12311) & "^x", "x", "x" & VBA.ChrW(8776) & "-15,29371    " & VBA.ChrW(8744) & "    x" & VBA.ChrW(8776) & "-1,249643    " & VBA.ChrW(8744) & "    x" & VBA.ChrW(8776) & "0,8581024", "Click maxima numeric"
    
    TestBeregn VBA.ChrW(8747) & "_(-" & VBA.ChrW(8734) & ")^" & VBA.ChrW(8734) & "" & VBA.ChrW(9618) & "1/(" & VBA.ChrW(8730) & "2" & VBA.ChrW(960) & "*a-*3)*a-*e^(-1/2*a-*((y-1)/3)^2 ) dy", "=1"
    
    GoTo slut
fejl:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error at test no.:" & TestCount + 1, vbOKOnly, "Error during test"
slut:
    Selection.TypeParagraph
    Selection.Font.Bold = True
    Selection.Font.Size = 14
    Selection.TypeText ("Test complete!   ")
    Selection.TypeText "Error count: " & ErrCount & "/" & TestCount
    Selection.Font.Bold = False
    Selection.Font.Size = 11
    Selection.TypeParagraph
    Selection.TypeText "Don't forget to do manual test, as the UI isn't tested using the testmodule"
    
#If Mac Then
#Else
        Oundo.EndCustomRecord
#End If
    Unload UfWait2
End Sub
Function StopNow() As Boolean
    If omax.StopNow Then
        StopNow = True
        Exit Function
    End If
    If ErrCount = 5 Then
        If MsgBox("5 errors detected. Do you want to continue?", vbYesNo, "Many errors") = vbNo Then
            StopNow = True
            Exit Function
        End If
    ElseIf ErrCount = 10 Then
        If MsgBox("10 errors detected. Do you still want to continue?", vbYesNo, "Many errors") = vbNo Then
            StopNow = True
            Exit Function
        End If
    End If
    StopNow = False
End Function
Sub PerformTest(TestType As Integer, komm As String, resul As String, Optional var As String, Optional Instruk As String)
    Dim s As String, TypeText As String, Oresul As String
    If TestType = 1 Then
        TypeText = "Calculating"
    ElseIf TestType = 2 Then
        TypeText = "Solving equation"
    End If
    s = TestCount & ": " & TypeText & vbCrLf & "Error count: " & ErrCount
    UfWait2.Label1.Caption = s
    UfWait2.Label_tip.Caption = komm
    If Instruk <> "" Then
        UfWait2.Label_tip.Caption = komm & vbCrLf & Instruk
    End If
    InsertTestMath komm
    If TestType = 1 Then
        beregn
    ElseIf TestType = 2 Then
        MaximaSolvePar (var)
    End If
    MoveCursorToEndOfCalculation
    
    If TestType = 1 Then ' ved beregn skrives resultatet sammen med input, så sammenligning ryger
        Oresul = omax.MaximaOutput
    Else
        Do While Selection.OMaths.Count = 0
            Selection.GoToPrevious (wdGoToLine)
            Selection.EndKey Unit:=wdLine
        Loop
        omax.ReadSelection
        Oresul = TrimR(omax.Kommando, vbCr)
        MoveCursorToEndOfCalculation False
    End If
    
    If Not Oresul = resul Then 'omax.MaximaOutput
        Selection.Font.ColorIndex = wdRed
        Selection.Font.Bold = True
        Selection.TypeText ("Error " & TypeText & ". Expected result: ")
        Selection.Font.Bold = False
        Selection.Font.ColorIndex = wdAuto
        Selection.TypeText ("  ")
        InsertTestMath resul
        UnicodeValsToString
        If TestType = 1 Then
            Selection.TypeParagraph
            Selection.TypeText omax.MaximaOutput
        End If
        Selection.TypeParagraph
        ErrCount = ErrCount + 1
    ElseIf visok Then
        Selection.Font.ColorIndex = wdGreen
        Selection.Font.Bold = True
        Selection.TypeText ("Test OK: " & TypeText)
        Selection.Font.Bold = False
        Selection.Font.ColorIndex = wdAuto
        Selection.TypeParagraph
    End If
    UfWait2.Label_progress.Caption = UfWait2.Label_progress.Caption & "*"
    TestCount = TestCount + 1
End Sub
Sub TestBeregn(komm As String, resul As String)
    PerformTest 1, komm, resul
End Sub

Sub TestSolve(komm As String, var As String, resul As String, Optional Instruk As String)
    PerformTest 2, komm, resul, var, Instruk
    Exit Sub
    Dim s As String
    s = TestCount & ": Solving equation" & vbCrLf & "Error count: " & ErrCount
    UfWait2.Label1.Caption = s
    UfWait2.Label_tip.Caption = komm
    If Instruk <> "" Then
        UfWait2.Label_tip.Caption = komm & vbCrLf & Instruk
    End If
    InsertTestMath komm
    MaximaSolvePar (var)
    MoveCursorToEndOfCalculation
    If Not omax.MaximaOutput = resul Then
        Selection.Font.ColorIndex = wdRed
        Selection.Font.Bold = True
        Selection.TypeText ("Error solving equation. Expected result: ")
        Selection.Font.Bold = False
        Selection.Font.ColorIndex = wdAuto
        Selection.TypeText ("  ")
        InsertTestMath resul
        UnicodeValsToString
        Selection.TypeParagraph
        ErrCount = ErrCount + 1
    ElseIf visok Then
        Selection.Font.ColorIndex = wdGreen
        Selection.Font.Bold = True
        Selection.TypeText ("Equation solved correctly")
        Selection.Font.Bold = False
        Selection.Font.ColorIndex = wdAuto
        Selection.TypeParagraph
    End If
    UfWait2.Label_progress.Caption = UfWait2.Label_progress.Caption & "*"
    TestCount = TestCount + 1
End Sub

Sub TestSolve2(komm As String, var As String, resul As String)
    omax.Kommando = komm
    omax.MaximaSolve (var)
    If Not omax.MaximaOutput = resul Then
        Selection.TypeText (" - Fejl - " & omax.Kommando & " - " & omax.MaximaOutput)
        Selection.TypeParagraph
    ElseIf visok Then
        Selection.TypeText (" - ok - " & omax.Kommando)
        Selection.TypeParagraph
    End If

End Sub

Sub testGetListItem()
Dim ea As New ExpressionAnalyser

ea.text = "dette er[ad;sdfs] en test ; hej(a;b{1;2}) ;{a;d} hallo"

MsgBox ea.GetNextListItem()
MsgBox ea.GetNextListItem()
MsgBox ea.GetNextListItem()
MsgBox ea.GetNextListItem()
MsgBox ea.GetNextListItem()

End Sub
Sub InsertTestMath(s As String)
' inds*ae*tter maxima output i word document
    Dim mo As Range

    Selection.Range.Font.Bold = False

    If Len(s) > 500 Then
        s = Left(s, 500) + Sprog.A(119)
    End If
    Set mo = Selection.OMaths.Add(Selection.Range)
    Selection.TypeText s
    mo.OMaths.BuildUp
End Sub
Sub MoveCursorToEndOfCalculation(Optional AddLine As Boolean = True)
    If Not Selection.Find.Execute("*^13^13", , , True) Then
'        MsgBox "ingen ny linje"
        '    Selection.GoToNext (wdGoToLine)
        Selection.EndKey Unit:=wdLine
        Selection.Collapse wdCollapseEnd
        Selection.TypeParagraph
        Selection.TypeParagraph
        Selection.GoToPrevious (wdGoToLine)
    End If
    Selection.Collapse wdCollapseEnd
    If AddLine Then
        Selection.TypeParagraph
        Selection.GoToPrevious (wdGoToLine)
    End If
    Selection.GoToPrevious (wdGoToLine)
End Sub
Sub ShowSettings()
    Selection.TypeParagraph
    Selection.TypeText "CAS engine: " & CASengine
    Selection.TypeParagraph
    Selection.TypeText "Exact: " & MaximaExact
    Selection.TypeParagraph
    Selection.TypeText "Cifre: " & MaximaCifre
    Selection.TypeParagraph
    Selection.TypeText "Radians: " & Radians
    Selection.TypeParagraph
    
End Sub
Sub InsertTestComment(s As String)
    Selection.TypeText (s)
    Selection.TypeParagraph
End Sub
Sub ShowDef()
Dim s As String
Dim i As Integer

For i = 0 To omax.defindex - 1
    s = s & omax.DefName(i) & " - " & omax.DefValue(i) & vbCrLf
Next

MsgBox s

End Sub


Sub OpenJanis()
Dim ReturnValue
Dim sti As String
    sti = """" & GetProgramFilesDir & "\WordMat\janis.jnlp"""

ReturnValue = Shell(sti, vbNormalFocus)
End Sub
