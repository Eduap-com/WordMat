Attribute VB_Name = "TestModul"
Option Explicit
    Private visok As Boolean
    Private UFwait2 As UserFormWaitForMaxima
    Private ErrCount As Long
    Private TestCount As Long
    Private ContCount As Integer
    Private NonInterA As Boolean

' Use CreateTestBeregn to create new tests

Sub RunTestSequenceNonInteractive()
    NonInterA = True
    RunTestSequence
End Sub
Sub RunTestSequenceInteractive()
    NonInterA = False
    RunTestSequence
End Sub

Sub RunTestSequence()
    ' runs a series of test calculations. The expressions are typed into Word and the test math action is performed. Result is shown
    '    Application.ScreenUpdating = False
    Dim s As String, i As Integer
    Dim AllR As Range
    Dim Tid As Single
    
    On Error GoTo fejl
    ErrCount = 0
    TestCount = 0
    ContCount = 0
    visok = True
    
    If MsgBox("Are you sure want to conduct a test. The document will be filled with calculations. It can take some time.", vbOKCancel, "Confirm") = vbCancel Then Exit Sub
    Tid = Timer
        
    Set UFwait2 = New UserFormWaitForMaxima
    UFwait2.Label1.Font.Size = 12
    UFwait2.Label_tip.Font.Size = 10
    UFwait2.Show vbModeless
    
' The test results are based on certain settings. They may change along the way, but they should be reset initially.
    CASengineTempOnly = 0 ' 0=maxima
    CASengine = 0
    MaximaExact = 1 ' 0 - auto ' 1 - exact ' 2 - num
    MaximaCifre = 7
    MaximaDecOutType = 2 ' bc
    Radians = False
    MaximaComplex = False
    AllTrig = False
    MaximaLogOutput = 0
    MaximaIndex = 0 ' subscript is just text. Otherwise will cause problems when x_1 is defined
    SettUseVBACAS = True
    
    
    PrepareMaxima
    
    Selection.Collapse wdCollapseEnd
    Selection.GoToNext (wdGoToLine)
    If Selection.OMaths.Count > 0 Then MoveCursorToEndOfCalculation
    
    Set AllR = Selection.Range
    
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
    
    Dim WT As Template
    Set WT = GetWordMatTemplate(False)
    If Not WT Is Nothing Then
        s = CheckKeyboardShortcutsNoninteractive()
        If s = "" Then
            Selection.Font.ColorIndex = wdGreen
            Selection.Font.Bold = True
            Selection.TypeText ("Keyboard Shortcuts ok")
            Selection.Font.Bold = False
            Selection.Font.ColorIndex = wdAuto
            Selection.TypeParagraph
        Else
            Selection.Font.ColorIndex = wdRed
            Selection.Font.Bold = True
            Selection.TypeText ("Keyboard shortcut problem:  " & s)
            Selection.Font.Bold = False
            Selection.Font.ColorIndex = wdAuto
            Selection.TypeParagraph
            '        Selection.TypeText ("  ")
            ErrCount = ErrCount + 1
        End If
        TestCount = TestCount + 1
    Else
            Selection.TypeText ("Could not check Keyboard shortcuts, because running as global template")
            Selection.TypeParagraph
    End If
    'GoTo ggbtest
    
    ' when creating new tests, it is easiest to use the function 'CreateTestBeregn'
    ' alternative: just running the test with an empty expected result. It is written which test string should be used if the result is correct.
    ' ellers brug GetTestString' eller UnicodeValsToString for at finde streng fra et matematikfelt for kommandoen.
    ' the result may have to be adjusted for calculations (not solve). Often brackets of exponents are different and spaces must be removed
    ' Multiple correct answers can be separated by @$
    '    DoEvents ' virker ikke
    '    Application.ScreenRefresh
    
    'til test af enkelt
'    GoTo slut
    
    
    DebugWM = False
    ' not the prettiest solution, but empty lines are eaten and if there is no space for calculations to grow in, it causes problems.
    For i = 0 To 45 ' must be same as below. Increase if adding more tests
        Selection.TypeParagraph
    Next
    Selection.MoveLeft wdCharacter, 45

    InsertSletDef
    ' You can enter multiple correct results, just separate with @$ ----------------------------------------
    
    'VBACAS test
    SettUseVBACAS = True
    
    If TestBeregn("2+3", "=5") Then GoTo slut
    If TestBeregn("1/5 2", "=2/5") Then GoTo slut
    
    ' Maxima test
    SettUseVBACAS = False

    'calculation tests
    TestBeregn "2+3^3,4/log" & ChrW$(8289) & "(889) -sin" & ChrW$(8289) & "(34)", "41,89983/log" & ChrW$(8289) & "(889) -sin" & ChrW$(8289) & "(34)+2"
    'TestBeregn "2+3^3,4/log" & ChrW$(8289) & "(889) -sin" & ChrW$(8289) & "(34)", "-sin" & ChrW$(8289) & "((17" & ChrW$(183) & "" & ChrW$(960) & ")/90)+(41,89983" & ChrW$(183) & "ln" & ChrW$(8289) & "(10))/ln" & ChrW$(8289) & "(889) +2"
    'TestBeregn "2+3^3,4/log" & chrw$(8289) & "(889) -sin" & chrw$(8289) & "(34)", "=(-sin" & chrw$(8289) & "((17" & chrw$(183) & "" & chrw$(960) & ")/90))+(41,89983049571472" & chrw$(183) & "ln" & chrw$(8289) & "(10))/ln" & chrw$(8289) & "(889)+2@$=-sin" & chrw$(8289) & "((17" & chrw$(183) & "" & chrw$(960) & ")/90)+(41,899830495714724" & chrw$(183) & "ln" & chrw$(8289) & "(10))/ln" & chrw$(8289) & "(889)+2"
    If StopNow Then GoTo slut
    TestBeregn ChrW$(8731) & "(-8)", "=-2"
    If StopNow Then GoTo slut
    TestBeregn "a b/c+a b+f_a (x)", "f_a (x)+(b" & ChrW$(183) & "a)/c+b" & ChrW$(183) & "a@$f_a (x)+(a" & ChrW$(183) & "b)/c+a" & ChrW$(183) & "b"
    If StopNow Then GoTo slut
    TestBeregn "f_a (x)+a (b+d)/c+f_c+2+a^x (2)", "f_a (x)+2" & ChrW$(183) & "a^x+((b+d)" & ChrW$(183) & "a)/c+f_c+2@$f_a (x)+(a" & ChrW$(183) & "(d+b))/c+f_c+2" & ChrW$(183) & "a^x+2"
    If StopNow Then GoTo slut
    'TestBeregn "2^2x+23/2x", "=2^(2" & ChrW$(183) & "x)+23/(2" & ChrW$(183) & "x)" ' fails if 2^2x is not interpreted as 2^(2*x)
    TestBeregn "2^2x+23/2x", "23/(2" & ChrW$(183) & "x)+2^(2" & ChrW$(183) & "x)" ' fails if 2^2x is not interpreted as 2^(2*x)
    If StopNow Then GoTo slut
    TestBeregn "log" & ChrW$(8289) & "(a)", "log" & ChrW$(8289) & "(a)"
    'TestBeregn "log" & ChrW$(8289) & "(a)", "=ln" & ChrW$(8289) & "(a)/ln" & ChrW$(8289) & "(10)"
    If StopNow Then GoTo slut
    TestBeregn "log_2" & ChrW$(8289) & "(4)", "=2"
    If StopNow Then GoTo slut
    TestBeregn "ln" & ChrW$(8289) & "(a)", "=ln" & ChrW$(8289) & "(a)"
    If StopNow Then GoTo slut
    TestBeregn "log_4" & ChrW$(8289) & "a", "=ln" & ChrW$(8289) & "(a)/ln" & ChrW$(8289) & "(4)"
    If StopNow Then GoTo slut
    TestBeregn ChrW$(12310) & "sin" & ChrW$(8289) & "(x)-sin" & ChrW$(12311) & "" & ChrW$(8289) & "(x_0 )/(x+y)", "(sin" & ChrW$(8289) & "(x)-sin" & ChrW$(8289) & "(x_0 ))/(x+y)"
    'TestBeregn ChrW$(12310) & "sin" & ChrW$(8289) & "(x)-sin" & ChrW$(12311) & "" & ChrW$(8289) & "(x_0 )/(x+y)", "=(sin" & ChrW$(8289) & "((" & ChrW$(960) & "" & ChrW$(183) & "x)/180)-sin" & ChrW$(8289) & "((" & ChrW$(960) & "" & ChrW$(183) & "x_0)/180))/(y+x)" ' Test af forkert placerede skjulte parenteser
    If StopNow Then GoTo slut
    TestBeregn ChrW$(12310) & "sin" & ChrW$(8289) & "(x)-sin" & ChrW$(12311) & "" & ChrW$(8289) & "(x_0 )/(x-x_0 )", "(sin" & ChrW$(8289) & "(x)-sin" & ChrW$(8289) & "(x_0 ))/(x-x_0 )"
    'TestBeregn ChrW$(12310) & "sin" & ChrW$(8289) & "(x)-sin" & ChrW$(12311) & "" & ChrW$(8289) & "(x_0 )/(x-x_0 )", "=(sin" & ChrW$(8289) & "((" & ChrW$(960) & "" & ChrW$(183) & "x)/180)-sin" & ChrW$(8289) & "((" & ChrW$(960) & "" & ChrW$(183) & "x_0)/180))/(x-x_0)"
    If StopNow Then GoTo slut
    TestBeregn "x/" & ChrW$(12310) & "2+x" & ChrW$(12311) & "^2", "=x/(x^(2)+2)"
    If StopNow Then GoTo slut
    TestBeregn "|2+(1+|2-3|)/(|2|+1)+1|", "=11/3"
    If StopNow Then GoTo slut
    TestBeregn "L=100^(1/5)/" & ChrW$(960) & "^(1/5)", "=100^(1/5)/" & ChrW$(960) & "^(1/5)"
    If StopNow Then GoTo slut
    TestBeregn "cos^(-1)" & ChrW$(8289) & "(2)", "cos^(-1)" & ChrW$(8289) & "(2)"
    'TestBeregn "cos^(-1)" & ChrW$(8289) & "(2)", "=(180" & ChrW$(183) & "cos^(-1)" & ChrW$(8289) & "(2))/" & ChrW$(960)
    If StopNow Then GoTo slut
    TestBeregn "(" & ChrW$(8730) & "((x+1)^2+(2" & ChrW$(183) & "(x+1)/x)^2 ))^'", "(x^4+x^3-4" & ChrW$(183) & "x-4)/(x" & ChrW$(183) & "" & ChrW$(8730) & "(x^4+2" & ChrW$(183) & "x^3+5" & ChrW$(183) & "x^2+8" & ChrW$(183) & "x+4)" & ChrW$(183) & "|x| )"
    'TestBeregn "(" & ChrW$(8730) & "((x+1)^2+(2" & ChrW$(183) & "(x+1)/x)^2 ))^'", "=((-(8" & ChrW$(183) & "(x+1)^(2))/x^(3))+(8" & ChrW$(183) & "(x+1))/x^(2)+2" & ChrW$(183) & "(x+1))/(2" & ChrW$(183) & "" & ChrW$(8730) & "((4" & ChrW$(183) & "(x+1)^(2))/x^(2)+(x+1)^(2)))@$=(-((8" & ChrW$(183) & "(x+1)^(2))/x^(3))+(8" & ChrW$(183) & "(x+1))/x^(2)+2" & ChrW$(183) & "(x+1))/(2" & ChrW$(183) & "" & ChrW$(8730) & "((4" & ChrW$(183) & "(x+1)^(2))/x^(2)+(x+1)^(2)))"
    If StopNow Then GoTo slut
    If TestBeregn("cos^(-1)" & ChrW$(8289) & "(1/2)", "=60") Then GoTo slut
    If StopNow Then GoTo slut
    If TestBeregn(ChrW$(12310) & "cos" & ChrW$(12311) & "^(-1) (0,5)", "=188,4956/" & ChrW$(960)) Then GoTo slut ' Skrevet vha. Design-Skrift-Hævet skrift.
    If StopNow Then GoTo slut
    
    InsertSletDef
    ' equation solving
    InsertHeadingtext "Equation solving"
    TestSolve "x^2=9", "x", "x=-3    " & ChrW$(8744) & "    x=3"
    If StopNow Then GoTo slut
    TestSolve "Hovedstol=Ydelse" & ChrW$(8729) & "" & ChrW$(12310) & "1-(1+r)" & ChrW$(12311) & "^(-n)/r", "n", "n=-ln" & ChrW$(8289) & "((Ydelse-Hovedstol" & ChrW$(183) & "r)/Ydelse)/ln" & ChrW$(8289) & "(r+1) @$n=-(ln" & ChrW$(8289) & "((Ydelse-Hovedstol" & ChrW$(183) & "r)/Ydelse)/ln" & ChrW$(8289) & "(r+1) )"
    If StopNow Then GoTo slut
    TestSolve ChrW$(8747) & "_0^a" & ChrW$(9618) & "x^2 dx=4", "a", "a=12^(1/3)"
    If StopNow Then GoTo slut
    TestSolve ChrW$(8747) & "_0^2" & ChrW$(9618) & "" & ChrW$(12310) & "a·x·(x-2)^2 " & ChrW$(12311) & " dx=20", "a", "a=15"
    If StopNow Then GoTo slut
    TestSolve "92,3*(e^(0,21278*x)-1)/(e^(0,21278*x)+1)=0.5", "x", "x=0,05091805"
    If StopNow Then GoTo slut
    TestSolve "x^3-x^2+x-2=0", "x", "x=1,35321"
    If StopNow Then GoTo slut
    TestSolve "x^4-x^3+x^2-x-1=0", "x", "x=-0,5187901    " & ChrW$(8744) & "    x=1,290649"
    If StopNow Then GoTo slut
    TestSolve "x-x=2", "x", "x" & ChrW$(8712) & "" & ChrW$(8709)
    If StopNow Then GoTo slut
    TestSolve "sin" & ChrW$(8289) & "(2x)=1/2", "x", "x=15"
    If StopNow Then GoTo slut
    TestSolve "340=34" & ChrW$(183) & "x^2,9", "x", "x=2,212216"
    If StopNow Then GoTo slut
    TestSolve "sin" & ChrW$(8289) & "(x)=2", "x", "x" & ChrW$(8712) & "" & ChrW$(8709)
    If StopNow Then GoTo slut
    TestSolve "x^(1/3)=V", "x", "x=V^3"
    If StopNow Then GoTo slut
    TestSolve "5/(x-1)=7/x", "x", "x=3,5"
    If StopNow Then GoTo slut
    TestSolve "120=429" & ChrW$(183) & "L^2,920", "L", "L=0,6464306"
    If StopNow Then GoTo slut
    TestSolve "x^17,3=1/2", "x", "x=0,9607257"
    If StopNow Then GoTo slut
    TestSolve "1,15=" & ChrW$(12310) & "0,9" & ChrW$(12311) & "^a", "a", "a=-1,326512"
    If StopNow Then GoTo slut
    TestSolve "y2/y1=(x2/x1)^a", "a", "a=ln" & ChrW$(8289) & "(y2/y1)/ln" & ChrW$(8289) & "(x2/x1)"
    If StopNow Then GoTo slut
    TestSolve "K_n=K_0" & ChrW$(183) & "(1+r)^n", "r", "r=(K_n/K_0 )^(1/n)-1"
    If StopNow Then GoTo slut
    TestSolve "(sin" & ChrW$(8289) & "(A))/a=(sin(B))/b", "A", "A=sin^(-1)" & ChrW$(8289) & "((sin" & ChrW$(8289) & "(B)" & ChrW$(183) & "a)/b)"
    If StopNow Then GoTo slut
    TestSolve "tan" & ChrW$(8289) & "(A)=x/50", "A", "A=tan^(-1)" & ChrW$(8289) & "(x/50)"
    If StopNow Then GoTo slut
    'TestSolve "L=10" & ChrW$(183) & "log" & ChrW$(8289) & "(I/I_0 )", "I", "I=I_0" & ChrW$(183) & "10^(L/10)"
    If TestSolve("L=10" & ChrW$(183) & "log" & ChrW$(8289) & "(I/I_0 )", "I", "I=10^(L/10)" & ChrW$(183) & "I_0") Then GoTo slut
    TestSolve "4" & ChrW$(183) & "" & ChrW$(8730) & "(x-1)=-1/5" & ChrW$(183) & "x^2+2" & ChrW$(183) & "x+3", "x", "x=5" ' should only provide this one solution
    If StopNow Then GoTo slut
    TestSolve "x^2" & ChrW$(183) & "(x^2-1)=9" & ChrW$(183) & "(x^2-1)", "x", "x=-3    " & ChrW$(8744) & "    x=-1    " & ChrW$(8744) & "    x=1    " & ChrW$(8744) & "    x=3"
    If StopNow Then GoTo slut
    TestSolve "54.4=(0.01767+2" & ChrW$(183) & "x)^2/((0.00600-x)(0.00313-x))", "x", "x=0,001432725    " & ChrW$(8744) & "    x=0,00982426" ' earlier problem with missing multsign in denominator
    If StopNow Then GoTo slut
    TestSolve ChrW$(8747) & "_0^a" & ChrW$(9618) & "" & ChrW$(12310) & "x^2+2dx=13" & ChrW$(12311), "a", "a=2,808381"
    If StopNow Then GoTo slut
    TestSolve ChrW$(8747) & "_0^2" & ChrW$(9618) & "" & ChrW$(12310) & "a" & ChrW$(183) & "x" & ChrW$(183) & "(x-2)^2 " & ChrW$(12311) & " dx=20", "a", "a=15"
    If StopNow Then GoTo slut
    TestSolve "0=(x^3+b)/x^2", "x", "x=-b^(1/3)"
    If StopNow Then GoTo slut
    TestSolve "0,1" & ChrW$(183) & "x^4+0,8" & ChrW$(183) & "x^3+0,6" & ChrW$(183) & "x^2-2" & ChrW$(183) & "x-1=0", "x", "x=-6,689559    " & ChrW$(8744) & "    x=-2,245805    " & ChrW$(8744) & "    x=-0,4727184    " & ChrW$(8744) & "    x=1,408082"
    If StopNow Then GoTo slut
    TestSolve "(x+1,5)^2+ (3/x+2)^2=25", "x", "x=-6,263041    " & ChrW$(8744) & "    x=-0,4357036    " & ChrW$(8744) & "    x=1,5    " & ChrW$(8744) & "    x=2,198745"
    If StopNow Then GoTo slut
    TestSolve "2" & ChrW$(183) & "3^x=4" & ChrW$(183) & "5^2x", "x", "x=ln" & ChrW$(8289) & "(2)/ln" & ChrW$(8289) & "(3/25)"
    If StopNow Then GoTo slut
    TestSolve "2" & ChrW$(183) & "" & ChrW$(12310) & "1,1" & ChrW$(12311) & "^x=3" & ChrW$(183) & "" & ChrW$(12310) & "1,05" & ChrW$(12311) & "^x", "x", "x=ln" & ChrW$(8289) & "(3/2)" & ChrW$(183) & "21,49612@$x=21,49612" & ChrW$(183) & "ln" & ChrW$(8289) & "(3/2)"
    If StopNow Then GoTo slut
    TestSolve "2" & ChrW$(183) & "" & ChrW$(12310) & "(11/10)" & ChrW$(12311) & "^x=3" & ChrW$(183) & "" & ChrW$(12310) & "(105/100)" & ChrW$(12311) & "^x", "x", "x=ln" & ChrW$(8289) & "(3/2)/ln" & ChrW$(8289) & "(22/21)"
    If StopNow Then GoTo slut
    TestSolve "e^(3,044522 x)=0,6666667 e^(3,091042 x)", "x", "x=8,71593"
    If StopNow Then GoTo slut
    TestSolve "92,3*(e^(0,21278*x)-1)/(e^(0,21278*x)+1)=0.5", "x", "x=0,05091805"
    If StopNow Then GoTo slut
    TestSolve "x^2+x-12" & ChrW$(8805) & "8", "x", "x" & ChrW$(8804) & "-5   " & ChrW$(8744) & "   x" & ChrW$(8805) & "4" ' ulighed
    If StopNow Then GoTo slut
    If TestSolve("x^3-4" & ChrW$(183) & "x+2=0", "x", "x=-2,21432    " & ChrW$(8744) & "    x=0,5391889    " & ChrW$(8744) & "    x=1,675131") Then GoTo slut 'Has 3 real solution. That look imaginary in Maxima
    If TestSolve("e^(-x)=2", "x", "x=-ln" & ChrW$(8289) & "(2)") Then GoTo slut
    If TestSolve("1=1,5" & ChrW$(183) & "x^2+0,5/x^0,5 -1", "x", "x=0,06287224    " & ChrW$(8744) & "    x=1") Then GoTo slut ' requires numerical solution. Previous to_poly error, with only one solution.
    If TestSolve("6174/(1+8652" & ChrW$(183) & "" & ChrW$(12310) & "0,711" & ChrW$(12311) & "^x )=8000000", "x", "x" & ChrW$(8712) & "" & ChrW$(8709)) Then GoTo slut
    If TestSolve("x^2=e^(-x^2 )", "x", "x=-" & ChrW$(8730) & "(W_0 (1) )    " & ChrW$(8744) & "    x=" & ChrW$(8730) & "(W_0 (1) )") Then GoTo slut
    If TestSolve("0=3" & ChrW$(183) & "" & ChrW$(8730) & "((x-2) )-x+1", "x", "x=-(3" & ChrW$(183) & "" & ChrW$(8730) & "5-11)/2    " & ChrW$(8744) & "    x=(3" & ChrW$(183) & "" & ChrW$(8730) & "5+11)/2@$x=2,145898    " & ChrW$(8744) & "    x=8,854102") Then GoTo slut ' this can be solved using dissolveradicals, but if disabled also has the numerical values for nsolve
    If TestSolve("0=3" & ChrW$(183) & "" & ChrW$(8730) & "(x-2)-x+1", "x", "x=-(3" & ChrW$(183) & "" & ChrW$(8730) & "5-11)/2    " & ChrW$(8744) & "    x=(3" & ChrW$(183) & "" & ChrW$(8730) & "5+11)/2@$x=2,145898    " & ChrW$(8744) & "    x=8,854102") Then GoTo slut ' this is almost a copy of the one above but entered differently
    If TestSolve(ChrW$(8730) & "(x+" & ChrW$(8730) & "(x-1)) =x^2", "x", "x=1    " & ChrW$(8744) & "    x=1,081031") Then GoTo slut
    If TestSolve("(3x+9)^(1/3)=x+3", "x", "x=-4,732051    " & ChrW$(8744) & "    x=-3    " & ChrW$(8744) & "    x=-1,267949") Then GoTo slut
    If TestSolve("0=3" & ChrW$(183) & "(x-2)^(-0,5)-x+1", "x", "x=3,472368") Then GoTo slut
    If TestSolve("1,2" & ChrW$(183) & "(" & ChrW$(8730) & "5" & ChrW$(183) & "x-x^1,5 )^0,5=0", "x", "x=0    " & ChrW$(8744) & "    x=5") Then GoTo slut
    If TestSolve("-x^2/9+2x/9+1/(x-2)-1/9=0", "x", "x=3,472368") Then GoTo slut
    If TestSolve(ChrW$(8730) & "(3x+9)=x+3", "x", "x=-3    " & ChrW$(8744) & "    x=0") Then GoTo slut
    If TestSolve("(2" & ChrW$(183) & "" & ChrW$(960) & "" & ChrW$(183) & "x" & ChrW$(183) & "" & ChrW$(8730) & "(1-x^2 ))/3-(" & ChrW$(960) & "" & ChrW$(183) & "x^3)/(3" & ChrW$(183) & "" & ChrW$(8730) & "(1-x^2 ))=0", "x", "x=0    " & ChrW$(8744) & "    x=-" & ChrW$(8730) & "2/" & ChrW$(8730) & "3    " & ChrW$(8744) & "    x=" & ChrW$(8730) & "2/" & ChrW$(8730) & "3") Then GoTo slut
    If TestSolve("-(80x^2-1600000)/" & ChrW$(8730) & "(40000-x^2 )=0", "x", "x=-25" & ChrW$(183) & "2^(5/2)    " & ChrW$(8744) & "    x=25" & ChrW$(183) & "2^(5/2)") Then GoTo slut
    If TestSolve("(3" & ChrW$(183) & "" & ChrW$(8730) & "x)/(2" & ChrW$(183) & "(x-2) )-x^(3/2)/(x-2)^2 =0", "x", "x=0    " & ChrW$(8744) & "    x=6") Then GoTo slut
    If TestSolve("x" & ChrW$(183) & "e^2x+e^2x=0", "x", "x" & ChrW$(8776) & "-1@$x=-1", "Click maxima numeric") Then GoTo slut  ' Only x=-1 is a solution
    If TestSolve("ln" & ChrW$(8289) & "(x)=x-1", "x", "x=1") Then GoTo slut
    If TestSolve("20,7=100" & ChrW$(8729) & "(1/2)^(100/T)", "T", "T=44,00833@$T=(100" & ChrW$(183) & "ln" & ChrW$(8289) & "(2))/(2" & ChrW$(183) & "ln" & ChrW$(8289) & "(2)+0,1887421)") Then GoTo slut
    If TestSolve("(b-a)" & ChrW$(183) & "x^2=" & ChrW$(8730) & "x", "x", "x=0    " & ChrW$(8744) & "    x=1/(a^2-2" & ChrW$(183) & "b" & ChrW$(183) & "a+b^2 )^(1/3)") Then GoTo slut
    If TestSolve("211,4885-10,4801" & ChrW$(183) & "(e^0,0329x+e^(-0,0329x) )=0", "x", "x=-91,25312    " & ChrW$(8744) & "    x=91,25312") Then GoTo slut
    If TestSolve(ChrW$(8730) & "x=1/8" & ChrW$(183) & "x^2", "x", "x=0    " & ChrW$(8744) & "    x=4") Then GoTo slut
    If TestSolve("0,004x+25/" & ChrW$(8730) & "x=10", "x", "x=6,281526    " & ChrW$(8744) & "    x=2371,663@$x=6,281527    " & ChrW$(8744) & "    x=2371,663") Then GoTo slut
    If TestSolve("x+(x^3-100)^(1/3)=0", "x", "x=3,684031") Then GoTo slut ' has caused problems because of 1/3 to 0.3333333
    If TestSolve("(" & ChrW$(9608) & "(x+y@x-y))=(" & ChrW$(9608) & "(1@2))", "x", "y=-1/2    " & ChrW$(8743) & "    x=3/2") Then GoTo slut ' vector equation that is a system of two simple equations
    If TestSolve("(x+1)^2+(y)^2=5^2  " & ChrW$(8743) & " (x-1)^2+(y)^2=5^2", "x", "(y=-2" & ChrW$(183) & "" & ChrW$(8730) & "6    " & ChrW$(8743) & "    x=0)   " & ChrW$(8744) & "   (y=2" & ChrW$(183) & "" & ChrW$(8730) & "6    " & ChrW$(8743) & "    x=0)") Then GoTo slut ' two circles exact
    If TestSolve("8" & ChrW$(183) & "" & ChrW$(960) & "" & ChrW$(183) & "r-(16" & ChrW$(183) & "" & ChrW$(960) & "" & ChrW$(183) & "r^2)/(100-(4" & ChrW$(183) & "" & ChrW$(960) & "" & ChrW$(183) & "r^3)/3)^(1/3) =0", "r", "r=0    " & ChrW$(8744) & "    r=2,016879") Then GoTo slut ' has caused problems before check of solutions
    If TestSolve("x^2-5=ln" & ChrW$(8289) & "(x)", "x", "x=0,006738253    " & ChrW$(8744) & "    x=2,426173") Then GoTo slut
    If TestSolve("1214729,56=21520,64" & ChrW$(183) & "(1-(1+r)^(-80))/r", "r", "r=-1,942866    " & ChrW$(8744) & "    r=0,009203618") Then GoTo slut ' failed because of problem in removeundefined

    'solvesystem
    Selection.TypeText "solving systems of equations": Selection.TypeParagraph
    If TestSolve("(x-38)^2+(y-18)^2=" & ChrW$(12310) & "17,5" & ChrW$(12311) & "^2  " & ChrW$(8743) & "  (x-27)^2+(y-33)^2=" & ChrW$(12310) & "9,4" & ChrW$(12311) & "^2", "x", "(y=25,05068    " & ChrW$(8743) & "    x=21,9832)   " & ChrW$(8744) & "   (y=35,39542    " & ChrW$(8743) & "    x=36,08966)") Then GoTo slut
    If TestSolve("x^2+y^2-12x+2y=-33 " & ChrW$(8743) & "  x^2+y^2+2x-6y=26", "x", "[x]" & ChrW$(8712) & "" & ChrW$(8709)) Then GoTo slut
    
    MaximaComplex = True ' complex
    If TestSolve("x^2=-4", "x", "x=-2" & ChrW$(183) & "i    " & ChrW$(8744) & "    x=2" & ChrW$(183) & "i") Then GoTo slut
    PolarOutput = True
    If TestBeregn("2" & ChrW$(183) & "i", "=2" & ChrW$(8736) & "90") Then GoTo slut
    PolarOutput = False
    MaximaComplex = False
    
    MaximaExact = 2 ' num
    If TestSolve("40=72" & ChrW$(183) & "e^((0,619/0,22" & ChrW$(8729) & "(e^22-e^0,22t )) )", "t", "t=100") Then GoTo slut
    If TestSolve(ChrW$(8730) & "(x^2+40000/x^4 )+x" & ChrW$(183) & "(2" & ChrW$(183) & "x-160000/x^5 )/(2" & ChrW$(183) & "" & ChrW$(8730) & "(x^2+40000/x^4 ))+2" & ChrW$(183) & "x-400/x^2 =0", "x", "x=5,520396") Then GoTo slut
    InsertTestMath "Definer: " & ChrW$(963) & ">0"
    If TestSolve("0,1=" & ChrW$(8747) & "_(-" & ChrW$(8734) & ")^5" & ChrW$(9618) & "1/(" & ChrW$(8730) & "2" & ChrW$(960) & "" & ChrW$(183) & "" & ChrW$(963) & ")" & ChrW$(183) & "e^(-1/2" & ChrW$(183) & "((y-7)/" & ChrW$(963) & ")^2 ) dy", "sigma", ChrW$(963) & "=1,560608") Then GoTo slut
    InsertSletDef
    InsertTestMath "Definer: f(x)={" & ChrW$(9608) & "(100" & ChrW$(183) & "" & ChrW$(12310) & "1,03" & ChrW$(12311) & "^x,  0" & ChrW$(8804) & "x" & ChrW$(8804) & "72@33743" & ChrW$(183) & "" & ChrW$(12310) & "0,95" & ChrW$(12311) & "^x,  72<x" & ChrW$(8804) & "150)" & ChrW$(9508)
    If TestSolve("f(x)=200", "x", "x=23,44977    " & ChrW$(8744) & "    x=99,97819") Then GoTo slut ' piecewise cannot be handled algebraically, so solvereal must go directly to nsolve
    InsertSletDef
    
    If TestSolve("0=3" & ChrW$(183) & "" & ChrW$(8730) & "((x-2) )-x+1", "x", "x=2,145898    " & ChrW$(8744) & "    x=8,854102") Then GoTo slut
    If TestSolve("112,02267" & ChrW$(183) & "e^(-(0,0057195" & ChrW$(183) & "x) )/(1,80517" & ChrW$(183) & "e^(-(0,0057195" & ChrW$(183) & "x) )+1)^2 -1,97=0", "x", "x=-488,2615    " & ChrW$(8744) & "    x=694,8022") Then GoTo slut ' This one can only be solved when set to num. If set to exact and STOPed and retried it will be solved. May be in the future it will be solved analytically using a substitution method.
    Selection.TypeText "OBS: Denne kan tage lang tid. Vent med at trykke STOP": Selection.TypeParagraph
    If TestSolve("8" & ChrW$(183) & "" & ChrW$(960) & "" & ChrW$(183) & "r-(16" & ChrW$(183) & "" & ChrW$(960) & "" & ChrW$(183) & "r^2)/(100-(4" & ChrW$(183) & "" & ChrW$(960) & "" & ChrW$(183) & "r^3)/3)^(1/3) =0", "r", "r=0    " & ChrW$(8744) & "    r=2,016879") Then GoTo slut ' this and the next are almost identical except for num radical. nsolve will solve it correctly, but dissolve radical has problems if applied. Previous problem: it will not return r=0 because of assume r#0 and the num version can cause float overflow
    If TestSolve("25,13274" & ChrW$(183) & "r-(50,26548" & ChrW$(183) & "r^2)/(100-4,18879" & ChrW$(183) & "r^3 )^0,3333333 =0", "r", "r=0    " & ChrW$(8744) & "    r=2,016879") Then GoTo slut
    If TestSolve("67" & ChrW$(183) & "(e^(-0,041" & ChrW$(183) & "t)-e^(-3,1" & ChrW$(183) & "t) )=25", "t", "t=0,1539059    " & ChrW$(8744) & "    t=24,04431") Then GoTo slut ' would hang due to float caused by two tellsimpafter, only works for numeric right now
    
    If Not NonInterA Then
        If TestSolve(ChrW$(8747) & "_(-" & ChrW$(8734) & ")^5" & ChrW$(9618) & "" & ChrW$(12310) & "1/(" & ChrW$(8730) & "2" & ChrW$(960) & "" & ChrW$(183) & "" & ChrW$(963) & ")" & ChrW$(183) & "e^(-1/2 ((x-" & ChrW$(956) & ")/" & ChrW$(963) & ")^2 ) dx=0,3" & ChrW$(12311) & "   " & ChrW$(8743) & "  " & ChrW$(8747) & "_(-" & ChrW$(8734) & ")^10" & ChrW$(9618) & "" & ChrW$(12310) & "1/(" & ChrW$(8730) & "2" & ChrW$(960) & "" & ChrW$(183) & "" & ChrW$(963) & ")" & ChrW$(183) & "e^(-1/2 ((x-" & ChrW$(956) & ")/" & ChrW$(963) & ")^2 ) dx=0,9" & ChrW$(12311), "mu", ChrW$(956) & "=6,451867    " & ChrW$(8743) & "    " & ChrW$(963) & "=2,768623") Then GoTo slut ' it is slow. in nsolve the taylorapprox halted. Now taylororder is lowered for long expression, and time is tested. will ask for sign.
        InsertTestMath "Press STOP and retry on this equation. "
        If TestBeregn("2" & ChrW$(960) & "" & ChrW$(8747) & "_(-r)^r" & ChrW$(9618) & "" & ChrW$(12310) & "" & ChrW$(8730) & "(r^2-x^2 )" & ChrW$(183) & "" & ChrW$(8730) & "(1+(-x/" & ChrW$(8730) & "(r^2-x^2 ))^2 ) dx" & ChrW$(12311), "=12,56637" & ChrW$(183) & "r^(2)") Then GoTo slut ' only asks for sign when set to num
    End If
    
    MaximaExact = 1 ' Exact
    
    '    TestSolve "", "x", "x="
    '    If StopNow Then GoTo slut
    
    'Often fails, but not always ??? then works numerically. It's something with an overflow error. Doesn't seem to be working numerically anymore
    ' result is numerical 100. Sometimes this runs in luck and needs to be repeated numerically. Unknown why.
    If Not NonInterA Then
        InsertTestMath "OBS: Denne kan tage lang tid. Vent med at trykke STOP"
        Selection.TypeParagraph
        DoEvents
        If TestSolve("40=72" & ChrW$(183) & "e^((0,619/0,22" & ChrW$(8729) & "(e^22-e^0,22t )) )", "t", "t=100@$t=ln" & ChrW$(8289) & "(ln" & ChrW$(8289) & "(5/9)" & ChrW$(183) & "-0,355412+e^22 )" & ChrW$(183) & "4,545455@$t=4,545455" & ChrW$(183) & "ln" & ChrW$(8289) & "(e^22-0,355412" & ChrW$(183) & "ln" & ChrW$(8289) & "(5/9) )") Then GoTo slut
    End If
        
    ' Defintion test
    InsertTestMath "T_2" & ChrW$(8788) & "8"
    Selection.TypeParagraph
    TestSolve "T_2=log" & ChrW$(8289) & "(2)/log" & ChrW$(8289) & "(a) ", "a", "a=2^(1/8)"
    InsertSletDef
    If StopNow Then GoTo slut

    InsertTestMath "f(x)" & ChrW$(8788) & "4x-2,5", True
    InsertTestMath "g(x)" & ChrW$(8788) & "2" & ChrW$(183) & "" & ChrW$(12310) & "0,8" & ChrW$(12311) & "^x", True
'    TestBeregn "f(x)+g(x)", "=2" & ChrW$(183) & "0,8^(x)+4" & ChrW$(183) & "x-2,5"
    TestBeregn "f(x)+g(x)", "4" & ChrW$(183) & "x+2" & ChrW$(183) & "" & ChrW$(12310) & "0,8" & ChrW$(12311) & "^x-2,5"
    InsertSletDef
    If StopNow Then GoTo slut
    
    InsertTestMath "f(x)" & ChrW$(8788) & "" & ChrW$(8730) & "(r^2-x^2 )"
    Selection.TypeParagraph
    TestBeregn "S=2" & ChrW$(960) & "" & ChrW$(8747) & "_(-r)^r" & ChrW$(9618) & "" & ChrW$(12310) & "f(x)" & ChrW$(183) & "" & ChrW$(8730) & "(1+(f^' (x))^2 ) dx" & ChrW$(12311), "=4" & ChrW$(183) & "" & ChrW$(960) & "" & ChrW$(183) & "r^(2)"
    InsertSletDef
    If StopNow Then GoTo slut

    InsertTestMath "Definer: a=2 ;g(x)=x^2"
    Selection.TypeParagraph
    InsertTestMath "h(x)" & ChrW$(8788) & "2x"
    Selection.TypeParagraph
    InsertTestMath "b" & ChrW$(8788) & "3 ;c" & ChrW$(8788) & "2"
    Selection.TypeParagraph
    TestBeregn "g(h(a+b))", "=100"
    InsertSletDef
    If StopNow Then GoTo slut

    InsertTestMath "Definer: x>0"
    Selection.TypeParagraph
    TestSolve "x^2=9", "x", "x=3"
    InsertSletDef
    If StopNow Then GoTo slut

    InsertTestMath "Definer: f(x;y)=x^2" & ChrW$(183) & "y+2y"
    If TestBeregn(ChrW$(8711) & "f(1;2)", "=(" & ChrW$(9632) & "(4@3))") Then GoTo slut
    InsertSletDef

    MaximaExact = 1 ' exact
    AllTrig = True
    Radians = True
    InsertTestMath "Definer: -7,5<x<7,5"
  ' If TestSolve("1,386" & ChrW$(183) & "cos" & ChrW$(8289) & "(0,63" & ChrW$(183) & "x)=0", "x", "x=" & ChrW$(960) & "" & ChrW$(183) & "-2,380952    " & ChrW$(8744) & "    x=" & ChrW$(960) & "" & ChrW$(183) & "-0,7936508    " & ChrW$(8744) & "    x=" & ChrW$(960) & "" & ChrW$(183) & "0,7936508    " & ChrW$(8744) & "    x=" & ChrW$(960) & "" & ChrW$(183) & "2,380952@$x=-(2,380952" & ChrW$(183) & "" & ChrW$(960) & ")    " & ChrW$(8744) & "    x=-(0,7936508" & ChrW$(183) & "" & ChrW$(960) & ")    " & ChrW$(8744) & "    x=0,7936508" & ChrW$(183) & "" & ChrW$(960) & "    " & ChrW$(8744) & "    x=2,380952" & ChrW$(183) & "" & ChrW$(960)) Then GoTo slut
    If TestSolve("1,386" & ChrW$(183) & "cos" & ChrW$(8289) & "(0,63" & ChrW$(183) & "x)=0", "x", "x=-2,380952" & ChrW$(183) & "" & ChrW$(960) & "    " & ChrW$(8744) & "    x=-0,7936508" & ChrW$(183) & "" & ChrW$(960) & "    " & ChrW$(8744) & "    x=0,7936508" & ChrW$(183) & "" & ChrW$(960) & "    " & ChrW$(8744) & "    x=2,380952" & ChrW$(183) & "" & ChrW$(960)) Then GoTo slut
    InsertSletDef
    
    
' this one fails, but not if it is swapped with the one above, then that is the one that fails ????
    Radians = True
    Selection.TypeText "The next one fails, but not if it is swapped with the one above, then that is the one that fails...."
    Selection.TypeParagraph
    InsertTestMath "Definer: -2<x<2"
    Selection.TypeParagraph
    If TestSolve("sin" & ChrW$(8289) & "" & ChrW$(12310) & "(x)" & ChrW$(12311) & " " & ChrW$(8730) & "(cos" & ChrW$(8289) & "" & ChrW$(12310) & "(x)" & ChrW$(12311) & " )=0", "x", "x=0    " & ChrW$(8744) & "    x=-" & ChrW$(960) & "/2    " & ChrW$(8744) & "    x=" & ChrW$(960) & "/2") Then GoTo slut

    InsertSletDef
    If StopNow Then GoTo slut
    
    Radians = True
    If TestBeregn("sin" & ChrW$(8289) & "" & ChrW$(12310) & "a/b" & ChrW$(12311) & "+2", "=sin" & ChrW$(8289) & "(a/b)+2") Then GoTo slut
    
    MaximaExact = 0
    AllTrig = False
    Radians = False

    '    TestSolve "", "x", "x="
    '    If StopNow Then GoTo slut
    
    If TestSolve("x^2=" & ChrW$(12310) & "0,7" & ChrW$(12311) & "^x", "x", "x " & ChrW$(8776) & " -15,29371    " & ChrW$(8744) & "    x " & ChrW$(8776) & " -1,249643    " & ChrW$(8744) & "    x " & ChrW$(8776) & " 0,8581024@$x=-15,29371    " & ChrW$(8744) & "    x=-1,249643    " & ChrW$(8744) & "    x=0,8581024", "Click maxima numeric") Then GoTo slut
    
    If Not NonInterA Then '****** Interactive start *******
        ' requires user to stop and choose retry numeric
        TestBeregn "(1+1/10^12 )^(10^12 )", ChrW$(8776) & "2,718523"
        If StopNow Then GoTo slut

        ' Requires user input. Equations that trigger numeric solution
    End If '****** Interactive end *******
    '    TestBeregn chrw$(8747) & "_(-" & chrw$(8734) & ")^" & chrw$(8734) & "" & chrw$(9618) & "1/(" & chrw$(8730) & "2" & chrw$(960) & "á3)áe^(-1/2á((y-1)/3)^2 ) dy", "=1"
    
    
    ' Numeric test
    MaximaExact = 2 ' numeric
    ShowSettings "Numerisk test"

    '    TestBeregn "cos^(-1)" & chrw$(8289) & "(2)", chrw$(8776) & "Ikke defineret indenfor " & chrw$(8477) ' med VBACAS kommer der fejlpopup
    '    If StopNow Then GoTo slut
    If TestBeregn(ChrW$(12310) & "cos" & ChrW$(12311) & "^(-1) (0,5)", ChrW$(8776) & "60") Then GoTo slut  ' testing inverse trig and entering via design-script-superscript
    If TestBeregn("23456789123", ChrW$(8776) & "2,345679" & ChrW$(183) & "10^10@$23456790000") Then GoTo slut
    TestBeregn "B=sin^(-1)" & ChrW$(8289) & "(sin" & ChrW$(8289) & "(40)" & ChrW$(183) & "0,8605341)", ChrW$(8776) & "33,58274"
    If TestBeregn("(1+1/10^12 )^(10^12 )", ChrW$(8776) & "2,718523") Then GoTo slut
    If TestBeregn(ChrW$(8747) & "_(-91,25312)^91,25312" & ChrW$(9618) & "" & ChrW$(8730) & "(1+(0,3447953" & ChrW$(183) & "e^(-(0,0329" & ChrW$(183) & "x) )-0,3447953" & ChrW$(183) & "e^(0,0329" & ChrW$(183) & "x) )^2 ) dx", ChrW$(8776) & "451,2555") Then GoTo slut
    
    MaximaExact = 0
    ShowSettings
        '****************************SolveDE ok here
    'numeric definition test
    InsertTestMath "definer: f(x)=-x^2" & ChrW$(8729) & "" & ChrW$(8730) & "x+2/x;x_1=0,25"
    Selection.TypeParagraph
    
    TestSolve "f^' (x)=-32", "x", "x=0,251239    " & ChrW$(8744) & "    x=5,464284@$x" & ChrW$(8776) & "0,251239    " & ChrW$(8744) & "    x" & ChrW$(8776) & "5,464284@$x=0,2512389    " & ChrW$(8744) & "    x=5,464284"
    If StopNow Then GoTo slut
        
    TestBeregn "y=f^' (x_1 )" & ChrW$(183) & "(x-x_1 )+f(x_1 )", ChrW$(8776) & "-32,3125" & ChrW$(183) & "x+16,04688@$" & ChrW$(8776) & "16,04688-32,3125" & ChrW$(183) & "x"
    InsertSletDef
    If StopNow Then GoTo slut

    If Not NonInterA Then '****** Interactive start *******
        InsertTestMath "Definer: f(x)=211,4885-10,4801" & ChrW$(183) & "(e^0,0329x+e^(-0,0329x) )"
        Selection.TypeParagraph
        TestBeregn "s=" & ChrW$(8747) & "_(-91,25312)^91,25312" & ChrW$(9618) & "" & ChrW$(8730) & "(1+(f^' (x))^2 ) dx", ChrW$(8776) & "451,2554"
        InsertSletDef
        If StopNow Then GoTo slut
    End If '****** Interactive end *******
    
    
    ' Scientific notation test
    MaximaExact = 2
    MaximaDecOutType = 3
    ShowSettings "Videnskabelig notation test"
    TestBeregn "123", "=1,23" & ChrW$(183) & "10^2"
    If StopNow Then GoTo slut
    MaximaDecOutType = 2

    ' Calculus
    MaximaExact = 0
    ShowSettings "Differential- og integralregning"
    If TestBeregn(ChrW$(8518) & "/" & ChrW$(8518) & "x (" & ChrW$(8747) & "x^2 dx)", "=x^(2)") Then GoTo slut
    If TestBeregn(ChrW$(8706) & "^2/" & ChrW$(8706) & "x" & ChrW$(8706) & "y (x^2" & ChrW$(183) & "y)", "=2" & ChrW$(183) & "x") Then GoTo slut
    If TestBeregn(ChrW$(8518) & "^2/(" & ChrW$(8518) & "x^2 ) (x^2+2x)", "=2") Then GoTo slut
    If TestBeregn(ChrW$(8747) & "_1^10" & ChrW$(9618) & "" & ChrW$(8730) & "(x^(-2)+1) dx", "=-(ln" & ChrW$(8289) & "(" & ChrW$(8730) & "(101)+1)-ln" & ChrW$(8289) & "(" & ChrW$(8730) & "(101)-1)-ln" & ChrW$(8289) & "(" & ChrW$(8730) & "(2)+1)+ln" & ChrW$(8289) & "(" & ChrW$(8730) & "(2)-1)-2" & ChrW$(183) & "" & ChrW$(8730) & "(101)+2^(3/2))/2" & ChrW$(8776) & "9,417202@$=-((ln" & ChrW$(8289) & "(" & ChrW$(8730) & "(101)+1)-ln" & ChrW$(8289) & "(" & ChrW$(8730) & "(101)-1)-ln" & ChrW$(8289) & "(" & ChrW$(8730) & "(2)+1)+ln" & ChrW$(8289) & "(" & ChrW$(8730) & "(2)-1)-2" & ChrW$(183) & "" & ChrW$(8730) & "(101)+2^(3/2))/2)" & ChrW$(8776) & "9,417202") Then GoTo slut
    InsertTestMath "Definer: f(x)=" & ChrW$(8730) & "(3x+9)  ;g(x)=x+3"
    If TestSolve(ChrW$(8747) & "_0^k" & ChrW$(9618) & "" & ChrW$(12310) & "g(x)-f(x) " & ChrW$(12311) & " dx=1,5", "k", "k=-3    " & ChrW$(8744) & "    k=7/3@$k=-3    " & ChrW$(8744) & "    k=2,333333") Then GoTo slut
    ' the next two previously gave incorrect results due to using ldefint in integrate, so now Nintegrate is used before ldefint, which is obviously fallible.
    If TestBeregn(ChrW$(8747) & "_(-1)^1" & ChrW$(9618) & "" & ChrW$(8730) & "(1+((4" & ChrW$(183) & "x-4" & ChrW$(183) & "x^3 )" & ChrW$(183) & "e^(2" & ChrW$(183) & "x^2-x^4 ) )^2 ) dx", ChrW$(8776) & "4,142057") Then GoTo slut
    InsertTestMath "Definer: f(x)=0,000003" & ChrW$(183) & "x^4-0,01676" & ChrW$(183) & "x^2+60"
    Selection.TypeParagraph
    If TestBeregn(ChrW$(8747) & "_(-52)^52" & ChrW$(9618) & "" & ChrW$(12310) & "" & ChrW$(8730) & "(1+(f^' (x))^2 ) dx" & ChrW$(12311), ChrW$(8776) & "115,7009") Then GoTo slut
    InsertTestMath "Definer: f(x,y)=x^2+3x" & ChrW$(183) & "y-2y+1"
    If TestBeregn("(" & ChrW$(8706) & "^2 f(x,y))/" & ChrW$(8706) & "x" & ChrW$(8706) & "y", "=3") Then GoTo slut
    InsertSletDef


    ' Vectortest
    MaximaExact = 0
    ShowSettings "Vektorer"
    If TestBeregn("((" & ChrW$(9608) & "(-4@3))" & ChrW$(183) & "(" & ChrW$(9608) & "(1@7)))/|((" & ChrW$(9608) & "(-4@3)))|^2 " & ChrW$(183) & "(" & ChrW$(9608) & "(-4@3))", "=(" & ChrW$(9632) & "(-68/25@51/25))=(" & ChrW$(9632) & "(-2,72@2,04))@$=(" & ChrW$(9632) & "(-(68/25)@51/25))=(" & ChrW$(9632) & "(-2,72@2,04))") Then GoTo slut
    If TestBeregn("(" & ChrW$(9632) & "(-400@0@320))" & ChrW$(215) & "(" & ChrW$(9632) & "(-120@280@0))", "=(" & ChrW$(9632) & "(-89600@-38400@-112000))") Then GoTo slut 'Two different ways to enter vectors:
    If TestBeregn("(" & ChrW$(9608) & "(-400@0@320))" & ChrW$(215) & "(" & ChrW$(9608) & "(-120@280@0))", "=(" & ChrW$(9632) & "(-89600@-38400@-112000))") Then GoTo slut
    If TestBeregn("det" & ChrW$(8289) & "((" & ChrW$(9608) & "(1@2)),(" & ChrW$(9608) & "(3@-4)))", "=-10") Then GoTo slut
    If TestSolve("(" & ChrW$(9608) & "(y+x@x-y))=(" & ChrW$(9608) & "(1@2))", "y", "y=-1/2    " & ChrW$(8743) & "    x=3/2") Then GoTo slut ' this one only works if equations written in correct order and correct variable chosen to solve for.

    MaximaExact = 2
    TestBeregn "rref([" & ChrW$(9632) & "(-1&0&1/2&1/2@1/2&-1&0&1/2@1/2&1&-1&0@0&0&1/2&-1)])", "=[" & ChrW$(9632) & "(1&0&0&-1,5@0&1&0&-1,25@0&0&1&-2@0&0&0&0)]"
    If StopNow Then GoTo slut
    MaximaExact = 0

    InsertTestMath "Definer: s" & ChrW$(8407) & "(t)=(" & ChrW$(9608) & "(2+t@t^2 ))"
    If StopNow Then GoTo slut
    TestBeregn "s" & ChrW$(8407) & "(2)", "=(" & ChrW$(9632) & "(4@4))"
    If StopNow Then GoTo slut
    TestBeregn "s" & ChrW$(8407) & "^'(t)", "=(" & ChrW$(9632) & "(1@2" & ChrW$(183) & "t))"
    If StopNow Then GoTo slut
    
    ' Differential equation test
    If TestSolveDE("N^'=1/10500" & ChrW$(183) & "N" & ChrW$(183) & "(1000-N)", "N,x", "N=0    " & ChrW$(8744) & "    N=1000    " & ChrW$(8744) & "    N=1000/(e^(-((2" & ChrW$(183) & "x)/21) )" & ChrW$(183) & "c+1)") Then GoTo slut
    If TestSolveDE("N^'=(0,025-0,0004t)" & ChrW$(183) & "N", "N,t", "N=e^(t/40-t^2/5000)" & ChrW$(183) & "c") Then GoTo slut
    If TestSolveDE("y^'+2x" & ChrW$(183) & "y=x", "y,x", "y=e^(-x^2 )" & ChrW$(183) & "c+1/2") Then GoTo slut
'    If TestSolveDE("y^'+2x" & ChrW$(183) & "y=x", "y,x", "y=c" & ChrW$(183) & "e^(-x^2 )+1/2") Then GoTo slut
    If TestSolveDE("L^'=k" & ChrW$(183) & "(100-L)", "L,x", "L=e^(-(k" & ChrW$(183) & "x) )" & ChrW$(183) & "c+100") Then GoTo slut
    If TestSolveDE("2y^'+y^2-5y=0", "y,x", "y=0    " & ChrW$(8744) & "    y=5    " & ChrW$(8744) & "    y=5/(e^(-((5" & ChrW$(183) & "x)/2) )" & ChrW$(183) & "c+1)") Then GoTo slut
    If TestSolveDE("y^'=5y" & ChrW$(183) & "(y+1)", "y,x", "y=0    " & ChrW$(8744) & "    y=-1    " & ChrW$(8744) & "    y=-1/(e^(-(5" & ChrW$(183) & "x) )" & ChrW$(183) & "c+1)") Then GoTo slut
    If TestSolveDE("y^'=b" & ChrW$(183) & "y" & ChrW$(183) & "(b/a-y)", "y,x", "y=0    " & ChrW$(8744) & "    y=b/a    " & ChrW$(8744) & "    y=b/(e^(-((b^2" & ChrW$(183) & "x)/a) )" & ChrW$(183) & "c" & ChrW$(183) & "a+a)") Then GoTo slut
    If TestSolveDE("N^'=0,00526" & ChrW$(183) & "N" & ChrW$(183) & "(209-N)", "N=30;x=103", "N=209/(e^(-(1,09934" & ChrW$(183) & "x) )" & ChrW$(183) & "8,948974" & ChrW$(183) & "10^49+1)") Then GoTo slut
    If TestSolveDE("y^'=b" & ChrW$(183) & "y" & ChrW$(183) & "(M-y)", "y,x", "y=0    " & ChrW$(8744) & "    y=M    " & ChrW$(8744) & "    y=M/(e^(-(M" & ChrW$(183) & "b" & ChrW$(183) & "x) )" & ChrW$(183) & "c+1)") Then GoTo slut

'    Other ordering for these c*e^x not e^x*c
'    TestSolveDE "N^'=1/10500" & ChrW$(183) & "N" & ChrW$(183) & "(1000-N)", "N,x", "N=0    " & ChrW$(8744) & "    N=1000    " & ChrW$(8744) & "    N=1000/(c" & ChrW$(183) & "e^(-((2" & ChrW$(183) & "x)/21) )+1)"
'    TestSolveDE "N^'=(0,025-0,0004t)" & ChrW$(183) & "N", "N,t", "N=c" & ChrW$(183) & "e^(t/40-t^2/5000)"
'    TestSolveDE "y^'=-2x/(1+x^2 )" & ChrW$(183) & "y+1/(1+x^2 )", "y,x", "y=(x+c)/(x^2+1)"
'    TestSolveDE "L^'=k" & ChrW$(183) & "(100-L)", "L,x", "L=c" & ChrW$(183) & "e^(-(k" & ChrW$(183) & "x) )+100"
'    TestSolveDE "2y^'+y^2-5y=0", "y,x", "y=0    " & ChrW$(8744) & "    y=5    " & ChrW$(8744) & "    y=5/(c" & ChrW$(183) & "e^(-((5" & ChrW$(183) & "x)/2) )+1)"
'    TestSolveDE "y^'=5y" & ChrW$(183) & "(y+1)", "y,x", "y=0    " & ChrW$(8744) & "    y=-1    " & ChrW$(8744) & "    y=-1/(c" & ChrW$(183) & "e^(-(5" & ChrW$(183) & "x) )+1)@$y=0    " & ChrW$(8744) & "    y=-1    " & ChrW$(8744) & "    y=-(1/(c" & ChrW$(183) & "e^(-(5" & ChrW$(183) & "x) )+1))"
'    TestSolveDE "y^'=b" & ChrW$(183) & "y" & ChrW$(183) & "(b/a-y)", "y,x", "y=0    " & ChrW$(8744) & "    y=b/a    " & ChrW$(8744) & "    y=b/(c" & ChrW$(183) & "a" & ChrW$(183) & "e^(-((b^2" & ChrW$(183) & "x)/a) )+a)"
'    If TestSolveDE("N^'=0,00526" & ChrW$(183) & "N" & ChrW$(183) & "(209-N)", "N=30;x=103", "N=209/(8,948974" & ChrW$(183) & "10^49" & ChrW$(183) & "e^(-(1,09934" & ChrW$(183) & "x) )+1)") Then GoTo slut
'    If TestSolveDE("y^'=b" & ChrW$(183) & "y" & ChrW$(183) & "(M-y)", "y,x", "y=0    " & ChrW$(8744) & "    y=M    " & ChrW$(8744) & "    y=M/(c" & ChrW$(183) & "e^(-(M" & ChrW$(183) & "b" & ChrW$(183) & "x) )+1)") Then GoTo slut
    
    If TestSolveDE("(y^' )^2+x" & ChrW$(183) & "y^'=0", "y,x", "y=c    " & ChrW$(8744) & "    y=c-x^2/2") Then GoTo slut
    ' particular solutions
    ' This has not previously given a solution. It was sorted out in ic1real when TESTTF found a small difference in the constants, and then thought it was a false solution. TestTF has now got numerical comparison
    If TestSolveDE("p^'=0,015" & ChrW$(183) & "p^1,2", "p=5,28;x=0", "p=-1000000000000000/(243" & ChrW$(183) & "(x-238,9747)^5 )") Then GoTo slut

    ' This has previously given the wrong solution, as there are two solutions, but when the constant is inserted, only one fits.
    If TestSolveDE("(x+5)" & ChrW$(183) & "y^'=" & ChrW$(8730) & "y", "y=1;x=-4", "y=(ln" & ChrW$(8289) & "(|x+5|)+2)^2/4") Then GoTo slut

    
    'unit test
    MaximaUnits = True
    If TestBeregn("10 km/time", "=2,777778  m/s") Then GoTo slut
    If TestBeregn("2L", "=0,002 m^(3)") Then GoTo slut
    If TestBeregn("2ml", "=2" & ChrW$(183) & "10^-6  m^(3)") Then GoTo slut
    If TestBeregn("2dage", "=172800 s") Then GoTo slut
    If TestBeregn("1,23 kg" & ChrW$(183) & "(2,998" & ChrW$(183) & "10^8  m/s)^2", "=1,105524" & ChrW$(183) & "10^17  J") Then GoTo slut
    If TestSolve("60W=2,1kJ/t", "t", "t=35 s") Then GoTo slut
    MaximaUnits = False


ggbtest:
    ' GeoGebra test
    CASengineTempOnly = 2
    MaximaExact = 1 ' 1=exact
    '    Selection.TypeParagraph
    '    Selection.TypeText "GeoGebra CAS Test"
    ShowSettings "GeoGebra CAS Test"
    
    TestBeregn "1/5 2", "=2/5"
    If StopNow Then GoTo slut
    ' Denne er ikke justeret
    TestBeregn "2+3^3,4/log" & ChrW$(8289) & "(889) -sin" & ChrW$(8289) & "(34)", "-cos" & ChrW$(8289) & "(14/45" & ChrW$(183) & "" & ChrW$(960) & ")+(27" & ChrW$(12310) & "" & ChrW$(8730) & "(5&3)" & ChrW$(12311) & "^2)/(ln" & ChrW$(8289) & "(889)/ln" & ChrW$(8289) & "(10) )+2"
    If StopNow Then GoTo slut
    TestBeregn ChrW$(8731) & "(-8)", "=-2"
    If StopNow Then GoTo slut
    
    TestBeregn "a b/c+a b+f_a (x)", "=a" & ChrW$(183) & "b+a" & ChrW$(183) & "b/c+f_a (x)"
    If StopNow Then GoTo slut
    TestBeregn "f_a (x)+a (b+d)/c+f_c+2+a^x (2)", "=a" & ChrW$(183) & "(b+d)/c+f_a (x)+f_c+2a^(x)+2"
    If StopNow Then GoTo slut
    TestBeregn "2^2x+23/2x", "=(2^(x))^(2)+23/(2x)" ' fails if 2^2x is not interpreted as 2^(2*x)
    If StopNow Then GoTo slut
    TestBeregn "log" & ChrW$(8289) & "(a)", "=ln(a)/ln(10)"
    If StopNow Then GoTo slut
    
    TestBeregn "log_2" & ChrW$(8289) & "(4)", "=ln(4)/ln(2)@$2" ' does not reduce exactly with ggb
    If StopNow Then GoTo slut
    TestBeregn "ln" & ChrW$(8289) & "(a)", "=ln(a)"
    If StopNow Then GoTo slut
    TestBeregn "log_4" & ChrW$(8289) & "a", "=ln(a)/ln(4)"
    If StopNow Then GoTo slut
    '    TestBeregn chrw$(12310) & "sin" & chrw$(8289) & "(x)-sin" & chrw$(12311) & "" & chrw$(8289) & "(x_0 )/(x+y)", "=(-sin" & chrw$(8289) & "(x_0)+sin" & chrw$(8289) & "(1/180 " & chrw$(960) & "" & chrw$(183) & "x))/(x+y)" ' Test af forkert placerede skjulte parenteser
    TestBeregn ChrW$(12310) & "sin" & ChrW$(8289) & "(x)-sin" & ChrW$(12311) & "" & ChrW$(8289) & "(x_0 )/(x+y)", "=(-sin" & ChrW$(8289) & "(1/180 x_0 " & ChrW$(960) & ")+sin" & ChrW$(8289) & "(1/180" & ChrW(183) & ChrW$(960) & "" & ChrW$(183) & "x))/(x+y)"
    If StopNow Then GoTo slut
    '    TestBeregn chrw$(12310) & "sin" & chrw$(8289) & "(x)-sin" & chrw$(12311) & "" & chrw$(8289) & "(x_0 )/(x-x_0 )", "=(sin" & chrw$(8289) & "(x_0)-sin" & chrw$(8289) & "(1/180 " & chrw$(960) & "" & chrw$(183) & "x))/(x_0-x)"
    TestBeregn ChrW$(12310) & "sin" & ChrW$(8289) & "(x)-sin" & ChrW$(12311) & "" & ChrW$(8289) & "(x_0 )/(x-x_0 )", "=(sin" & ChrW$(8289) & "(1/180 x_0 " & ChrW$(960) & ")-sin" & ChrW$(8289) & "(1/180" & ChrW(183) & ChrW$(960) & "" & ChrW$(183) & "x))/(x_0-x)"
    If StopNow Then GoTo slut
    
    If TestBeregn("-20" & ChrW$(183) & "ln" & ChrW$(8289) & "(4)/ln" & ChrW$(8289) & "(10) +98", "=-20" & ChrW$(183) & "ln" & ChrW$(8289) & "(4)/ln" & ChrW$(8289) & "(10)+98") Then GoTo slut ' had problem with Bracket and equal sign

    ' Equations
    If TestSolve("x^2=9", "x", "x=-3    " & ChrW$(8744) & "    x=3") Then GoTo slut

    InsertTestMath "Definer: f(x,y)=x^2+y^2"
    If TestBeregn("f_x^' (x,y)", "=2x") Then GoTo slut ' test of partial derivative
    InsertTestMath "slet definitioner:"

    ' Calculus
    InsertTestMath "Definer: f(x)=2*x^3+e^(2x)"
    If TestBeregn("f^' (x)", "2e^2x+6x^2") Then GoTo slut
    If TestBeregn("(2" & ChrW$(183) & "x^3 )^'", "=6x^(2)") Then GoTo slut
    
    ' Differential equations
    If TestSolveDE("y^'+2x" & ChrW$(183) & "y=x", "y,x", "y=c_1 e^(-x^2 )+1/2") Then GoTo slut
    If TestSolveDE("y^'=b" & ChrW$(183) & "y" & ChrW$(183) & "(M-y)", "y,x", "y=-M e^(M" & ChrW$(183) & "b" & ChrW$(183) & "x)/(c_1-e^(M" & ChrW$(183) & "b" & ChrW$(183) & "x) )") Then GoTo slut
    
    'Fails
    MaximaExact = 0
    CASengineTempOnly = 0
    If TestSolve("(" & ChrW$(9608) & "(t^3-t@t))=(" & ChrW$(9608) & "((-15)/64@ 1/4))", "t", "t=1/4") Then GoTo slut ' This gives 3 solutions, but there is only one. It is directly in the 2nd coordinate. GeoGebra works

    
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
    Selection.TypeParagraph
    Selection.TypeText "Tid: " & (Timer - Tid) / 60 & " min"
    
    MaximaDecOutType = 2
    MaximaUnits = False
    MaximaExact = 0 ' Auto
    CASengineTempOnly = 0 ' Maxima
    SettUseVBACAS = True
    
    AllR.End = Selection.End
    AllR.Select
    
    Unload UFwait2
End Sub
Function StopNow() As Boolean
    If omax.StopNow Then
        StopNow = True
        Exit Function
    End If
    If UFwait2.StopNow Then
        omax.StopNow = True
        StopNow = True
        Exit Function
    End If
    If ErrCount = 5 And ContCount = 0 Then
        If MsgBox("5 errors detected. Do you want to continue?", vbYesNo, "Many errors") = vbNo Then
            StopNow = True
            Exit Function
         Else
            ContCount = ContCount + 1
         End If
    ElseIf ErrCount = 10 And ContCount <= 1 Then
        If MsgBox("10 errors detected. Do you still want to continue?", vbYesNo, "Many errors") = vbNo Then
            StopNow = True
            Exit Function
         Else
            ContCount = ContCount + 1
        End If
    ElseIf ErrCount = 50 And ContCount <= 2 Then
        If MsgBox("50 errors detected. Do you still want to continue?", vbYesNo, "Many errors") = vbNo Then
            StopNow = True
            Exit Function
         Else
            ContCount = ContCount + 1
        End If
    End If
    StopNow = False
End Function
Sub PerformTest(TestType As Integer, komm As String, resul As String, Optional Var As String, Optional Instruk As String)
    Dim s As String, TypeText As String, Oresul As String, Oresul2 As String, Arr() As String, ResultOK As Boolean, i As Integer, p As Integer, p2 As Integer, res As String
    If TestType = 1 Then
        TypeText = "Calculate"
    ElseIf TestType = 2 Then
        TypeText = "Solve"
    ElseIf TestType = 3 Then
        TypeText = "SolveDE"
    End If
    s = TestCount & ": " & TypeText & vbCrLf & "Error count: " & ErrCount
    UFwait2.Label1.Caption = s
    UFwait2.Label_tip.Caption = komm
    If Instruk <> "" Then
        UFwait2.Label_tip.Caption = komm & vbCrLf & Instruk
    End If
    InsertTestMath komm, False
    DoEvents
    If TestType = 1 Then
        beregn
    ElseIf TestType = 2 Then
        MaximaSolvePar (Var)
    ElseIf TestType = 3 Then 'solvede
        Arr = Split(Var, ";")
        If UBound(Arr) < 1 Then Arr = Split(Var, ",")
        SolveDEpar Arr(0), Arr(1)
    End If
    Wait 0.2
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    MoveCursorToEndOfCalculation
        
    If TestType = 1 Then ' when calculating, the result is written together with the input, so comparison is eliminated
        Oresul2 = omax.MaximaOutput
        GotoPrevEq
        omax.ReadSelection
        Oresul = TrimR(omax.Kommando, vbCr)
        p = InStrRev(Oresul, "=")
        p2 = InStrRev(Oresul, ChrW$(8776))
        If p2 > p Then p = p2
        If p > 0 Then Oresul = Right$(Oresul, Len(Oresul) - p)
        MoveCursorToEndOfCalculation False
    Else
        GotoPrevEq
        If Selection.OMaths.Count > 0 Then
            If HasAssumptions Then
                Selection.MoveRight 2
                GotoNextEq
            End If
        End If
        omax.ReadSelection
        Oresul = TrimR(omax.Kommando, vbCr)
        MoveCursorToEndOfCalculation False
    End If
    Oresul = Trim$(Oresul)
    If Left$(Oresul, 1) = "=" Or Left$(Oresul, 1) = ChrW$(8776) Then Oresul = Right$(Oresul, Len(Oresul) - 1)
    ResultOK = False
    Arr = Split(resul, "@$")
    For i = 0 To UBound(Arr)
        res = Trim$(Arr(i))
        If Left$(res, 1) = "=" Or Left$(res, 1) = ChrW$(8776) Then res = Right$(res, Len(res) - 1)
        If res = Oresul Or Trim$(Arr(i)) = Oresul2 Then
            ResultOK = True
            Exit For
        End If
    Next
    
    If Not ResultOK Then 'omax.MaximaOutput
        Selection.Font.ColorIndex = wdRed
        Selection.Font.Bold = True
        Selection.TypeText (TypeText & " error. Expected result: ")
        Selection.Font.Bold = False
        Selection.Font.ColorIndex = wdAuto
        Selection.TypeText ("  ")
        Selection.TypeParagraph
        Selection.TypeText "resul(forventet resultat):   " & resul
        Selection.TypeParagraph
        Selection.TypeText "Oresul(egentlige resultat):  " & Oresul
        Selection.TypeParagraph
        Selection.TypeText "Brug denne kode-streng:"
        Selection.TypeParagraph
        If TestType = 1 Then
            Selection.TypeText "TestBeregn " & Trim$(ConvertToVBAString(komm)) & " , " & Trim$(ConvertToVBAString(Oresul))
        ElseIf TestType = 2 Then
            Selection.TypeText vbTab & "If TestSolve(" & Trim$(ConvertToVBAString(komm)) & " , """ & Var & """ , " & Trim$(ConvertToVBAString(Oresul)) & ") Then GoTo slut"
        ElseIf TestType = 3 Then
            Selection.TypeText vbTab & "If TestSolveDE(" & Trim$(ConvertToVBAString(komm)) & " , """ & Var & """ , " & Trim$(ConvertToVBAString(Oresul)) & ") Then GoTo slut"
        End If
        Selection.TypeParagraph
        ErrCount = ErrCount + 1
    ElseIf visok Then
        Selection.Font.ColorIndex = wdGreen
        Selection.Font.Bold = True
        Selection.TypeText (TypeText & " test was successful")
        Selection.Font.Bold = False
        Selection.Font.ColorIndex = wdAuto
        Selection.TypeParagraph
    End If
    UFwait2.Label_progress.Caption = UFwait2.Label_progress.Caption & "*"
    TestCount = TestCount + 1
End Sub
Sub CreateTestBeregn()
' laver en ny test
    CreateTestBeregnPar
End Sub
Sub CreateTestBeregnDE()
    CreateTestBeregnPar 4
End Sub

Sub CreateTestBeregnPar(Optional TestType As Integer = 0)
    Dim Oresul As String, Arr() As String
    Dim komm As String, Var As String
    
    PrepareMaxima
    
    omax.ReadSelection
    komm = TrimR(omax.Kommando, vbCr)
    DoEvents
    If InStr(LCase$(komm), "define") > 0 Or TestType = 3 Then
        TestType = 3
    ElseIf InStr(komm, "y^'=") > 0 Or TestType = 4 Then
        TestType = 4
        Var = InputBox("Enter dependent and independent variable to DEsolve for. Separate by semicolon. You can add initial condition." & vbCrLf & "Examples: y;x " & vbCrLf & "y=5;x=0", "Variable", "y;x")
        If Trim$(Var) = vbNullString Then Exit Sub
        Arr = Split(Var, ";")
        If UBound(Arr) < 1 Then Exit Sub
        SolveDEpar Trim$(Arr(0)), Trim$(Arr(1))
    ElseIf InStr(komm, "=") > 0 Or TestType = 2 Then
        TestType = 2
        Var = InputBox("Enter variable to solve for", "Variable", "x")
        If Trim$(Var) = vbNullString Then Exit Sub
        MaximaSolvePar (Var)
    Else
        TestType = 1
        beregn
    End If
    '   InsertTestMath komm
    Wait 0.5
    MoveCursorToEndOfCalculation
        
    If TestType = 1 Then ' when calculating, the result is written together with the input, so comparison is eliminated
        Oresul = omax.MaximaOutput
    Else
        GotoPrevEq
        If Selection.OMaths.Count > 0 Then
            If HasAssumptions Then
                Selection.MoveRight 2
                GotoNextEq
            End If
        End If
        omax.ReadSelection
        Oresul = TrimR(omax.Kommando, vbCr)
        '      omax.Kommando = komm & "=" & Split(omax.Kommando, "=", 1)(1)
        MoveCursorToEndOfCalculation False
    End If
        
    If TestType = 1 Then
        Selection.TypeParagraph
        Selection.TypeText "Insert this code string into sub 'RunTestSequence':"
        Selection.TypeParagraph
        Selection.TypeText vbTab & "If TestBeregn(" & Trim$(ConvertToVBAString(komm)) & " , " & Trim$(ConvertToVBAString(Oresul)) & ") Then GoTo slut"
    ElseIf TestType = 2 Then
        Selection.TypeParagraph
        Selection.TypeText "Insert this code string into sub 'RunTestSequence':"
        Selection.TypeParagraph
        Selection.TypeText vbTab & "If TestSolve(" & Trim$(ConvertToVBAString(komm)) & " , """ & Var & """ , " & Trim$(ConvertToVBAString(Oresul)) & ") Then GoTo slut"
    ElseIf TestType = 3 Then
        Selection.TypeParagraph
        Selection.TypeText "Insert this code string into sub 'RunTestSequence':"
        Selection.TypeParagraph
        Selection.TypeText "InsertTestMath " & Trim$(ConvertToVBAString(komm))
    ElseIf TestType = 4 Then
        Selection.TypeParagraph
        Selection.TypeText "Insert this code string into sub 'RunTestSequence':"
        Selection.TypeParagraph
        Selection.TypeText vbTab & "If TestSolveDE(" & Trim$(ConvertToVBAString(komm)) & " , """ & Var & """ , " & Trim$(ConvertToVBAString(Oresul)) & ") Then GoTo slut"
    End If
    Selection.TypeParagraph

End Sub
Sub GetTestString()
' place the cursor in the result of a calculation, then this sub gives the expression that performtest should have as a result
    Dim Oresul As String, s As String
    omax.ReadSelection
    Oresul = TrimR(omax.Kommando, vbCr)
    
   s = ConvertToVBAString(Oresul)
    
    Selection.Collapse wdCollapseEnd
    Selection.EndKey unit:=wdLine
    Selection.TypeParagraph
    Selection.TypeText (s)

End Sub
Function ConvertToVBAString(text As String) As String
    Dim s As String, j As Integer, i As Integer
    s = ""
    For j = 1 To Len(text)
        i = AscW(Mid$(text, j, 1))
        If i > 200 Or i = 183 Then
            s = s & """ & chrw$(" & i & ") & """
        Else
            s = s & Mid$(text, j, 1)
        End If
    Next
    If Left$(s, 4) = """ & " Then
        s = Right$(s, Len(s) - 4)
    ElseIf Left$(s, 1) <> """" Then
        s = """" & s
    End If
    If Right$(s, 4) = " & """ Then
        s = Left$(s, Len(s) - 4)
    ElseIf Right$(s, 1) <> """" Then
        s = s & """"
    End If
    ConvertToVBAString = s
End Function
Function TestBeregn(komm As String, resul As String) As Boolean
    PerformTest 1, komm, resul
    If StopNow Then TestBeregn = True ' betyder stop
End Function
Function TestSolveDE(komm As String, Var As String, resul As String) As Boolean
    PerformTest 3, komm, resul, Var
    If StopNow Then TestSolveDE = True ' betyder stop
End Function
Sub GotoPrevEq()
Dim i As Integer
    Do While Selection.OMaths.Count = 0 And i < 100
        Selection.GoToPrevious (wdGoToLine)
        Selection.EndKey unit:=wdLine
        i = i + 1 ' there are some equations where it just gets stuck on the same line. Something with vectors
    Loop
End Sub
Sub GotoNextEq()
Dim i As Integer
    Do While Selection.OMaths.Count = 0 And i < 100
        Selection.GoToNext (wdGoToLine)
        Selection.EndKey unit:=wdLine
        i = i + 1 ' there are some equations where it just gets stuck on the same line. Something with vectors
    Loop
End Sub
Function TestSolve(komm As String, Var As String, resul As String, Optional Instruk As String) As Boolean
    PerformTest 2, komm, resul, Var, Instruk
    If StopNow Then TestSolve = True ' means stop
End Function

Sub TestSolve2(komm As String, Var As String, resul As String)
    omax.Kommando = komm
    DoEvents
    omax.MaximaSolve (Var)
    Wait 0.5
    If Not omax.MaximaOutput = resul Then
        Selection.TypeText (" - Fejl - " & omax.Kommando & " - " & omax.MaximaOutput)
        Selection.TypeParagraph
    ElseIf visok Then
        Selection.TypeText (" - ok - " & omax.Kommando)
        Selection.TypeParagraph
    End If

End Sub

Sub InsertTestMath(s As String, Optional NewLine As Boolean = True)
' inserts maxima output into word document
    Dim mo As Range

    Selection.Range.Font.Bold = False

    If Len(s) > 500 Then
'        If MsgBox("Output is " & Len(s) & " characters long. Click OK to insert all characters in Word. Cancel to insert the first 500", vbOKCancel, TT.Warning) = vbCancel Then
        s = Left$(s, 500) & " " & TT.A(119)
    End If
    Set mo = Selection.OMaths.Add(Selection.Range)
    Selection.TypeText s
    mo.OMaths.BuildUp
    If NewLine Then
        Selection.EndKey unit:=wdLine
        Selection.Collapse wdCollapseEnd
        Selection.TypeParagraph
    End If
    DoEvents
End Sub
Sub MoveCursorToEndOfCalculation(Optional AddLine As Boolean = True)
    If Not Selection.Find.Execute("*^13^13", , , True) Then
        '    Selection.GoToNext (wdGoToLine)
        Selection.EndKey unit:=wdLine
        Selection.Collapse wdCollapseEnd
        Selection.TypeParagraph
        Selection.TypeParagraph
        Selection.TypeParagraph
        Selection.TypeParagraph
        Selection.GoToPrevious (wdGoToLine)
        Selection.GoToPrevious (wdGoToLine)
        Selection.EndKey unit:=wdLine
    End If
    Selection.Collapse wdCollapseEnd
    If AddLine Then
        Selection.TypeParagraph
        Selection.GoToPrevious (wdGoToLine)
    End If
    Selection.GoToPrevious (wdGoToLine)
    Selection.EndKey unit:=wdLine

End Sub
Sub InsertHeadingtext(Htext As String)
    Selection.TypeParagraph
    Selection.Font.Size = 18
    Selection.ParagraphFormat.SpaceAfter = 0
    Selection.Font.Bold = True
    Selection.TypeText Htext
    Selection.Font.Bold = False
    Selection.TypeParagraph
    Selection.Font.Size = 11
    Selection.ParagraphFormat.SpaceAfter = 6
End Sub

Sub ShowSettings(Optional Htext As String)
    If Htext <> "" Then InsertHeadingtext Htext
    Selection.TypeParagraph
    Selection.Font.Size = 9
    Selection.ParagraphFormat.SpaceAfter = 0
    Selection.Font.Bold = True
    Selection.TypeText "Settings"
    Selection.Font.Bold = False
    Selection.TypeParagraph
    Selection.TypeText "CAS engine: " & CASengine & vbTab & "All Trig. solutions: " & AllTrig
    Selection.TypeParagraph
    Selection.TypeText "Exact: " & MaximaExact & vbTab & "Cifre: " & MaximaCifre
    Selection.TypeParagraph
    Selection.TypeText "Decimaloutput: " & MaximaDecOutType
    Selection.TypeText "Scientific notation: " & (MaximaDecOutType = 3) & vbTab & "Units: " & MaximaUnits
    Selection.TypeParagraph
    Selection.TypeText "Radians: " & Radians & vbTab & "Complex: " & MaximaComplex
    Selection.Font.Size = 11
    Selection.ParagraphFormat.SpaceAfter = 6
    Selection.TypeParagraph
End Sub
Sub InsertTestComment(s As String)
    Selection.TypeText (s)
    Selection.TypeParagraph
End Sub
Sub unicodevals()
    Dim s As String
    Dim i As Integer
    Dim c As Range
    MsgBox Selection.text
    For Each c In Selection.Characters
        i = AscW(c)
        s = s & c & " - " & i & vbCrLf
    Next

    MsgBox s
End Sub

Sub unicodevals2()
    Dim text As String
    Dim i As Integer
    Dim j As Integer
    Dim s As String
    If Selection.OMaths.Count > 0 Then
        Selection.OMaths.Linearize
        Selection.OMaths(1).ConvertToNormalText
        text = Selection.text
        Selection.OMaths(1).ConvertToMathText
        Selection.OMaths(1).Range.Select
        Selection.OMaths.BuildUp
    Else
        text = Selection.text
    End If

    For j = 1 To Len(text)
        i = AscW(Mid$(text, j, 1))
        s = s & Mid$(text, j, 1) & " - " & i & vbCrLf
    Next
    MsgBox s

End Sub
Sub UnicodeValsToString()
' converts all Omaths in selection to a string that can be inserted into VBA code. Used primarily for test modules
' The strings are inserted after selection in order. Each on a new line
    Dim text As String
    Dim j As Integer
    Dim i As Integer
    Dim k As Integer, n As Integer
    Dim s As String
    Dim mo As OMath
    Dim Arr() As String
    Dim MoArr() As Variant

    n = Selection.OMaths.Count
    If n = 0 Then
        MsgBox "You must select an equation", vbOKOnly, "Error"
        Exit Sub
    End If
    ReDim Arr(n - 1)
    ReDim MoArr(n - 1)
'    Selection.OMaths.Linearize
    For k = 0 To n - 1
        Set MoArr(k) = Selection.OMaths(k + 1)
    Next
    For k = 0 To n - 1
        Set mo = MoArr(k)
        mo.Linearize
        mo.ConvertToNormalText
        Arr(k) = Trim$(mo.Range.text)
        mo.ConvertToMathText
        mo.Range.Select
        mo.BuildUp
    Next
    Selection.Collapse wdCollapseEnd
    Selection.EndKey unit:=wdLine

    For k = 0 To UBound(Arr)
        text = Arr(k)
        s = ""
        For j = 1 To Len(text)
            i = AscW(Mid$(text, j, 1))
            If i > 200 Or i = 183 Then
                s = s & """ & chrw$(" & i & ") & """
            Else
                s = s & Mid$(text, j, 1)
            End If
        Next
        If Left$(s, 4) = """ & " Then
            s = Right$(s, Len(s) - 4)
        ElseIf Left$(s, 1) <> """" Then
            s = """" & s
        End If
        If Right$(s, 4) = " & """ Then
            s = Left$(s, Len(s) - 4)
        ElseIf Right$(s, 1) <> """" Then
            s = s & """"
        End If
        Selection.Collapse wdCollapseEnd
        Selection.TypeParagraph
        Selection.TypeText (s)
    Next

End Sub

