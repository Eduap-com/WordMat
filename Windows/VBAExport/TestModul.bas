Attribute VB_Name = "TestModul"
Option Explicit
    Private visok As Boolean
    Private UFwait2 As UserFormWaitForMaxima
    Private ErrCount As Long
    Private TestCount As Long
    Private ContCount As Integer
    Private NonInterA As Boolean

' Brug CreateTestBeregn til at lave nye test

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
    
    On Error GoTo Fejl
    ErrCount = 0
    TestCount = 0
    ContCount = 0
    visok = True
    
    If MsgBox("Are you sure want to conduct a test. The document will be filled with calculations. It can take some time.", vbOKCancel, "Confirm") = vbCancel Then Exit Sub
        
    Set UFwait2 = New UserFormWaitForMaxima
    UFwait2.Label1.Font.Size = 12
    UFwait2.Label_tip.Font.Size = 10
    UFwait2.Show vbModeless
    
    ' Testresultaterne er baseret på bestemte indstillinger. Undervejs bliver de måske ændret, men i starten skal de nulstilles
    CASengineTempOnly = 0 ' 0=maxima
    MaximaExact = 1 ' 0 - auto ' 1 - exact ' 2 - num
    MaximaCifre = 7
    MaximaDecOutType = 2 ' bc
    Radians = False
    MaximaComplex = False
    AllTrig = False
    MaximaLogOutput = 0
    
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

    'GoTo ggbtest
    
    ' når der skal laves nye test, er det nemmest at bruge funktion 'CreateTestberegn'
    ' alternativt: bare at køre testen med et tomt expected result. Der skrives hvilken teststreng der skal bruges hvis resultat er korrekt
    ' ellers brug GetTestString' eller UnicodeValsToString for at finde streng fra et matematikfelt for kommandoen.
    ' the result may have to be adjusted for calculations (not solve). Often brackets of exponents are different and spaces must be removed
    ' Multiple correct answers can be separated by @$
    '    DoEvents ' virker ikke
    '    Application.ScreenRefresh
    
    'til test af enkelt
    
'    GoTo slut
    
    
    DebugWM = False
    ' ikke den kønneste løsning, men tomme linjer bliver spist og hvis der ikke er plads som beregninger kan vokse i, så giver det problemer.
    For i = 0 To 30
    Selection.TypeParagraph
    Next
    Selection.MoveLeft wdCharacter, 30
    
    InsertSletDef
    ' Man kan indtaste flere korrekte resultater, bare adskil med @$   ----------------------------------------
    'calculation tests
    If TestBeregn("2+3", "=5") Then GoTo slut
    TestBeregn "2+3^3,4/log" & VBA.ChrW(8289) & "(889) -sin" & VBA.ChrW(8289) & "(34)", "-sin" & VBA.ChrW(8289) & "((17" & VBA.ChrW(183) & "" & VBA.ChrW(960) & ")/90)+(41,89983" & VBA.ChrW(183) & "ln" & VBA.ChrW(8289) & "(10))/ln" & VBA.ChrW(8289) & "(889) +2"

'    TestBeregn "2+3^3,4/log" & VBA.ChrW(8289) & "(889) -sin" & VBA.ChrW(8289) & "(34)", "=(-sin" & VBA.ChrW(8289) & "((17" & VBA.ChrW(183) & "" & VBA.ChrW(960) & ")/90))+(41,89983049571472" & VBA.ChrW(183) & "ln" & VBA.ChrW(8289) & "(10))/ln" & VBA.ChrW(8289) & "(889)+2@$=-sin" & VBA.ChrW(8289) & "((17" & VBA.ChrW(183) & "" & VBA.ChrW(960) & ")/90)+(41,899830495714724" & VBA.ChrW(183) & "ln" & VBA.ChrW(8289) & "(10))/ln" & VBA.ChrW(8289) & "(889)+2"
    If StopNow Then GoTo slut
    TestBeregn "1/5 2", "=2/5"
    If StopNow Then GoTo slut
    TestBeregn VBA.ChrW(8731) & "(-8)", "=-2"
    If StopNow Then GoTo slut
    TestBeregn "a b/c+a b+f_a (x)", "=f_a (x)+(a" & VBA.ChrW(183) & "b)/c+a" & VBA.ChrW(183) & "b"
    If StopNow Then GoTo slut
    TestBeregn "f_a (x)+a (b+d)/c+f_c+2+a^x (2)", "=f_a (x)+f_c+(a" & VBA.ChrW(183) & "(d+b))/c+2" & VBA.ChrW(183) & "a^(x)+2"
    If StopNow Then GoTo slut
    TestBeregn "2^2x+23/2x", "=2^(2" & VBA.ChrW(183) & "x)+23/(2" & VBA.ChrW(183) & "x)" ' fails if 2^2x is not interpreted as 2^(2*x)
    If StopNow Then GoTo slut
    TestBeregn "log" & VBA.ChrW(8289) & "(a)", "=ln" & VBA.ChrW(8289) & "(a)/ln" & VBA.ChrW(8289) & "(10)"
    If StopNow Then GoTo slut
    TestBeregn "log_2" & VBA.ChrW(8289) & "(4)", "=2"
    If StopNow Then GoTo slut
    TestBeregn "ln" & VBA.ChrW(8289) & "(a)", "=ln" & VBA.ChrW(8289) & "(a)"
    If StopNow Then GoTo slut
    TestBeregn "log_4" & VBA.ChrW(8289) & "a", "=ln" & VBA.ChrW(8289) & "(a)/ln" & VBA.ChrW(8289) & "(4)"
    If StopNow Then GoTo slut
    TestBeregn VBA.ChrW(12310) & "sin" & VBA.ChrW(8289) & "(x)-sin" & VBA.ChrW(12311) & "" & VBA.ChrW(8289) & "(x_0 )/(x+y)", "=(sin" & VBA.ChrW(8289) & "((" & VBA.ChrW(960) & "" & VBA.ChrW(183) & "x)/180)-sin" & VBA.ChrW(8289) & "((" & VBA.ChrW(960) & "" & VBA.ChrW(183) & "x_0)/180))/(y+x)" ' Test af forkert placerede skjulte parenteser
    If StopNow Then GoTo slut
    TestBeregn VBA.ChrW(12310) & "sin" & VBA.ChrW(8289) & "(x)-sin" & VBA.ChrW(12311) & "" & VBA.ChrW(8289) & "(x_0 )/(x-x_0 )", "=(sin" & VBA.ChrW(8289) & "((" & VBA.ChrW(960) & "" & VBA.ChrW(183) & "x)/180)-sin" & VBA.ChrW(8289) & "((" & VBA.ChrW(960) & "" & VBA.ChrW(183) & "x_0)/180))/(x-x_0)"
    If StopNow Then GoTo slut
    TestBeregn "x/" & VBA.ChrW(12310) & "2+x" & VBA.ChrW(12311) & "^2", "=x/(x^(2)+2)"
    If StopNow Then GoTo slut
    TestBeregn "|2+(1+|2-3|)/(|2|+1)+1|", "=11/3"
    If StopNow Then GoTo slut
    TestBeregn "L=100^(1/5)/" & VBA.ChrW(960) & "^(1/5)", "=100^(1/5)/" & VBA.ChrW(960) & "^(1/5)"
    If StopNow Then GoTo slut
    TestBeregn "cos^(-1)" & VBA.ChrW(8289) & "(2)", "=(180" & VBA.ChrW(183) & "cos^(-1)" & VBA.ChrW(8289) & "(2))/" & VBA.ChrW(960) ' burde måske give fejl
    If StopNow Then GoTo slut
    TestBeregn "(" & VBA.ChrW(8730) & "((x+1)^2+(2" & VBA.ChrW(183) & "(x+1)/x)^2 ))^'", "=((-(8" & VBA.ChrW(183) & "(x+1)^(2))/x^(3))+(8" & VBA.ChrW(183) & "(x+1))/x^(2)+2" & VBA.ChrW(183) & "(x+1))/(2" & VBA.ChrW(183) & "" & VBA.ChrW(8730) & "((4" & VBA.ChrW(183) & "(x+1)^(2))/x^(2)+(x+1)^(2)))@$=(-((8" & VBA.ChrW(183) & "(x+1)^(2))/x^(3))+(8" & VBA.ChrW(183) & "(x+1))/x^(2)+2" & VBA.ChrW(183) & "(x+1))/(2" & VBA.ChrW(183) & "" & VBA.ChrW(8730) & "((4" & VBA.ChrW(183) & "(x+1)^(2))/x^(2)+(x+1)^(2)))"
    If StopNow Then GoTo slut
    
    InsertSletDef
    ' equation solving
    InsertHeadingtext "Ligningsløsning"
    TestSolve "x^2=9", "x", "x=-3    " & VBA.ChrW(8744) & "    x=3"
    If StopNow Then GoTo slut
    TestSolve "Hovedstol=Ydelse" & VBA.ChrW(8729) & "" & VBA.ChrW(12310) & "1-(1+r)" & VBA.ChrW(12311) & "^(-n)/r", "n", "n=-ln" & VBA.ChrW(8289) & "((Ydelse-Hovedstol" & VBA.ChrW(183) & "r)/Ydelse)/ln" & VBA.ChrW(8289) & "(r+1) @$n=-(ln" & VBA.ChrW(8289) & "((Ydelse-Hovedstol" & VBA.ChrW(183) & "r)/Ydelse)/ln" & VBA.ChrW(8289) & "(r+1) )"
    If StopNow Then GoTo slut
    TestSolve VBA.ChrW(8747) & "_0^a" & VBA.ChrW(9618) & "x^2 dx=4", "a", "a=12^(1/3)"
    If StopNow Then GoTo slut
    TestSolve VBA.ChrW(8747) & "_0^2" & VBA.ChrW(9618) & "" & VBA.ChrW(12310) & "a·x·(x-2)^2 " & VBA.ChrW(12311) & " dx=20", "a", "a=15"
    If StopNow Then GoTo slut
    TestSolve "92,3*(e^(0,21278*x)-1)/(e^(0,21278*x)+1)=0.5", "x", "x=0,05091805"
    If StopNow Then GoTo slut
    TestSolve "x^3-x^2+x-2=0", "x", "x=1,35321"
    If StopNow Then GoTo slut
    TestSolve "x^4-x^3+x^2-x-1=0", "x", "x=-0,5187901    " & VBA.ChrW(8744) & "    x=1,290649"
    If StopNow Then GoTo slut
    TestSolve "x-x=2", "x", "x" & VBA.ChrW(8712) & "" & VBA.ChrW(8709)
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
    TestSolve "120=429" & VBA.ChrW(183) & "L^2,920", "L", "L=0,6464306"
    If StopNow Then GoTo slut
    TestSolve "x^17,3=1/2", "x", "x=0,9607257"
    If StopNow Then GoTo slut
    TestSolve "1,15=" & VBA.ChrW(12310) & "0,9" & VBA.ChrW(12311) & "^a", "a", "a=-1,326512"
    If StopNow Then GoTo slut
    TestSolve "y2/y1=(x2/x1)^a", "a", "a=ln" & VBA.ChrW(8289) & "(y2/y1)/ln" & VBA.ChrW(8289) & "(x2/x1)"
    If StopNow Then GoTo slut
    TestSolve "K_n=K_0" & VBA.ChrW(183) & "(1+r)^n", "r", "r=(K_n/K_0 )^(1/n)-1"
    If StopNow Then GoTo slut
    TestSolve "(sin" & VBA.ChrW(8289) & "(A))/a=(sin(B))/b", "A", "A=sin^(-1)" & VBA.ChrW(8289) & "((sin" & VBA.ChrW(8289) & "(B)" & VBA.ChrW(183) & "a)/b)"
    If StopNow Then GoTo slut
    TestSolve "tan" & VBA.ChrW(8289) & "(A)=x/50", "A", "A=tan^(-1)" & VBA.ChrW(8289) & "(x/50)"
    If StopNow Then GoTo slut
    TestSolve "L=10" & VBA.ChrW(183) & "log" & VBA.ChrW(8289) & "(I/I_0 )", "I", "I=I_0" & VBA.ChrW(183) & "10^(L/10)"
    If StopNow Then GoTo slut
    TestSolve "4" & VBA.ChrW(183) & "" & VBA.ChrW(8730) & "(x-1)=-1/5" & VBA.ChrW(183) & "x^2+2" & VBA.ChrW(183) & "x+3", "x", "x=5" ' skal kun give denne ene løsning
    If StopNow Then GoTo slut
    TestSolve "x^2" & VBA.ChrW(183) & "(x^2-1)=9" & VBA.ChrW(183) & "(x^2-1)", "x", "x=-3    " & VBA.ChrW(8744) & "    x=-1    " & VBA.ChrW(8744) & "    x=1    " & VBA.ChrW(8744) & "    x=3"
    If StopNow Then GoTo slut
    TestSolve "54.4=(0.01767+2" & VBA.ChrW(183) & "x)^2/((0.00600-x)(0.00313-x))", "x", "x=0,001432725    " & VBA.ChrW(8744) & "    x=0,00982426" ' earlier problem with missing multsign in denominator
    If StopNow Then GoTo slut
    TestSolve VBA.ChrW(8747) & "_0^a" & VBA.ChrW(9618) & "" & VBA.ChrW(12310) & "x^2+2dx=13" & VBA.ChrW(12311), "a", "a=2,808381"
    If StopNow Then GoTo slut
    TestSolve VBA.ChrW(8747) & "_0^2" & VBA.ChrW(9618) & "" & VBA.ChrW(12310) & "a" & VBA.ChrW(183) & "x" & VBA.ChrW(183) & "(x-2)^2 " & VBA.ChrW(12311) & " dx=20", "a", "a=15"
    If StopNow Then GoTo slut
    TestSolve "0=(x^3+b)/x^2", "x", "x=-b^(1/3)"
    If StopNow Then GoTo slut
    TestSolve "0,1" & VBA.ChrW(183) & "x^4+0,8" & VBA.ChrW(183) & "x^3+0,6" & VBA.ChrW(183) & "x^2-2" & VBA.ChrW(183) & "x-1=0", "x", "x=-6,689559    " & VBA.ChrW(8744) & "    x=-2,245805    " & VBA.ChrW(8744) & "    x=-0,4727184    " & VBA.ChrW(8744) & "    x=1,408082"
    If StopNow Then GoTo slut
    TestSolve "(x+1,5)^2+ (3/x+2)^2=25", "x", "x=-6,263041    " & VBA.ChrW(8744) & "    x=-0,4357036    " & VBA.ChrW(8744) & "    x=1,5    " & VBA.ChrW(8744) & "    x=2,198745"
    If StopNow Then GoTo slut
    TestSolve "2" & VBA.ChrW(183) & "3^x=4" & VBA.ChrW(183) & "5^2x", "x", "x=ln" & VBA.ChrW(8289) & "(2)/ln" & VBA.ChrW(8289) & "(3/25)"
    If StopNow Then GoTo slut
    TestSolve "2" & VBA.ChrW(183) & "" & VBA.ChrW(12310) & "1,1" & VBA.ChrW(12311) & "^x=3" & VBA.ChrW(183) & "" & VBA.ChrW(12310) & "1,05" & VBA.ChrW(12311) & "^x", "x", "x=ln" & VBA.ChrW(8289) & "(3/2)" & VBA.ChrW(183) & "21,49612@$x=21,49612" & VBA.ChrW(183) & "ln" & VBA.ChrW(8289) & "(3/2)"
    If StopNow Then GoTo slut
    TestSolve "2" & VBA.ChrW(183) & "" & VBA.ChrW(12310) & "(11/10)" & VBA.ChrW(12311) & "^x=3" & VBA.ChrW(183) & "" & VBA.ChrW(12310) & "(105/100)" & VBA.ChrW(12311) & "^x", "x", "x=ln" & VBA.ChrW(8289) & "(3/2)/ln" & VBA.ChrW(8289) & "(22/21)"
    If StopNow Then GoTo slut
    TestSolve "e^(3,044522 x)=0,6666667 e^(3,091042 x)", "x", "x=8,71593"
    If StopNow Then GoTo slut
    TestSolve "92,3*(e^(0,21278*x)-1)/(e^(0,21278*x)+1)=0.5", "x", "x=0,05091805"
    If StopNow Then GoTo slut
    TestSolve "x^2+x-12" & VBA.ChrW(8805) & "8", "x", "x" & VBA.ChrW(8804) & "-5   " & VBA.ChrW(8744) & "   x" & VBA.ChrW(8805) & "4" ' ulighed
    If StopNow Then GoTo slut
    If TestSolve("e^(-x)=2", "x", "x=-ln" & VBA.ChrW(8289) & "(2)") Then GoTo slut
    If TestSolve("1=1,5" & VBA.ChrW(183) & "x^2+0,5/x^0,5 -1", "x", "x=0,06287224    " & VBA.ChrW(8744) & "    x=1") Then GoTo slut ' kræver numerisk løsning. Tidligere to_poly fejl, med kun en løsning.
    If TestSolve("6174/(1+8652" & VBA.ChrW(183) & "" & VBA.ChrW(12310) & "0,711" & VBA.ChrW(12311) & "^x )=8000000", "x", "x" & VBA.ChrW(8712) & "" & VBA.ChrW(8709)) Then GoTo slut
    If TestSolve("x^2=e^(-x^2 )", "x", "x=-" & VBA.ChrW(8730) & "(lambert_w (1) )    " & VBA.ChrW(8744) & "    x=" & VBA.ChrW(8730) & "(lambert_w (1) )") Then GoTo slut

    MaximaExact = 2 ' num
    InsertTestMath "Definer: " & VBA.ChrW(963) & ">0"
    If TestSolve("0,1=" & VBA.ChrW(8747) & "_(-" & VBA.ChrW(8734) & ")^5" & VBA.ChrW(9618) & "1/(" & VBA.ChrW(8730) & "2" & VBA.ChrW(960) & "" & VBA.ChrW(183) & "" & VBA.ChrW(963) & ")" & VBA.ChrW(183) & "e^(-1/2" & VBA.ChrW(183) & "((y-7)/" & VBA.ChrW(963) & ")^2 ) dy", "sigma", VBA.ChrW(963) & "=1,560608") Then GoTo slut
    InsertSletDef
    InsertTestMath "Definer: f(x)={" & VBA.ChrW(9608) & "(100" & VBA.ChrW(183) & "" & VBA.ChrW(12310) & "1,03" & VBA.ChrW(12311) & "^x,  0" & VBA.ChrW(8804) & "x" & VBA.ChrW(8804) & "72@33743" & VBA.ChrW(183) & "" & VBA.ChrW(12310) & "0,95" & VBA.ChrW(12311) & "^x,  72<x" & VBA.ChrW(8804) & "150)" & VBA.ChrW(9508)
    If TestSolve("f(x)=200", "x", "x=23,44977    " & VBA.ChrW(8744) & "    x=99,97819") Then GoTo slut ' stykkevis kan ikke håndteres algebraisk, så solvereal skal gå direkte til nsolve
    InsertSletDef
    
    MaximaExact = 1
    
    '    TestSolve "", "x", "x="
    '    If StopNow Then GoTo slut
    
    'Often fails, but not always ??? then works numerically. Det er noget med en overflow fejl. Virker ikke mere numerisk
    ' resultat er numerisk 100. Nogle gange kører denne i lykke og skal gentages numerisk. Ukendt hvorfor.
    If Not NonInterA Then
        InsertTestMath "OBS: Denne kan tage lang tid. Vent med at trykke STOP"
        Selection.TypeParagraph
        DoEvents
        If TestSolve("40=72" & VBA.ChrW(183) & "e^((0,619/0,22" & VBA.ChrW(8729) & "(e^22-e^0,22t )) )", "t", "t=100@$t=ln" & VBA.ChrW(8289) & "(ln" & VBA.ChrW(8289) & "(5/9)" & VBA.ChrW(183) & "-0,355412+e^22 )" & VBA.ChrW(183) & "4,545455@t=4,545455" & VBA.ChrW(183) & "ln" & VBA.ChrW(8289) & "(e^22-0,355412" & VBA.ChrW(183) & "ln" & VBA.ChrW(8289) & "(5/9) )") Then GoTo slut
    End If
        
    ' Defintion test
    InsertTestMath "T_2" & VBA.ChrW(8788) & "8"
    Selection.TypeParagraph
    TestSolve "T_2=log" & VBA.ChrW(8289) & "(2)/log" & VBA.ChrW(8289) & "(a) ", "a", "a=2^(1/8)"
    InsertSletDef
    If StopNow Then GoTo slut

    InsertTestMath "f(x)" & VBA.ChrW(8788) & "4x-2,5", True
    InsertTestMath "g(x)" & VBA.ChrW(8788) & "2" & VBA.ChrW(183) & "" & VBA.ChrW(12310) & "0,8" & VBA.ChrW(12311) & "^x", True
    TestBeregn "f(x)+g(x)", "=2" & VBA.ChrW(183) & "0,8^(x)+4" & VBA.ChrW(183) & "x-2,5"
    InsertSletDef
    If StopNow Then GoTo slut
    
    InsertTestMath "f(x)" & VBA.ChrW(8788) & "" & VBA.ChrW(8730) & "(r^2-x^2 )"
    Selection.TypeParagraph
    TestBeregn "S=2" & VBA.ChrW(960) & "" & VBA.ChrW(8747) & "_(-r)^r" & VBA.ChrW(9618) & "" & VBA.ChrW(12310) & "f(x)" & VBA.ChrW(183) & "" & VBA.ChrW(8730) & "(1+(f^' (x))^2 ) dx" & VBA.ChrW(12311), "=4" & VBA.ChrW(183) & "" & VBA.ChrW(960) & "" & VBA.ChrW(183) & "r^(2)"
    InsertSletDef
    If StopNow Then GoTo slut

    InsertTestMath "Definer: a=2 ;g(x)=x^2"
    Selection.TypeParagraph
    InsertTestMath "h(x)" & VBA.ChrW(8788) & "2x"
    Selection.TypeParagraph
    InsertTestMath "b" & VBA.ChrW(8788) & "3 ;c" & VBA.ChrW(8788) & "2"
    Selection.TypeParagraph
    TestBeregn "g(h(a+b))", "=100"
    InsertSletDef
    If StopNow Then GoTo slut

    InsertTestMath "Definer: f(x;y)=x^2" & VBA.ChrW(183) & "y+2y"
    If TestBeregn(VBA.ChrW(8711) & "f(1;2)", "=(" & VBA.ChrW(9632) & "(4@3))") Then GoTo slut
    InsertSletDef

    MaximaExact = 1 ' exact
    AllTrig = True
    Radians = True
    InsertTestMath "Definer: -7,5<x<7,5"
    If TestSolve("1,386" & VBA.ChrW(183) & "cos" & VBA.ChrW(8289) & "(0,63" & VBA.ChrW(183) & "x)=0", "x", "x=" & VBA.ChrW(960) & "" & VBA.ChrW(183) & "-2,380952    " & VBA.ChrW(8744) & "    x=" & VBA.ChrW(960) & "" & VBA.ChrW(183) & "-0,7936508    " & VBA.ChrW(8744) & "    x=" & VBA.ChrW(960) & "" & VBA.ChrW(183) & "0,7936508    " & VBA.ChrW(8744) & "    x=" & VBA.ChrW(960) & "" & VBA.ChrW(183) & "2,380952@$x=-(2,380952" & VBA.ChrW(183) & "" & VBA.ChrW(960) & ")    " & VBA.ChrW(8744) & "    x=-(0,7936508" & VBA.ChrW(183) & "" & VBA.ChrW(960) & ")    " & VBA.ChrW(8744) & "    x=0,7936508" & VBA.ChrW(183) & "" & VBA.ChrW(960) & "    " & VBA.ChrW(8744) & "    x=2,380952" & VBA.ChrW(183) & "" & VBA.ChrW(960)) Then GoTo slut
    InsertSletDef
    
    
    ' denne fejler, men ikke hvis den byttes om med ovenstående, så er det den som fejler ????
    Radians = True
    Selection.TypeText "Den næste fejler, men ikke hvis den byttes om med ovenstående, så er det den som fejler..."
    Selection.TypeParagraph
    InsertTestMath "Definer: -2<x<2"
    Selection.TypeParagraph
'    TestSolve "sin" & VBA.ChrW(8289) & "" & VBA.ChrW(12310) & "(x)" & VBA.ChrW(12311) & " " & VBA.ChrW(8730) & "(cos" & VBA.ChrW(8289) & "" & VBA.ChrW(12310) & "(x)" & VBA.ChrW(12311) & " )=0", "x", "x=0    " & VBA.ChrW(8744) & "    x=" & VBA.ChrW(960) & "/2"
    If TestSolve("sin" & VBA.ChrW(8289) & "" & VBA.ChrW(12310) & "(x)" & VBA.ChrW(12311) & " " & VBA.ChrW(8730) & "(cos" & VBA.ChrW(8289) & "" & VBA.ChrW(12310) & "(x)" & VBA.ChrW(12311) & " )=0", "x", "x=0    " & VBA.ChrW(8744) & "    x=-(" & VBA.ChrW(960) & "/2)    " & VBA.ChrW(8744) & "    x=" & VBA.ChrW(960) & "/2") Then GoTo slut
    
    InsertSletDef
    If StopNow Then GoTo slut
    
    Radians = True
    If TestBeregn("sin" & VBA.ChrW(8289) & "" & VBA.ChrW(12310) & "a/b" & VBA.ChrW(12311) & "+2", "=sin" & VBA.ChrW(8289) & "(a/b)+2") Then GoTo slut
    
    MaximaExact = 0
    AllTrig = False
    Radians = False

    '    TestSolve "", "x", "x="
    '    If StopNow Then GoTo slut
    
    If Not NonInterA Then '****** Interactive start *******
        ' requires user to stop and choose retry numeric
        TestBeregn "(1+1/10^12 )^(10^12 )", VBA.ChrW(8776) & "2,718523"
        If StopNow Then GoTo slut

        ' Requires user input. Equations that trigger numeric solution
        TestSolve "x^2=" & VBA.ChrW(12310) & "0,7" & VBA.ChrW(12311) & "^x", "x", "x " & VBA.ChrW(8776) & " -15,29371    " & VBA.ChrW(8744) & "    x " & VBA.ChrW(8776) & " -1,249643    " & VBA.ChrW(8744) & "    x " & VBA.ChrW(8776) & " 0,8581024", "Click maxima numeric"
        If StopNow Then GoTo slut
        TestSolve "x" & VBA.ChrW(183) & "e^2x+e^2x=0", "x", "x" & VBA.ChrW(8776) & "-1@x=-1", "Click maxima numeric"  ' Only x=-1 is a solution
        If StopNow Then GoTo slut
    
    End If '****** Interactive end *******
    '    TestBeregn VBA.ChrW(8747) & "_(-" & VBA.ChrW(8734) & ")^" & VBA.ChrW(8734) & "" & VBA.ChrW(9618) & "1/(" & VBA.ChrW(8730) & "2" & VBA.ChrW(960) & "á3)áe^(-1/2á((y-1)/3)^2 ) dy", "=1"
    
    
    ' Numrerisk test
    MaximaExact = 2
    ShowSettings "Numerisk test"

'    TestBeregn "cos^(-1)" & VBA.ChrW(8289) & "(2)", VBA.ChrW(8776) & "Ikke defineret indenfor " & VBA.ChrW(8477) ' med VBACAS kommer der fejlpopup
'    If StopNow Then GoTo slut
    TestBeregn VBA.ChrW(12310) & "cos" & VBA.ChrW(12311) & "^(-1) (0,5)", VBA.ChrW(8776) & "60" ' test af invers trig og indtastning via design-skrift-hævet skrift
    If StopNow Then GoTo slut
    TestBeregn "23456789123", VBA.ChrW(8776) & "2,345679" & VBA.ChrW(183) & "10^10@$23456790000"
    If StopNow Then GoTo slut
    TestBeregn "B=sin^(-1)" & VBA.ChrW(8289) & "(sin" & VBA.ChrW(8289) & "(40)" & VBA.ChrW(183) & "0,8605341)", VBA.ChrW(8776) & "33,58274"
    If StopNow Then GoTo slut

    MaximaExact = 0
    ShowSettings
    
    'numeric definition test
    InsertTestMath "definer: f(x)=-x^2" & VBA.ChrW(8729) & "" & VBA.ChrW(8730) & "x+2/x;x_1=0,25"
    Selection.TypeParagraph
    TestSolve "f^' (x)=-32", "x", "x=0,2512389    " & VBA.ChrW(8744) & "    x=5,464284"
    If StopNow Then GoTo slut
    TestBeregn "y=f^' (x_1 )" & VBA.ChrW(183) & "(x-x_1 )+f(x_1 )", VBA.ChrW(8776) & "-32,3125" & VBA.ChrW(183) & "x+16,04688@$" & VBA.ChrW(8776) & "16,04688-32,3125" & VBA.ChrW(183) & "x"
    InsertSletDef
    If StopNow Then GoTo slut

    If Not NonInterA Then '****** Interactive start *******
        InsertTestMath "Definer: f(x)=211,4885-10,4801" & VBA.ChrW(183) & "(e^0,0329x+e^(-0,0329x) )"
        Selection.TypeParagraph
        TestBeregn "s=" & VBA.ChrW(8747) & "_(-91,25312)^91,25312" & VBA.ChrW(9618) & "" & VBA.ChrW(8730) & "(1+(f^' (x))^2 ) dx", VBA.ChrW(8776) & "451,2554"
        InsertSletDef
        If StopNow Then GoTo slut
    End If '****** Interactive end *******
    
    ' Scientific notation test
    MaximaExact = 2
    MaximaDecOutType = 3
    ShowSettings "Videnskabelig notation test"
    TestBeregn "123", "=1,23" & VBA.ChrW(183) & "10^2"
    If StopNow Then GoTo slut
    MaximaDecOutType = 2

    ' Differential- og integralregning
    MaximaExact = 0
    ShowSettings "Differential- og integralregning"
    If TestBeregn(VBA.ChrW(8518) & "/" & VBA.ChrW(8518) & "x (" & VBA.ChrW(8747) & "x^2 dx)", "=x^(2)") Then GoTo slut
    If TestBeregn(VBA.ChrW(8706) & "^2/" & VBA.ChrW(8706) & "x" & VBA.ChrW(8706) & "y (x^2" & VBA.ChrW(183) & "y)", "=2" & VBA.ChrW(183) & "x") Then GoTo slut
    If TestBeregn(VBA.ChrW(8518) & "^2/(" & VBA.ChrW(8518) & "x^2 ) (x^2+2x)", "=2") Then GoTo slut
    If TestBeregn(VBA.ChrW(8747) & "_1^10" & VBA.ChrW(9618) & "" & VBA.ChrW(8730) & "(x^(-2)+1) dx", "=-(ln" & VBA.ChrW(8289) & "(" & VBA.ChrW(8730) & "(101)+1)-ln" & VBA.ChrW(8289) & "(" & VBA.ChrW(8730) & "(101)-1)-ln" & VBA.ChrW(8289) & "(" & VBA.ChrW(8730) & "(2)+1)+ln" & VBA.ChrW(8289) & "(" & VBA.ChrW(8730) & "(2)-1)-2" & VBA.ChrW(183) & "" & VBA.ChrW(8730) & "(101)+2^(3/2))/2" & VBA.ChrW(8776) & "9,417202@$=-((ln" & VBA.ChrW(8289) & "(" & VBA.ChrW(8730) & "(101)+1)-ln" & VBA.ChrW(8289) & "(" & VBA.ChrW(8730) & "(101)-1)-ln" & VBA.ChrW(8289) & "(" & VBA.ChrW(8730) & "(2)+1)+ln" & VBA.ChrW(8289) & "(" & VBA.ChrW(8730) & "(2)-1)-2" & VBA.ChrW(183) & "" & VBA.ChrW(8730) & "(101)+2^(3/2))/2)" & VBA.ChrW(8776) & "9,417202") Then GoTo slut
    InsertTestMath "Definer: f(x)=" & VBA.ChrW(8730) & "(3x+9)  ;g(x)=x+3"
    If TestSolve(VBA.ChrW(8747) & "_0^k" & VBA.ChrW(9618) & "" & VBA.ChrW(12310) & "g(x)-f(x) " & VBA.ChrW(12311) & " dx=1,5", "k", "k=-3    " & VBA.ChrW(8744) & "    k=7/3@$k=-3    " & VBA.ChrW(8744) & "    k=2,333333") Then GoTo slut
    ' de næste to gav før forkert resultat pga brug ldefint i integrate, så nu anvendes Nintegrate før ldefint, som åbenbart er fejlbarlig.
    If TestBeregn(VBA.ChrW(8747) & "_(-1)^1" & VBA.ChrW(9618) & "" & VBA.ChrW(8730) & "(1+((4" & VBA.ChrW(183) & "x-4" & VBA.ChrW(183) & "x^3 )" & VBA.ChrW(183) & "e^(2" & VBA.ChrW(183) & "x^2-x^4 ) )^2 ) dx", VBA.ChrW(8776) & "4,142057") Then GoTo slut
    InsertTestMath "Definer: f(x)=0,000003" & VBA.ChrW(183) & "x^4-0,01676" & VBA.ChrW(183) & "x^2+60"
    Selection.TypeParagraph
    If TestBeregn(VBA.ChrW(8747) & "_(-52)^52" & VBA.ChrW(9618) & "" & VBA.ChrW(12310) & "" & VBA.ChrW(8730) & "(1+(f^' (x))^2 ) dx" & VBA.ChrW(12311), VBA.ChrW(8776) & "115,7009") Then GoTo slut
    InsertSletDef


    ' Vektortest
    MaximaExact = 0
    ShowSettings "Vektorer"
    If TestBeregn("((" & VBA.ChrW(9608) & "(-4@3))" & VBA.ChrW(183) & "(" & VBA.ChrW(9608) & "(1@7)))/|((" & VBA.ChrW(9608) & "(-4@3)))|^2 " & VBA.ChrW(183) & "(" & VBA.ChrW(9608) & "(-4@3))", "=(" & VBA.ChrW(9632) & "(-68/25@51/25))=(" & VBA.ChrW(9632) & "(-2,72@2,04))@$=(" & VBA.ChrW(9632) & "(-(68/25)@51/25))=(" & VBA.ChrW(9632) & "(-2,72@2,04))") Then GoTo slut
    If TestBeregn("(" & VBA.ChrW(9632) & "(-400@0@320))" & VBA.ChrW(215) & "(" & VBA.ChrW(9632) & "(-120@280@0))", "=(" & VBA.ChrW(9632) & "(-89600@-38400@-112000))") Then GoTo slut 'To forskellige måde at taste vektorer på:
    If TestBeregn("(" & VBA.ChrW(9608) & "(-400@0@320))" & VBA.ChrW(215) & "(" & VBA.ChrW(9608) & "(-120@280@0))", "=(" & VBA.ChrW(9632) & "(-89600@-38400@-112000))") Then GoTo slut
    If TestSolve("(" & VBA.ChrW(9608) & "(x+y@x-y))=(" & VBA.ChrW(9608) & "(1@2))", "x", "x=3/2    " & VBA.ChrW(8743) & "    y=-(1/2)") Then GoTo slut
        
    MaximaExact = 2
    TestBeregn "rref([" & VBA.ChrW(9632) & "(-1&0&1/2&1/2@1/2&-1&0&1/2@1/2&1&-1&0@0&0&1/2&-1)])", "=[" & VBA.ChrW(9632) & "(1&0&0&-1,5@0&1&0&-1,25@0&0&1&-2@0&0&0&0)]"
    If StopNow Then GoTo slut
    MaximaExact = 0

    InsertTestMath "Definer: s" & VBA.ChrW(8407) & "(t)=(" & VBA.ChrW(9608) & "(2+t@t^2 ))"
    If StopNow Then GoTo slut
    TestBeregn "s" & VBA.ChrW(8407) & "(2)", "=(" & VBA.ChrW(9632) & "(4@4))"
    If StopNow Then GoTo slut
    TestBeregn "s" & VBA.ChrW(8407) & "^'(t)", "=(" & VBA.ChrW(9632) & "(1@2" & VBA.ChrW(183) & "t))"
    If StopNow Then GoTo slut
    
    ' Differentialligning test
    TestSolveDE "N^'=1/10500" & VBA.ChrW(183) & "N" & VBA.ChrW(183) & "(1000-N)", "N,x", "N=0    " & VBA.ChrW(8744) & "    N=1000    " & VBA.ChrW(8744) & "    N=1000/(c" & VBA.ChrW(183) & "e^(-((2" & VBA.ChrW(183) & "x)/21) )+1)"
    If StopNow Then GoTo slut
    TestSolveDE "N^'=(0,025-0,0004t)" & VBA.ChrW(183) & "N", "N,t", "N=c" & VBA.ChrW(183) & "e^(t/40-t^2/5000)"
    If StopNow Then GoTo slut
    TestSolveDE "y^'=-2x/(1+x^2 )" & VBA.ChrW(183) & "y+1/(1+x^2 )", "y,x", "y=(x+c)/(x^2+1)"
    If StopNow Then GoTo slut
    TestSolveDE "y^'+2x" & VBA.ChrW(183) & "y=x", "y,x", "y=c" & VBA.ChrW(183) & "e^(-x^2 )+1/2"
    If StopNow Then GoTo slut
    TestSolveDE "L^'=k" & VBA.ChrW(183) & "(100-L)", "L,x", "L=c" & VBA.ChrW(183) & "e^(-(k" & VBA.ChrW(183) & "x) )+100"
    If StopNow Then GoTo slut
    TestSolveDE "2y^'+y^2-5y=0", "y,x", "y=0    " & VBA.ChrW(8744) & "    y=5    " & VBA.ChrW(8744) & "    y=5/(c" & VBA.ChrW(183) & "e^(-((5" & VBA.ChrW(183) & "x)/2) )+1)"
    If StopNow Then GoTo slut
    TestSolveDE "y^'=5y" & VBA.ChrW(183) & "(y+1)", "y,x", "y=0    " & VBA.ChrW(8744) & "    y=-1    " & VBA.ChrW(8744) & "    y=-1/(c" & VBA.ChrW(183) & "e^(-(5" & VBA.ChrW(183) & "x) )+1)@$y=0    " & VBA.ChrW(8744) & "    y=-1    " & VBA.ChrW(8744) & "    y=-(1/(c" & VBA.ChrW(183) & "e^(-(5" & VBA.ChrW(183) & "x) )+1))"
    If StopNow Then GoTo slut
    TestSolveDE "y^'=b" & VBA.ChrW(183) & "y" & VBA.ChrW(183) & "(b/a-y)", "y,x", "y=0    " & VBA.ChrW(8744) & "    y=b/a    " & VBA.ChrW(8744) & "    y=b/(c" & VBA.ChrW(183) & "a" & VBA.ChrW(183) & "e^(-((b^2" & VBA.ChrW(183) & "x)/a) )+a)"
    If StopNow Then GoTo slut
    If TestSolveDE("y^'=b" & VBA.ChrW(183) & "y" & VBA.ChrW(183) & "(M-y)", "y,x", "y=0    " & VBA.ChrW(8744) & "    y=M    " & VBA.ChrW(8744) & "    y=M/(c" & VBA.ChrW(183) & "e^(-(M" & VBA.ChrW(183) & "b" & VBA.ChrW(183) & "x) )+1)") Then GoTo slut
    If TestSolveDE("(y^' )^2+x" & VBA.ChrW(183) & "y^'=0", "y,x", "y=c    " & VBA.ChrW(8744) & "    y=c-x^2/2") Then GoTo slut
    ' partikulære løsninger
    ' Denne har tidligere ikke givet en løsning. Den blev sorteret fra i ic1real, da TESTTF fandt en lille forskel i konstanterne, og så troede det var en falsk løsning. TestTF har nu fået numerisk sammenligning
    If TestSolveDE("p^'=0,015" & VBA.ChrW(183) & "p^1,2", "p=5,28;x=0", "p=-(1000000000000000/(243" & VBA.ChrW(183) & "(x-238,9747)^5 ))") Then GoTo slut
    ' Denne har tidligere givet forkert løsning, da der kommer to løsninger, men når konstanten er sat ind, passer kun den ene.
    If TestSolveDE("(x+5)" & VBA.ChrW(183) & "y^'=" & VBA.ChrW(8730) & "y", "y=1;x=-4", "y=(ln" & VBA.ChrW(8289) & "(|x+5|)+2)^2/4") Then GoTo slut

ggbtest:
    ' GeoGebra test
    CASengineTempOnly = 2
    MaximaExact = 1 ' 1=exact
    '    Selection.TypeParagraph
    '    Selection.TypeText "GeoGebra CAS Test"
    ShowSettings "GeoGebra CAS Test"
    
    TestBeregn "2+3", "=5"
    If StopNow Then GoTo slut
    ' Denne er ikke justeret
    TestBeregn "2+3^3,4/log" & VBA.ChrW(8289) & "(889) -sin" & VBA.ChrW(8289) & "(34)", "=-cos" & VBA.ChrW(8289) & "(14/45 " & VBA.ChrW(960) & ")+27" & VBA.ChrW(8730) & "(5&3)^(2)/(ln(889)/ln(10))+2@$=-cos" & VBA.ChrW(8289) & "(14/45 " & VBA.ChrW(960) & ")+27" & VBA.ChrW(8730) & "(5,00000000000000&3)^(2)/(ln(889)/ln(10))+2"
    If StopNow Then GoTo slut
    TestBeregn "1/5 2", "=2/5"
    If StopNow Then GoTo slut
    TestBeregn VBA.ChrW(8731) & "(-8)", "=-2"
    If StopNow Then GoTo slut
    
    TestBeregn "a b/c+a b+f_a (x)", "=a" & VBA.ChrW(183) & "b+a" & VBA.ChrW(183) & "b/c+f_a (x)"
    If StopNow Then GoTo slut
    TestBeregn "f_a (x)+a (b+d)/c+f_c+2+a^x (2)", "=a" & VBA.ChrW(183) & "(b+d)/c+f_a (x)+f_c+2a^(x)+2"
    If StopNow Then GoTo slut
    TestBeregn "2^2x+23/2x", "=(2^(x))^(2)+23/(2x)" ' fails if 2^2x is not interpreted as 2^(2*x)
    If StopNow Then GoTo slut
    TestBeregn "log" & VBA.ChrW(8289) & "(a)", "=ln(a)/ln(10)"
    If StopNow Then GoTo slut
    
    TestBeregn "log_2" & VBA.ChrW(8289) & "(4)", "=ln(4)/ln(2)@$2" ' reducerer ikke eksakt med ggb
    If StopNow Then GoTo slut
    TestBeregn "ln" & VBA.ChrW(8289) & "(a)", "=ln(a)"
    If StopNow Then GoTo slut
    TestBeregn "log_4" & VBA.ChrW(8289) & "a", "=ln(a)/ln(4)"
    If StopNow Then GoTo slut
    '    TestBeregn VBA.ChrW(12310) & "sin" & VBA.ChrW(8289) & "(x)-sin" & VBA.ChrW(12311) & "" & VBA.ChrW(8289) & "(x_0 )/(x+y)", "=(-sin" & VBA.ChrW(8289) & "(x_0)+sin" & VBA.ChrW(8289) & "(1/180 " & VBA.ChrW(960) & "" & VBA.ChrW(183) & "x))/(x+y)" ' Test af forkert placerede skjulte parenteser
    TestBeregn VBA.ChrW(12310) & "sin" & VBA.ChrW(8289) & "(x)-sin" & VBA.ChrW(12311) & "" & VBA.ChrW(8289) & "(x_0 )/(x+y)", "=(-sin" & VBA.ChrW(8289) & "(1/180 x_0 " & VBA.ChrW(960) & ")+sin" & VBA.ChrW(8289) & "(1/180 " & VBA.ChrW(960) & "" & VBA.ChrW(183) & "x))/(x+y)"
    If StopNow Then GoTo slut
    '    TestBeregn VBA.ChrW(12310) & "sin" & VBA.ChrW(8289) & "(x)-sin" & VBA.ChrW(12311) & "" & VBA.ChrW(8289) & "(x_0 )/(x-x_0 )", "=(sin" & VBA.ChrW(8289) & "(x_0)-sin" & VBA.ChrW(8289) & "(1/180 " & VBA.ChrW(960) & "" & VBA.ChrW(183) & "x))/(x_0-x)"
    TestBeregn VBA.ChrW(12310) & "sin" & VBA.ChrW(8289) & "(x)-sin" & VBA.ChrW(12311) & "" & VBA.ChrW(8289) & "(x_0 )/(x-x_0 )", "=(sin" & VBA.ChrW(8289) & "(1/180 x_0 " & VBA.ChrW(960) & ")-sin" & VBA.ChrW(8289) & "(1/180 " & VBA.ChrW(960) & "" & VBA.ChrW(183) & "x))/(x_0-x)"
    If StopNow Then GoTo slut
    
    ' Ligninger
    If TestSolve("x^2=9", "x", "x=-3    " & VBA.ChrW(8744) & "    x=3") Then GoTo slut

    ' Differentialligninger
    If TestSolveDE("y^'+2x" & VBA.ChrW(183) & "y=x", "y,x", "y=c_1 e^(-x^2 )+1/2") Then GoTo slut
    If TestSolveDE("y^'=b" & VBA.ChrW(183) & "y" & VBA.ChrW(183) & "(M-y)", "y,x", "y=-M e^(M" & VBA.ChrW(183) & "b" & VBA.ChrW(183) & "x)/(c_1-e^(M" & VBA.ChrW(183) & "b" & VBA.ChrW(183) & "x) )") Then GoTo slut
    
    
    'Fejler
    MaximaExact = 0
    CASengineTempOnly = 0
    If TestSolve("(" & VBA.ChrW(9608) & "(t^3-t@t))=(" & VBA.ChrW(9608) & "((-15)/64@ 1/4))", "t", "t=1/4") Then GoTo slut ' Denne giver 3 løsninger, men der er kun den ene. Det står direkte i 2. koordinaten. GeoGebra virker

    
    GoTo slut
Fejl:
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
    
    MaximaDecOutType = 2
    MaximaUnits = False
    MaximaExact = 0 ' Auto
    CASengineTempOnly = 0 ' Maxima
    
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
        
    If TestType = 1 Then ' ved beregn skrives resultatet sammen med input, så sammenligning ryger
        Oresul2 = omax.MaximaOutput
        GotoPrevEq
        omax.ReadSelection
        Oresul = TrimR(omax.Kommando, vbCr)
        p = InStrRev(Oresul, "=")
        p2 = InStrRev(Oresul, ChrW(8776))
        If p2 > p Then p = p2
        If p > 0 Then Oresul = right(Oresul, Len(Oresul) - p)
        MoveCursorToEndOfCalculation False
    Else
        GotoPrevEq
        omax.ReadSelection
        Oresul = TrimR(omax.Kommando, vbCr)
        MoveCursorToEndOfCalculation False
    End If
    Oresul = Trim(Oresul)
    If Left(Oresul, 1) = "=" Or Left(Oresul, 1) = ChrW(8776) Then Oresul = right(Oresul, Len(Oresul) - 1)
    ResultOK = False
    Arr = Split(resul, "@$")
    For i = 0 To UBound(Arr)
        res = Trim(Arr(i))
        If Left(res, 1) = "=" Or Left(res, 1) = ChrW(8776) Then res = right(res, Len(res) - 1)
        If res = Oresul Or Trim(Arr(i)) = Oresul2 Then
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
            Selection.TypeText "TestBeregn " & Trim(ConvertToVBAString(komm)) & " , " & Trim(ConvertToVBAString(Oresul))
        ElseIf TestType = 2 Then
            Selection.TypeText vbTab & "If TestSolve(" & Trim(ConvertToVBAString(komm)) & " , """ & Var & """ , " & Trim(ConvertToVBAString(Oresul)) & ") Then GoTo slut"
        ElseIf TestType = 3 Then
            Selection.TypeText vbTab & "If TestSolveDE(" & Trim(ConvertToVBAString(komm)) & " , """ & Var & """ , " & Trim(ConvertToVBAString(Oresul)) & ") Then GoTo slut"
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
    
   omax.ReadSelection
   komm = TrimR(omax.Kommando, vbCr)
   DoEvents
   If InStr(LCase(komm), "define") > 0 Or TestType = 3 Then
      TestType = 3
   ElseIf InStr(komm, "y^'=") > 0 Or TestType = 4 Then
     TestType = 4
      Var = InputBox("Enter dependent and independent variable to DEsolve for. Separate by semicolon. You can add initial condition." & vbCrLf & "Examples: y;x " & vbCrLf & "y=5;x=0", "Variable", "y;x")
      If Trim(Var) = vbNullString Then Exit Sub
      Arr = Split(Var, ";")
      If UBound(Arr) < 1 Then Exit Sub
      SolveDEpar Trim(Arr(0)), Trim(Arr(1))
   ElseIf InStr(komm, "=") > 0 Or TestType = 2 Then
      TestType = 2
      Var = InputBox("Enter variable to solve for", "Variable", "x")
      If Trim(Var) = vbNullString Then Exit Sub
      MaximaSolvePar (Var)
   Else
      TestType = 1
      beregn
   End If
'   InsertTestMath komm
   Wait 0.5
   MoveCursorToEndOfCalculation
        
   If TestType = 1 Then ' ved beregn skrives resultatet sammen med input, så sammenligning ryger
      Oresul = omax.MaximaOutput
   Else
      GotoPrevEq
      omax.ReadSelection
      Oresul = TrimR(omax.Kommando, vbCr)
'      omax.Kommando = komm & "=" & Split(omax.Kommando, "=", 1)(1)
      MoveCursorToEndOfCalculation False
   End If
        
   If TestType = 1 Then
      Selection.TypeParagraph
      Selection.TypeText "Indsæt denne kode-streng i sub 'RunTestSequence':"
      Selection.TypeParagraph
      Selection.TypeText vbTab & "If TestBeregn(" & Trim(ConvertToVBAString(komm)) & " , " & Trim(ConvertToVBAString(Oresul)) & ") Then GoTo slut"
   ElseIf TestType = 2 Then
      Selection.TypeParagraph
      Selection.TypeText "Indsæt denne kode-streng i sub 'RunTestSequence':"
      Selection.TypeParagraph
      Selection.TypeText vbTab & "If TestSolve(" & Trim(ConvertToVBAString(komm)) & " , """ & Var & """ , " & Trim(ConvertToVBAString(Oresul)) & ") Then GoTo slut"
   ElseIf TestType = 3 Then
      Selection.TypeParagraph
      Selection.TypeText "Indsæt denne kode-streng i sub 'RunTestSequence':"
      Selection.TypeParagraph
      Selection.TypeText "InsertTestMath " & Trim(ConvertToVBAString(komm))
   ElseIf TestType = 4 Then
      Selection.TypeParagraph
      Selection.TypeText "Indsæt denne kode-streng i sub 'RunTestSequence':"
      Selection.TypeParagraph
      Selection.TypeText vbTab & "If TestSolveDE(" & Trim(ConvertToVBAString(komm)) & " , """ & Var & """ , " & Trim(ConvertToVBAString(Oresul)) & ") Then GoTo slut"
   End If
   Selection.TypeParagraph

End Sub
Sub GetTestString()
    ' placer cursoren i resultat af en beregning, så giver denne sub det udtryk som performtest skal have som resultat
    Dim Oresul As String, s As String
    omax.ReadSelection
    Oresul = TrimR(omax.Kommando, vbCr)
    
   s = ConvertToVBAString(Oresul)
    
    Selection.Collapse wdCollapseEnd
    Selection.EndKey unit:=wdLine
    Selection.TypeParagraph
    Selection.TypeText (s)

End Sub
Function ConvertToVBAString(Text As String) As String
    Dim s As String, j As Integer, i As Integer
    s = ""
    For j = 1 To Len(Text)
        i = AscW(Mid(Text, j, 1))
        If i > 200 Or i = 183 Then
            s = s & """ & VBA.ChrW(" & i & ") & """
        Else
            s = s & Mid(Text, j, 1)
        End If
    Next
    If Left(s, 4) = """ & " Then
        s = right(s, Len(s) - 4)
    ElseIf Left(s, 1) <> """" Then
        s = """" & s
    End If
    If right(s, 4) = " & """ Then
        s = Left(s, Len(s) - 4)
    ElseIf right(s, 1) <> """" Then
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
        i = i + 1 ' der er nogle ligninger, hvor den bare sidder fast ved samme linje. Noget med vektorer
    Loop
End Sub
Function TestSolve(komm As String, Var As String, resul As String, Optional Instruk As String) As Boolean
    PerformTest 2, komm, resul, Var, Instruk
    If StopNow Then TestSolve = True ' betyder stop
    
    Exit Function
    
    Dim s As String
    s = TestCount & ": Solving equation" & vbCrLf & "Error count: " & ErrCount
    UFwait2.Label1.Caption = s
    UFwait2.Label_tip.Caption = komm
    If Instruk <> "" Then
        UFwait2.Label_tip.Caption = komm & vbCrLf & Instruk
    End If
    InsertTestMath komm, False
    MaximaSolvePar (Var)
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
    UFwait2.Label_progress.Caption = UFwait2.Label_progress.Caption & "*"
    TestCount = TestCount + 1
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

Sub testGetListItem()
Dim ea As New ExpressionAnalyser

ea.Text = "dette er[ad;sdfs] en test ; hej(a;b{1;2}) ;{a;d} hallo"

MsgBox ea.GetNextListItem()
MsgBox ea.GetNextListItem()
MsgBox ea.GetNextListItem()
MsgBox ea.GetNextListItem()
MsgBox ea.GetNextListItem()

End Sub
Sub InsertTestMath(s As String, Optional NewLine As Boolean = True)
' indsætter maxima output i word document
    Dim mo As Range

    Selection.Range.Font.Bold = False

    If Len(s) > 500 Then
'        If MsgBox("Output is " & Len(s) & " characters long. Click OK to insert all characters in Word. Cancel to insert the first 500", vbOKCancel, Sprog.Warning) = vbCancel Then
        s = Left(s, 500) + Sprog.A(119)
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
'        MsgBox "ingen ny linje"
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
    MsgBox Selection.Text
    For Each c In Selection.Characters
        i = AscW(c)
        s = s & c & " - " & i & vbCrLf
    Next

    MsgBox s
End Sub

Sub unicodevals2()
    Dim Text As String
    Dim i As Integer
    Dim j As Integer
    Dim s As String
    Selection.OMaths.Linearize
    Selection.OMaths(1).ConvertToNormalText
    Text = Selection.Text
    Selection.OMaths(1).ConvertToMathText
    Selection.OMaths(1).Range.Select
    Selection.OMaths.BuildUp

    For j = 1 To Len(Text)
        i = AscW(Mid(Text, j, 1))
        s = s & Mid(Text, j, 1) & " - " & i & vbCrLf
    Next
    MsgBox s

End Sub
Sub UnicodeValsToString()
' laver alle Omaths i selection om til en streng der kan indsættes i VBA-kode. Bruges primært til testmodul
' Strengene indsættes efter selection i rækkefølge. Hver på ny linje
    Dim Text As String
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
        Arr(k) = Trim(mo.Range.Text)
        mo.ConvertToMathText
        mo.Range.Select
        mo.BuildUp
    Next
    Selection.Collapse wdCollapseEnd
    Selection.EndKey unit:=wdLine

    For k = 0 To UBound(Arr)
        Text = Arr(k)
        s = ""
        For j = 1 To Len(Text)
            i = AscW(Mid(Text, j, 1))
            If i > 200 Or i = 183 Then
                s = s & """ & VBA.ChrW(" & i & ") & """
            Else
                s = s & Mid(Text, j, 1)
            End If
        Next
        If Left(s, 4) = """ & " Then
            s = right(s, Len(s) - 4)
        ElseIf Left(s, 1) <> """" Then
            s = """" & s
        End If
        If right(s, 4) = " & """ Then
            s = Left(s, Len(s) - 4)
        ElseIf right(s, 1) <> """" Then
            s = s & """"
        End If
        Selection.Collapse wdCollapseEnd
        Selection.TypeParagraph
        Selection.TypeText (s)
    Next

End Sub


