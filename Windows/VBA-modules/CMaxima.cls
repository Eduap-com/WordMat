VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMaxima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Kommando As String    ' kommando som maxima skal udf*oe*re
Public StartKommando As String
Public KommandoForklaring As String ' til skrivning af kommando
Private Kommandoer() As String    ' kommandoer som maxima skal udf*oe*re. KommandoerStreng m*aa* bruges public
Public MaximaOutput As String    ' resultat fra Maxima
Public KommentarOutput As String    ' evt. kommentar fra Maxima
Private AntalKommandoer As Integer    ' bruges af runmaxima til at se hvorn*aa*r f*ae*rdig
Public vars As String    ' variable i udtryk adskilt ad ;
Public AntalVars As Integer    ' antal variable fundet af findvariable
Private IntegrationVars As String    ' variable der bruges til integration skal ikke med i varlisten
Public DiffEqVar As String    ' variabel til differentialligninger. sikrer at subst ikke bruges i diff
Public MaximaInstalled As Boolean
Public CmdTask As Task
Public StopNow As Boolean
Public TempDefs As String    ' midlertidige definitioner kun for denne kommando
Public deffejl As Boolean    ' s*ae*ttes hvis der er fejl i definitioner
Public Definitions As Boolean    ' s*ae*ttes hvis der er definitioner i dokumentet
Public DefString As String    ' contains the output from finddefinition
Public defstringtext As String    ' contains definitions from finddefinitions before ConvertToMaxima
Private ldefname() As String    ' array med alle definitioner
Private ldefvalue() As String    ' array med alle definitioner
Public defindex As Integer
Public KillDef As String    'indeholder listseparatoret liste over sidste definitioner
Public ForgetList As String    ' indeholder liste over assumes der skal slettes
Private Minput() As String
Private moutput() As String
Public ConvertErrorText As String    ' hvis der opst*aa*r fejl i fors*oe*get p*aa* at overs*ae*tte syntaks indeholder denne fejlteksten
Public MaximaInputStreng As String
Public MaximaInputStrengSec As String    ' secondary inputstring to be run after question
Private matvekt As Boolean    ' s*ae*ttes hvis der registreres matricer. S*aa* aktiveres prik og krydsprodukt.
Public matrixstartbracket As String
Public matrixendbracket As String
Private at As AutoTextEntry
Public ConvertLnLog As Boolean    ' whethter converttomaxima converts ln(x) -> log(x) and log(x)-> log(x)/log(10)
Public prevspr As String    ' bruges til n*aa*r Maxima spr. gentagne gange om det samme
Public prevsvar As String


Private Sub Class_Initialize()
    On Error Resume Next
    ConvertLnLog = True
#If Mac Then
    MaximaInstalled = True
#Else
    Dim maximasti As String
    maximasti = Dir(Environ("ProgramFiles") & "\WordMat\Maxima*", vbDirectory)
    If maximasti = "" Then
        maximasti = Dir(Environ("ProgramFiles") & "\Maxima*", vbDirectory)
        If maximasti = "" Then
            maximasti = Dir(GetProgramFilesDir() & "\WordMat\Maxima*", vbDirectory)
        End If
    End If

    If maximasti = "" Then
        MaximaInstalled = False
        MsgBox "WordMat could not locate the Maxima installation. Try to reinstall", vbOKOnly, Sprog.Error
        Exit Sub
    Else
        MaximaInstalled = True
    End If
    OpretTempdoc
#End If
    'If at Is Nothing Then ' giver fejl hvis k*oe*res i autoexec
    '    Set at = NormalTemplate.AutoTextEntries.Add(name:="WordMat", Range:=Selection.Range)
    'End If
End Sub
Public Sub PrepareNewCommand(Optional finddef As Boolean = True)
    Dim DefS As String
    On Error Resume Next    ' der kommer fejl ved autoexec
'    If at Is Nothing Then    ' giver fejl hvis k*oe*res i autoexec s*aa* skal ogs*aa* pr*oe*ves her . 19/2  bruges vel ikke mere
'        Set at = NormalTemplate.AutoTextEntries.Add(name:="WordMat", Range:=Selection.Range)
'    End If
    StopNow = False
    deffejl = False
    DiffEqVar = ""
    Kommando = ""
    matvekt = False
    matrixstartbracket = "("
    matrixendbracket = ")"
    ReDim Kommandoer(1 To 1)
    ' kommandoer?
    MaximaOutput = ""

    KommentarOutput = ""
    vars = ""
    IntegrationVars = ""
    AntalVars = 0
    TempDefs = ""
    ReDim Kommandoer(0 To 0)
    ReDim Minput(1 To 100) As String
    ReDim moutput(1 To 100) As String
    MaximaInputStreng = ""
    MaximaInputStrengSec = ""
    ConvertErrorText = ""

    If finddef Then
        Definitions = False
        InsertKillDef
        DefS = FindDefinitions
        If Len(DefS) > 0 Then
            MaximaInputStreng = MaximaInputStreng & "[" & Left(Replace(DefS, "$", ","), Len(DefS) - 1) & "]$"
        End If
        MaximaInputStreng = MaximaInputStreng & "fejl;"
    End If

    Exit Sub

    ' skal nok bruges til mac p*aa* et tidspunkt:
    Dim fkommando As String
    MaximaInputStreng = "["
    'MaximaInputStreng = MaximaInputStreng & "display2d:false,"
    'MaximaInputStreng = MaximaInputStreng & "breakup:false$"  ' ikke dele udtryk op
    'MaximaInputStreng = MaximaInputStreng & "fpprec:5$" 'antal cifre der skal bruges i beregninger, default 16 mindre g*oe*r beregninger upr*ae*cise
    MaximaInputStreng = MaximaInputStreng & "fpprintprec:" & MaximaCifre & ","    ' antal cifre der skal vises fpprec styrer hvor mange der skal bruges i beregninger
    MaximaInputStreng = MaximaInputStreng & "ratprint:false,"    'ingen advarsel om konvertering fra decimal til br*oe*k
    MaximaInputStreng = MaximaInputStreng & "ratsimpexpons:true,"  ' simplificer ogs*aa* eksponenter,
    MaximaInputStreng = MaximaInputStreng & "algebraic:false,"  ' false er default value, simplifikation af algebraiske heltal bl.a. noget med %i. Udtryk der er relle bliver ikke reduceret s*aa* %i fjernes helt hvis denne er false
    MaximaInputStreng = MaximaInputStreng & "logexpand:super,"  ' alle log regneregler
    'MaximaInputStreng = MaximaInputStreng & "realonly:" & VBA.LCase(MaximaComplex) & "," 'reelle eller komplekse l*oe*sninger, men kun til algsys
    If MaximaComplex Then
        MaximaInputStreng = MaximaInputStreng & "domain:complex,"  '
    Else
        MaximaInputStreng = MaximaInputStreng & "realonly:true,"    'reelle eller komplekse l*oe*sninger, men kun til algsys
    End If
    If MaximaExact = 2 Then
        MaximaInputStreng = MaximaInputStreng & "numer:true,"  ' numeriske l*oe*sninger
    Else
        MaximaInputStreng = MaximaInputStreng & "numer:false,"  ' exact
    End If
    If MaximaUnits Then
        MaximaInputStreng = MaximaInputStreng & "load(unit),setunits([N,J,W]),"    'enheder har nok bugs
    Else
        '    MaximaInputStreng = MaximaInputStreng & "load(noninteractive)," 'nogle gange sp*oe*rges om fortegn p*aa* variable. giver output differentieret. er vidst ikke stabil
    End If
    If Not Radians Then
        '    MaximaInputStreng = MaximaInputStreng & "HasInvTrig(udt):=block([i,del],partswitch:true,i:0,start,del:inpart(udt,i),if del=end then return (false),if del=asin then return (true),if del=acos then return (true),if del=atan then return (true), i:i+1, go (start))," ' n*oe*dvendigt at checke for inverse trig da det tager ca. 0,5s ekstra at implicere inverse trig funktioner uden grund
        '    MaximaInputStreng = MaximaInputStreng & "ConvertToDegr(udt):=block([aliz],remfunction(sing),matchdeclare(aliz,all),tellsimp(sin(aliz),sing(180/%pi*aliz)),tellsimp(cos(aliz),cosg(180/%pi*aliz)),tellsimp(tan(aliz),tang(180/%pi*aliz)),if not(HasInvTrig(udt)) then go(slut),tellsimp(asin(aliz),%pi/180*asing(aliz)),tellsimp(acos(aliz),%pi/180*acosg(aliz)),tellsimp(atan(aliz),%pi/180*atang(aliz)),slut,udt:ev(udt,simp),clear_rules(),udt),"
    End If

    'MaximaInputStreng = MaximaInputStreng & "assume_pos_pred:true," '
    'MaximaInputStreng = MaximaInputStreng & "assume_pos:true," 'nogle gange sp*oe*rges om fortegn p*aa* variable. Nu antages pos. Men s*aa* kan den ikke l*oe*se y2/y1=(x2/x1)^a for a
    MaximaInputStreng = MaximaInputStreng & "radexpand:true,"  ' default true, styrer reduktion af r*oe*dder. hvis all er sqrt(x^2)=x hvis true sqrt(x^2)=abs(x) ' all giver problemer ved reduktion af imagin*ae*re udtryk til reelle
    MaximaInputStreng = MaximaInputStreng & "listconstvars:true,"  ' for at %i ogs*aa* kommer med blandt variable s*aa* kompleks kan findes
    MaximaInputStreng = MaximaInputStreng & "solveradcan:false"    'default false, kan l*oe*se flere ligninger men langsommere. Bl.a. bruges log nu som l*oe*sning. Men kan ikke l*oe*se f.eks. 20=3*x^2.9 s*aa*
    MaximaInputStreng = MaximaInputStreng & "]$"
    'If MaximaUnits Then
    '    MaximaInputStreng = MaximaInputStreng & "setunits([N,J])$"  '
    '    AntalKommandoer = AntalKommandoer + 1
    'End If

Fejl:
    MsgBox Sprog.ErrorGeneral & vbCrLf & " preparing maxima", vbOKOnly, Sprog.Error

End Sub
Private Sub Class_Terminate()
'lukker kommandoprompt
    On Error Resume Next
    CloseCmd
    '    tempDoc.Close (False)
End Sub
Public Sub luk()
    On Error Resume Next
    CloseCmd
    KillMaxima
    '    tempDoc.Close (False)

End Sub
Sub OpretMaximaProcess()
' men kun hvis ikke eksisterer allerede
    If MaxProc Is Nothing Then
        Set MaxProc = GetMaxProc() 'CreateObject("MaximaProcessClass")
    End If
End Sub
Public Sub ResetMaxima()
    Dim kd As String
    
#If Mac Then ' indf*oe*rt fra version 1.23
    KillDef = ""
    ForgetList = ""
    Exit Sub
#End If
    
    If Len(KillDef) > 0 Then
        KillDef = Left(KillDef, Len(KillDef) - 1)    'fjern sidste komma
        kd = kd & "kill(" & KillDef & ")"
        KillDef = ""
    End If
    If Len(ForgetList) > 0 Then
        ForgetList = Left(ForgetList, Len(ForgetList) - 1)    'fjern sidste komma
        If Len(kd) > 0 Then kd = kd & ","
        kd = kd & "forget([" & ForgetList & "])"
        ForgetList = ""
    End If
    MaxProc.Reset kd
    If MaximaUnits Then
#If Mac Then
#Else
        MaxProcUnit.Reset kd
#End If
    End If
End Sub
Public Sub RunMaxima(Optional StopTime As Integer, Optional Reset As Boolean = True, Optional ConvertWordSymbols As Boolean = True)
' starts maxima sends input and receives output

'AntalKommandoer = AntalKommandoer + 2
'If Definitions Then AntalKommandoer = AntalKommandoer + 1

    Dim i As Integer
    Dim svar As String
    Dim spr As Boolean
    
'    If CASengine > 0 Then Exit Sub  ' giver problemer med gnuplot
    
#If Mac Then
    RunMaximaFile
    Exit Sub
#Else
#End If
    
    If StopTime = 0 Then StopTime = 10000
    AntalB = AntalB + 1

    '    MaximaInputStreng = "mu:2;sigma:1;Integrate(1/(sqrt(2*%pi)*sigma)*%e^(-1/2*((y-mu)/sigma)^2),y,minf,1000);"
    '     MsgBox (MaximaInputStreng)
    '    Selection.InsertAfter MaximaInputStreng

    If DebugWM Then
        '        UserFormDebug.Label_time.Caption = UserFormDebug.Label_time.Caption & "vba tid: " & Timer - tid & vbCrLf
        '        ShowDebug
    End If

    If MaxProc Is Nothing Then
       Set MaxProc = GetMaxProc() 'CreateObject("MaximaProcessClass")
       MaxProc.StartMaximaProcess
    End If

    MaxProc.AntalCifre = MaximaCifre

    If MaximaComplex Then
        MaxProc.Complex = 1
    Else
        MaxProc.Complex = 0
    End If
    '    If MaximaUnits Then
    '        MaxProc.Units = 1
    '        MaxProc.OutUnits = omax.CodeForMaxima(Replace(OutUnits, ",", ";"))
    '    Else
    '        MaxProc.Units = 0
    '    End If
    i = 0
    Dim UfWait2 As UserFormWaitForMaxima
    Set UfWait2 = New UserFormWaitForMaxima

'    Dim UFStartup As New UserFormWaitStartup  ' slettet 18/2 2017 Hvad skulle den?
    If StopTime < -2 Then    ' ved grafer og enkelte andre nsolve
        '        MsgBox MaximaInputStreng
        MaxProc.ExecuteMaximaCommand MaximaInputStreng, 0
#If Mac Then
        MaxProc.WaitForMaximaUntil (-1 * StopTime)
#Else
        WaitForMaximaUntil (-1 * StopTime)
#End If
        '        Wait (2)
        If StopNow Then GoTo Slut
        If MaxProc.Question = 1 Then
            AnswerQuestion StopTime, False ' muligvis problematisk efter *ae*ndret til at form kan vises med answerquestion
#If Mac Then
            MaxProc.WaitForMaximaUntil (-1 * StopTime)
#Else
            WaitForMaximaUntil (-1 * StopTime)
#End If

        End If
        If MaxProc.Finished = 0 Then
            MaxProc.CloseProcess
            MaxProc.StartMaximaProcess
        End If
        FindLastMaximaOutput
        ResetMaxima
        '    MsgBox (MaxProc.LastMaximaOutput)
        Exit Sub
    ElseIf StopTime < 0 Then
        UfWait2.Show vbModeless
        DoEvents
        MaxProc.ExecuteMaximaCommand MaximaInputStreng, 1
        FindLastMaximaOutput
        ResetMaxima
        Unload UfWait2
        '        MaxProc.CloseProcess ' g*oe*r at den h*ae*nger lidt, m*aa*ske n*oe*dvendig
        ShowDebug
        Exit Sub
        GoTo Slut
    Else

        If MaxProc.Finished = 0 Then
            UfWait2.Show vbModeless
            DoEvents
            Do While MaxProc.Finished = 0 And i < 50
                If Round(i / 8) = i / 8 Then
                    UfWait2.Label_progress.Caption = UfWait2.Label_progress.Caption & "*"
                End If
                Wait (0.1)
                i = i + 1
            Loop
            UfWait2.Label_tip.Font.Size = 10
            UfWait2.Label_tip.Font.Italic = False
            UfWait2.Label_tip.Caption = GetRandomTip
        End If
        If StopNow Then
            Unload UfWait2
            MaxProc.CloseProcess
            MaxProc.StartMaximaProcess
            Exit Sub
        End If

        MaxProc.ExecuteMaximaCommand MaximaInputStreng, 0
        UfWait2.Label_progress.Caption = UfWait2.Label_progress.Caption & "*"
    End If


    '    If stoptime < 0 Then
    '        Do While MaxProc.Finished = 0
    '        Wait (0.1)
    '        If StopNow Then
    '            Unload ufwait2
    '            MaxProc.CloseProcess
    '            Exit Sub
    '        End If
    '        Loop
    '        GoTo slut
    '    End If
    
    WaitForMaximaForm StopTime

    AnswerQuestion StopTime

    ShowDebug
    '    MsgBox (MaxProc.LastMaximaOutput)

    If Reset Then
        FindLastMaximaOutput ConvertWordSymbols
        ResetMaxima
    End If
    On Error Resume Next
    Unload UfWait2
    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:
End Sub
Function ConvertMaximaUnits(Expr As String, Optional basedim As Boolean = True) As String
' Converts a maxima expression, using the maxima-image with units loaded
#If Mac Then ' noget af nedenst*aa*ende m*aa* kunne bruges, men f*oe*rst pr*oe*ves dette
    ConvertMaximaUnits = Expr
#Else
    Dim i As Integer, k As String, Arr As Variant
    
    Expr = Replace(Expr, """", "") ' svaret vil typisk v*ae*re k*oe*rt igennem num-funktion der tilf*oe*jer "" om tal
    
    If basedim Then ' all units are reverted to base SI-units
        k = Expr & ",numer;applyb1(%,unitrule),numer;dectalallNum(%),numer;"
    Else
        k = "numer:true$uforget(append(globalbaseunitlisting,globalderivedunitlisting))$"
        k = k & "setunits(usersetunits);"
        If Not OutUnits = vbNullString Then
            Arr = Split(OutUnits, ",")
            For i = 0 To UBound(Arr) ' Det skal g*oe*res p*aa* denne m*aa*de fordi man ikke kan k*oe*re setunits p*aa* en liste hvor der optr*ae*der den samme type enhed to gange. fx eV og J
                k = k & "setunits(" & omax.ConvertUnits(Arr(i)) & ")$"
            Next
        End If
'        k = k & expr & ",numer;applyb1(%,unitrule),numer;scanmap(processunits,%);dectalallNum(%),numer;"
        k = k & Expr & ",numer;ConvertUnit(%);"
        If MaximaVidNotation Then
            k = k & "scinotall(%);"
        Else
'            k = k & "dectalallNum(%),numer;" ' giver problemer for enheder med % prefix da enhederne s*aa* vises f*oe*rst.
            k = k & "%,numer;"
        End If
    End If
'    MsgBox k
    MaxProcUnit.ExecuteMaximaCommand k, 0
    WaitForMaximaUnitUntil
'    MsgBox MaxProcUnit.LastMaximaOutput
    ConvertMaximaUnits = GetMaximaUnitOutput() 'MaxProcUnit.LastMaximaOutput
#End If
End Function
Sub WaitForMaximaForm(StopTime As Integer)
    Dim i As Integer
    Dim UfWait2 As New UserFormWaitForMaxima
    
    i = 0
#If Mac Then
    MaxProc.WaitForMaximaUntil 1
#Else
    Do While MaxProc.Finished = 0 And i < 11
        Wait (0.1)
        i = i + 1
    Loop
#End If
    If MaxProc.Finished = 0 Then
        '    ufwait2.omax = omax
        On Error Resume Next
        UfWait2.Show vbModeless
        DoEvents
        On Error GoTo Fejl
        i = 0
        Do While MaxProc.Finished = 0 And i < StopTime
#If Mac Then
    MaxProc.WaitForMaximaUntil 1
            i = i + 10
#Else
            Wait (0.1)
            i = i + 1
#End If
            If Round(i / 8) = i / 8 Then
                UfWait2.Label_progress.Caption = UfWait2.Label_progress.Caption & "*"
            ElseIf Round(i / 60) = i / 60 Then
                UfWait2.Label_tip.Font.Size = 10
                UfWait2.Label_tip.Font.Italic = False
                UfWait2.Label_tip.Caption = GetRandomTip
            ElseIf Len(UfWait2.Label_progress.Caption) > 33 Then
                UfWait2.Label_progress.Caption = "*"
            End If
            '        If i = 50 Then
            '            ufwait2.Label_tip.Caption = "         Regner stadig..."
            '        ElseIf i = 150 Then
            '            ufwait2.Label_tip.Caption = "       Det tager sin tid..."
            '        ElseIf i = 250 Then
            '            ufwait2.Label_progress.Caption = "**"
            '            ufwait2.Label_tip.Caption = " Jeg arbejder stadig p*aa* det."
            '        ElseIf i = 350 Then
            '            ufwait2.Label_tip.Caption = "Tryk evt. stop nu."
            '        ElseIf stoptime > 450 Then
            '            DoEvents
            If i = StopTime - 200 Then
                UfWait2.Label_tip.Caption = "om 20s stoppes automatisk."
            ElseIf i = StopTime - 150 Then
                UfWait2.Label_tip.Caption = "om 15s stoppes automatisk."
            ElseIf i = StopTime - 100 Then
                UfWait2.Label_tip.Caption = "om 10s stoppes automatisk."
            ElseIf i = StopTime - 50 Then
                UfWait2.Label_tip.Caption = "om 5s stoppes automatisk."
            ElseIf i = StopTime Then
                UfWait2.Label_tip.Caption = "Det lykkedes ikke. Stopper."
            End If
            If StopNow Then
                '            MaxProc.ConsoleInterrupt
                Unload UfWait2
                ShowDebug
                Dim tmis As String
                tmis = MaximaInputStreng
                MaxProc.CloseProcess
                MaxProc.StartMaximaProcess
                Exit Sub
            End If
        Loop
        If i >= StopTime Then
            StopNow = True
            Unload UfWait2
            MaxProc.CloseProcess
            Exit Sub
        Else
            Unload UfWait2
        End If
    End If
    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:
End Sub
Sub ShowDebug(Optional vis As Boolean = False)
    If DebugWM Or vis Then
        UserFormDebug.TextBox_input = MaximaInputStreng
#If Mac Then
#Else
        UserFormDebug.TextBox_output = MaxProc.MaximaOutput
#End If
        UserFormDebug.TextBox_maximaoutput = MaxProc.LastMaximaOutput
        UserFormDebug.Label_time.Caption = UserFormDebug.Label_time.Caption & "Samlet tid: " & Timer - tid & vbCrLf
        UserFormDebug.Show
    End If
End Sub
Sub AnswerQuestion(StopTime As Integer, Optional ShowWaitForm As Boolean = True)
    Dim svar As String
    Dim answered As Boolean

    Do While MaxProc.Question = 1
        Dim Pos As Integer
        Dim pos2 As Integer
        Pos = InStr(MaxProc.LastMaximaOutput, "positive, negative or zero?")
        pos2 = InStr(MaxProc.LastMaximaOutput, "positive or negative?")
        If Pos > 0 Or pos2 > 0 Then
            Dim udtryk As String
            Dim UF As New UserFormAskSign
            If Pos <= 0 And pos2 > 0 Then Pos = pos2
            pos2 = InStr(MaxProc.LastMaximaOutput, "Is ")
            UF.Label_udtryk.Caption = Mid(MaxProc.LastMaximaOutput, pos2 + 3, Pos - pos2 - 4)
            If UF.Label_udtryk.Caption = prevspr Then
                svar = prevsvar
            Else
                prevspr = UF.Label_udtryk.Caption
                UF.Show
                If UF.OptionButton_positiv.Value = True Then
                    svar = "pos"
                ElseIf UF.OptionButton_negativ.Value = True Then
                    svar = "neg"
                Else
                    svar = "zero"
                End If
                Unload UF
            End If
        ElseIf InStr(MaxProc.LastMaximaOutput, " integer?") And MaximaComplex = False Then
            svar = "y"
        ElseIf InStr(MaxProc.LastMaximaOutput, "Is the matrix  1. Diagonal  2. Symmetric  3. Antisymmetric  4. General") Then
            svar = InputBox(MaxProc.LastMaximaOutput, "spr", "")    ' MaxProc.Question skal f*oe*rst til passes
        ElseIf InStr(MaxProc.LastMaximaOutput, "variables to solve for?") Then
            Dim antalvar As Integer, var As String, ea As New ExpressionAnalyser, sprtext As String
            Pos = InStr(MaxProc.LastMaximaOutput, " variables to solve for?")
            antalvar = val(Mid(MaxProc.LastMaximaOutput, Pos - 2, 2))
            sprtext = Sprog.A(416) & " " & antalvar & " " & Sprog.A(417)
            ea.Text = right(MaxProc.LastMaximaOutput, Len(MaxProc.LastMaximaOutput) - Pos - 24)
            ea.SetSquareBrackets
            ea.Pos = 0
            var = ea.GetNextBracketContent
igen:
            svar = InputBox(sprtext, "spr", var)
            If svar = "" Then
                svar = "end;%;%;%"
            Else
                If UBound(Split(svar, ",")) = antalvar - 1 Then
                    svar = "[" & svar & "];%;%;%"
                Else
                    sprtext = Sprog.A(418) & " " & antalvar & " " & Sprog.A(419)
                    GoTo igen
                End If
            End If
        Else
            svar = InputBox(MaxProc.LastMaximaOutput, "spr", "y")
        End If
        prevsvar = svar
        MaxProc.AnswerQuestion (svar)
'        MsgBox MaxProc.LastMaximaOutput
        WaitForMaximaForm StopTime
        answered = True
    Loop

    If answered And MaximaInputStrengSec <> "" Then
        MaximaInputStrengSec = "%th(2);" & MaximaInputStrengSec    ' pga. slutprut skal hoppes over.
        If right(MaximaInputStrengSec, 1) = ";" Then    ' ved answerquestion s*ae*ttes automatisk ;
            MaximaInputStrengSec = Left(MaximaInputStrengSec, Len(MaximaInputStrengSec) - 1)
        End If
        MaxProc.AnswerQuestion (MaximaInputStrengSec)
        WaitForMaximaForm StopTime
    End If

End Sub
Public Sub RunMaximaFile(Optional StopTime As Integer, Optional ConvertWordSymbols As Boolean = True)
' for mac. Starts maxima, sends command, receives output. closes maxima
#If Mac Then

    Dim s As String, i As Integer, WT As String
    Dim FN As String
    Dim UfWait2 As UserFormWaitForMaxima
    If StopTime > 0 Then
        WT = StopTime
    Else
        WT = "5"
    End If
    Set UfWait2 = New UserFormWaitForMaxima
    On Error Resume Next ' hvis en form allerede er *aa*ben giver det fejl
    UfWait2.Show vbModeless
    UfWait2.Label_tip.Font.Size = 10
    UfWait2.Label_tip.Font.Italic = False
    On Error GoTo 0
start:
    UfWait2.Label_tip.Caption = GetRandomTip
    DoEvents
  FN = DataFolder & "output.txt"
    s = ExecuteMaximaViaFile(MaximaInputStreng, WT, MaximaUnits)
    If Left(s, 5) = "Fejln" Then
        WT = InputBox("Kommandoen blev afbrudt efter " & WT & "s. Vil du pr*oe*ve igen, og hvor lang tid vil du vente. (sekunder) Computeren l*aa*ses i ventetiden.", "Afbrudt", 20)
        If Trim(WT) <> "" Then
            WT = Trim(WT)
            GoTo start
        Else
            StopNow = True
        End If
    End If
'    i = 0
'    Do Until Dir(FN) > "" Or i > 80
'        i = i + 1
'        Wait (0.1)
'    Loop
    FindLastMaximaOutput ConvertWordSymbols, s
'    GenerateOutputArray (s)
    Unload UfWait2
#End If
End Sub
Public Sub RunMaximaFileOld()
' starts maxima and sends inputfile. Then waits til Outputfile is generated.
    Dim maximasti As String
    Dim inputfil As String
    Dim outputfil As String
    Dim appnr As Integer

    AntalKommandoer = AntalKommandoer + 2
    If Definitions Then AntalKommandoer = AntalKommandoer + 1

    Dim shellcmd As String
    Dim i As Integer

    'UserFormWaitForMaxima.sh
    shellcmd = "cmd /C " & maximasti & " < " & inputfil & " > " & outputfil   '/K holder cmd *aa*ben
    'shellcmd = "cmd /C " & maximasti & " maxima.bat < " & InputFil & " > " & OutputFil   '/K holder cmd *aa*ben
    'shellcmd = "cmd /C " & maximasti & " start /high ""maxima.bat < " & InputFil & " > " & OutputFil & """"   '/K holder cmd *aa*ben
    appnr = Shell(shellcmd, vbHide)    'vbNormalFocus vbMinimizedNoFocus vbhide

    ' vent p*aa* maxima har genereret outputfil
    i = 0
    Do
        i = i + 1
        Wait (0.1)
    Loop Until Dir(outputfil) > "" Or i > 80

    FindCmdTask    ' find kommandoprompten mens der ventes alligevel

    ' vent p*aa* outputfilen er skrevet f*ae*rdig
ventmere:
    If WaitForMaxima(30) Then
        FindLastMaximaOutput
    Else
        FindLastMaximaOutput
        '    Dim result As VbMsgBoxResult
        '    result = MsgBox("Maxima tager meget lang tid om beregningen. Tryk Ok hvis du vil vente mere eller annuller for at afslutte Maxima.", vbOKCancel, "Langsom")
        '    If result = vbOK Then
        '        i = 0
        '        GoTo ventmere
        '    Else
        '        KommentarOutput = "Maxima tager meget lang tid om beregning. Maxima blev afsluttet"
        '        On Error Resume Next
        '    End If
        CloseCmd

        '    MsgBox "Maxima tog meget lang tid om beregningen. S*aa* den blev afsluttet inden den var f*ae*rdig.", vbOKOnly, "Fejl"
    End If


    ' vent p*aa* maximakommandoprompt er lukket
    'Dim maximaopen As Boolean
    'maximaopen = False
    'Do
    'i = i + 1
    'Wait (0.1)
    'For Each task1 In Tasks
    ''    MsgBox task1.Name & " - " & task1.Application & " - " & task1.Creator
    '    maximaopen = False
    '    If InStr(task1.Name, "cmd.exe") > 0 Then '  And InStr(task1.Name, "Maxima") > 0
    '        MsgBox task1.Application & " - " & task1.Creator
    '        Set CmdTask = task1
    '        maximaopen = True
    '        Exit For
    '    Else
    '    End If
    'Next
    'Loop Until Not maximaopen Or i > 80

End Sub
Public Sub RunwxMaxima()
    Dim shellcmd As String
    Dim i As Integer
    Dim result As VbMsgBoxResult
    Dim MaximaInputStrengNonInteractive As String
    Dim maximasti As String
    Dim inputfil As String
    Dim appnr As Integer

    MaximaInputStreng = Left(MaximaInputStreng, Len(MaximaInputStreng) - 7) & "print(outpudtr);"
    MaximaInputStreng = Replace(MaximaInputStreng, "display2d:false", "display2d:true", 1, 1)
    MaximaInputStrengNonInteractive = ""

    WriteInputFile (False)
    maximasti = Dir(Environ("ProgramFiles") & "\Maxima*", vbDirectory)
    maximasti = """" & Environ("ProgramFiles") & "\" & maximasti & "\wxMaxima\wxMaxima.exe"""

    shellcmd = maximasti & " /o " & inputfil

    result = MsgBox("Det er ikke muligt at l*oe*se problemet uden at stille nogle sp*oe*rgsm*aa*l." & vbCrLf & "Du bliver nu sendt til wxMaxima hvor du skal svare p*aa* et eller flere sp*oe*rgsm*aa*l. Angiv svar som y for yes eller n for no etc. tryk shift-enter efter svarene." & vbCrLf & vbCrLf & "Det resulterende udtryk kan s*aa* kopieres til Word igen.", vbOKCancel, "wxMaxima")

    If result = vbOK Then
        appnr = Shell(shellcmd, vbNormalFocus)    'vbNormalFocus vbMinimizedNoFocus vbhide
    End If

End Sub

Public Function WaitForMaxima(t As Integer) As Boolean
' vent p*aa* outputfil er genereret f*ae*rdig eller indtil der er g*aa*et t sekunder
' returnerer true hvis den er f*ae*rdig
    Dim fillen As Integer
    Dim fillen2 As Integer
    Dim i As Integer
    Dim filtext As String
    Dim NoRuns As Integer
    Dim Arr As Variant
    Dim MaximaInputStrengNonInteractive As String
    i = 0
    StopNow = False
    Do
        i = i + 1
        filtext = FilOutput
        fillen2 = fillen
        fillen = Len(filtext)
        Wait (0.1)    ' vent til beregnet
        If StopNow Then
            On Error Resume Next
            CloseCmd
            Exit Function
        End If
    Loop Until (fillen >= 297 And InStr(280, filtext, "(%i" & AntalKommandoer & ")") > 0) Or i > t * 10 Or fillen > 4000 Or InStr(279, filtext, "Incorrect syntax")  ' 301 er korteste mulige outputfil der er f*ae*rdig
    'ekstra sikkerhed fjernet: fillen = fillen2 And
    'MsgBox Len(FilOutput)

    If fillen > 4000 And NoRuns = 0 Then
        CloseCmd
        Arr = Split(vars, ";")
        Dim decl As String
        decl = "declare(["
        decl = decl & Replace(vars, ";", ",")
        decl = decl & "],integer)"

        decl = "[" & decl & ",assume_pos:true,load(noninteractive)]$"
        '    decl = decl & "[matchdeclare ([aa, bb, cc, dd], true),simp: false,tellsimpafter (if aa then bb else bb, bb),tellsimpafter (if aa then bb elseif cc then bb else bb, bb),tellsimpafter (if aa then bb elseif cc then dd else dd, if aa then bb else dd),tellsimpafter (if aa then bb elseif cc then bb else dd, if aa or cc then bb else dd),tellsimpafter ('if aa then bb else bb, bb),tellsimpafter ('if aa then bb elseif cc then bb else bb, bb),tellsimpafter ('if aa then bb elseif cc then dd else dd, 'if aa then bb else dd),tellsimpafter ('if aa then bb elseif cc then bb else dd, 'if aa or cc then bb else dd),simp: true]$"
        MaximaInputStrengNonInteractive = decl

        AntalKommandoer = AntalKommandoer - 1
        WriteInputFile (False)
        WaitTilClosed
        '    Kill outputfil
        NoRuns = NoRuns + 1
        RunMaxima
    End If

    If i > t * 10 Then
        WaitForMaxima = False
    Else
        WaitForMaxima = True
    End If

End Function
Public Sub CloseCmd()
    On Error GoTo Slut

    If Not CmdTask Is Nothing Then
        CmdTask.Close
    End If
    GoTo Slut
Fejl:
    KillMaxima
Slut:
End Sub
Sub FindCmdTask()
    On Error Resume Next
    'For Each task1 In Tasks
    '    If InStr(task1.Name, "cmd.exe") > 0 And task1.Visible = False Then '  And InStr(task1.Name, "Maxima") > 0
    ''        MsgBox task1.Application & " - " & task1.Creator
    '        Set CmdTask = task1
    '        Exit For
    '    Else
    '    End If
    'Next

End Sub
Private Sub WaitTilClosed()
    On Error Resume Next
    'dim b As Boolean
    'Dim i As Integer
    'start:
    'b = False
    'i = i + 1
    'For Each task1 In Tasks
    '    If InStr(task1.Name, "cmd.exe") > 0 Then '  And InStr(task1.Name, "Maxima") > 0
    ''        MsgBox task1.Application & " - " & task1.Creator
    '        b = True
    '        Exit For
    '    Else
    '    End If
    'Next
    'If b And i < 1000 Then GoTo start

End Sub
Public Sub WriteInputFile(Optional quit As Boolean = True)
' generates the inputfile from the variable MaximaInputStreng
    Dim filnr As Integer

    'If quit Then MaximaInputStreng = MaximaInputStreng & "quit()$"

End Sub
Sub InsertKillDef()
    If Len(KillDef) > 0 Then
        KillDef = Left(KillDef, Len(KillDef) - 1)    'fjern sidste komma
        MaximaInputStreng = MaximaInputStreng & "kill(" & KillDef & ")$"
        KillDef = ""
    End If
    If Len(ForgetList) > 0 Then
        ForgetList = Left(ForgetList, Len(ForgetList) - 1)    'fjern sidste komma
        MaximaInputStreng = MaximaInputStreng & "forget(" & ForgetList & ")$"
        ForgetList = ""
    End If

End Sub
Sub AddStandardSecondary()
' typisk skal denne k*oe*res efter maximastregen er sat op til at k*oe*re kommandoen
' denne s*oe*rger s*aa* for at output kommer som anvist i indstillinger

    If MaximaVidNotation Then     ' videnskabelig notation
        MaximaInputStrengSec = MaximaInputStrengSec & "scinotall(%),numer"    ' uden ,numer bliver der ikke reduceret i udtrykket sqrt(2) f.eks.
    ElseIf MaximaExact = 2 Or MaximaUnits Then    ' numerisk
        MaximaInputStrengSec = MaximaInputStrengSec & "dectalall(%),numer"
    Else
        If MaximaExact = 0 Then    ' kun eksakt hvis udtrykket ikke er for langt
            MaximaInputStrengSec = MaximaInputStrengSec & "if length(charlist(string(%)))>70 then ev(dectalall(%),numer) else dectalall(%)"
        Else    ' eksakt
            MaximaInputStrengSec = MaximaInputStrengSec & "if length(charlist(string(%)))>150 then ev(dectalall(%),numer) else dectalall(%)"
        End If
    End If

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec & ";"
End Sub
Public Sub ExecuteMaximaCommand(Optional MaxCmd As String, Optional TranslateToMaxima As Boolean = False)
' udf*oe*rer den kommando som er i kommando-strengen eller angives som parameter
    Dim Text As String
    If MaxCmd <> "" Then
        Kommando = MaxCmd    ' VBA.LCase m*aa*ske
    End If


    Application.ScreenUpdating = False
    If TranslateToMaxima Then
        MaximaInputStreng = MaximaInputStreng & CodeForMaxima(Kommando) & ";"
    Else
        Text = Kommando
        Text = Replace(Text, VBA.ChrW(9633), "")    ' tegn foran paranteser der g*oe*r dem usynlige. Muligvis problematisk da paranteserne ikke fjernes. g*oe*r ogs*aa* br*oe*ker sm*aa*
        Text = Replace(Text, VBA.ChrW(9508), "")    ' tegn for tom plads ved ihvertfald [ ]
        Text = Replace(Text, VBA.ChrW(11), "")    ' shift-enter
        Text = Replace(Text, vbLf, "")    ' shift-enter og enter
        Text = Replace(Text, vbCrLf, "")
        Text = Replace(Text, vbCr, "")
        Text = Replace(Text, VBA.ChrW(183), "*")    ' prik erstattes med gange
        Text = Replace(Text, VBA.ChrW(8901), "*")    ' \cdot , ved ikke hvorfor der tilsyneladende er to
        Text = Replace(Text, VBA.ChrW(8729), "*")    ' prik \cdot erstattes med gange
        Text = Replace(Text, VBA.ChrW(8226), "*")    ' tyk prik erstattes med gange
        Text = Replace(Text, VBA.ChrW(247), "/")    '
        Text = Replace(Text, VBA.ChrW(8800), "#")    ' ikkeligmed
        Text = Replace(Text, VBA.ChrW(8804), "<=")    '
        Text = Replace(Text, VBA.ChrW(8805), ">=")    '
        Text = Replace(Text, VBA.ChrW(160), "")    ' foran accent f.eks. vector pil 160 8407
        Text = Replace(Text, VBA.ChrW(8289), "") ' funktionstegn
        Text = Replace(Text, VBA.ChrW(9618), "") ' funktionstegn
        Text = Replace(Text, VBA.ChrW(12310), "(") ' skjulte parenteser
        Text = Replace(Text, VBA.ChrW(12311), ")") ' skjulte parenteser
        
        Kommando = Text
        Kommando = Replace(Kommando, VBA.ChrW(8289), "")
        MaximaInputStreng = MaximaInputStreng & Kommando & ";"
    End If
    Call RunMaxima
    If StopNow Then GoTo Slut


Slut:
End Sub

Sub MaximaSolve(variabel As String)
    Dim nr As Integer
    Dim i As Integer
    Dim gemMaxExct As Integer
    gemMaxExct = -1
    'ti = Timer

    If variabel = "" Then Exit Sub
    Dim rkommando As String, ikommando As String, pkommando As String, ckommando As String

    ikommando = Kommando    ' kommando gemmes. Denne pilles der ikke ved
start:
    KommandoForklaring = "Solve(" & Kommando & "," & variabel & ")"    ' kun til output af kommando forklaring

    ckommando = CodeForMaxima(ikommando)    ' coded kommando

    '    kommando = "CSolve(" & ckommando & "," & variabel & ")"
    If TempDefs <> "" Then
        '        MaximaInputStreng = MaximaInputStreng & "[" & CodeForMaxima(Replace(TempDefs, "=", ":")) & "]$"
        '        AddToKillDefs (TempDefs)

        MaximaInputStreng = MaximaInputStreng & AddDefinition("definer:" & TempDefs)
    End If
    If AllTrig Then
        MaximaInputStreng = MaximaInputStreng & "AllTrig:true$"
    Else
        MaximaInputStreng = MaximaInputStreng & "AllTrig:false$"
    End If

    If MaximaExact = 2 Then
        rkommando = "ev(%,numer)"    ' fors*oe*g p*aa* at l*oe*se : 6^2=5^2+3^2-2*5*3*cos(C) l*oe*s for C med numer, giver delvist eksakt
    Else
        rkommando = "ev(%)"    ' kommando hvor log og sin konverteres
    End If
    '    rkommando = "ev(%th(2))" ' kommando hvor log og sin konverteres  th2 pga slutprut n*aa*r der bruges answerquestion


    '    If Not Radians And Not MaximaExact = 2 Then ' kan ikke huske hvorfor  not num, men det er et problem n*aa*r der kommer udtryk med variabel
    If Not Radians Then
        rkommando = "ConvertToDegr(" & rkommando & ")"
    End If
    '    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
    If MaximaLogOutput = 0 Then
        If (InStr(ikommando, "log" & VBA.ChrW(8289)) > 0 Or InStr(ikommando, "10" & VBA.ChrW(12311) & "^") > 0) Then
            rkommando = "ConvertLog(" & rkommando & ")"
        End If
        '    ElseIf MaximaLogOutput = 2 And Not MaximaExact = 2 Then
    ElseIf MaximaLogOutput = 2 Then
        rkommando = "ConvertLog(" & rkommando & ")"
    End If

    ' inputstrengen bygges op**************

    ' f*oe*rst en simpel reducering af ligningen. nok ikke n*oe*dvendig. Giver problemer med sp*oe*rgsm*aa*l da kommando med spr. er sidste der udf*oe*res
    '    If MaximaExact = 2 Then
    '        MaximaInputStreng = MaximaInputStreng & "fullratsimp(" & ckommando & "),numer;"
    '    Else
    '        MaximaInputStreng = MaximaInputStreng & ckommando & ";"
    '    End If
    
    If MaximaUnits Then 'reducer f*oe*rst enheder p*aa* udtrykket inden kommandoen. Kan nogen gange give et mere kompliceret udtryk.
        If Not (InStr(ckommando, "diff") > 0 And InStr(ckommando, "subst") > 0) Then ' hvis der indg*aa*r noget f' m*aa* det ikke reduceres inden
            ckommando = Replace(ConvertMaximaUnits(ckommando), """", "")
        End If
    End If
    
'    If MaximaUnits Then ' fjernet pga maximaunitimage
'        MaximaInputStreng = MaximaInputStreng & "rhs(" & ckommando & ");"    ' kryptisk, men ellers kan komme lisp fejl f.eks. (-t+1)*m/s=0
'        MaximaInputStreng = MaximaInputStreng & "lhs(" & ckommando & ");"    ' kryptisk, men ellers kan komme lisp fejl f.eks. -(252178095*m*x-206157266*m)/(342482374*s)=0
'        MaximaInputStreng = MaximaInputStreng & "%=%th(2);"    ' kryptisk, men ellers kan komme lisp fejl f.eks. (-t+1)*m/s=0
'        MaximaInputStreng = MaximaInputStreng & "fullratsimp(%),numer;"    ' hvis lisp fejl p*aa* denne bliver % bare lig forrige
'        MaximaInputStreng = MaximaInputStreng & "applyunitrule(lhs(%)-rhs(%))=0;"
'        MaximaInputStreng = MaximaInputStreng & "CSolve(%," & variabel & ");"
    If MaximaExact = 2 Then    ' n*oe*dvendigt for at trig ligninger med sing ikke l*oe*ses eksakt
        MaximaInputStreng = MaximaInputStreng & "Solve(ev(" & ckommando & ",numer)," & variabel & ");"
    Else
        MaximaInputStreng = MaximaInputStreng & "Solve(" & ckommando & "," & variabel & ");"
    End If


    '    RunMaxima reset:=False
    '    MaximaInputStreng = ""

    MaximaInputStrengSec = MaximaInputStrengSec & "fullratsimp(" & rkommando & ");"    's*aa* log og trig konverteringer


    AddStandardSecondary

    Call RunMaxima

    'MsgBox MaximaInputStreng
    '    MaxProc.AnswerQuestion (MaximaInputStreng)
    '    FindLastMaximaOutput

    If StopNow Then
        If MaximaExact < 2 Then
            If MsgBox(Sprog.RetryNum, vbRetryCancel, "Stop") = vbRetry Then
                gemMaxExct = MaximaExact
                MaximaExact = 2
                StopNow = False
                PrepareNewCommand
                Kommando = ikommando
                GoTo start
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If


    '    MsgBox MaxProc.LastMaximaOutput
    '    ShowOutputArray
    nr = UBound(Minput)
    If nr < 2 Or nr = 1000 Then Exit Sub

    KommentarOutput = ""
    For i = 0 To nr - 1
        KommentarOutput = KommentarOutput & Minput(nr - i)
    Next
    KommentarOutput = Replace(KommentarOutput, vbCrLf, "")

    '    MaximaOutput = Trim(ConvertToWordSymbols(Moutput(nr))) ' numerisk med radcan


    If MaximaOutput = "[]" Then
        If LmSet Then
            MaximaOutput = "L=" & VBA.ChrW(8709)
        Else
            MaximaOutput = variabel & VBA.ChrW(8712) & VBA.ChrW(8709)
        End If
    ElseIf MaximaOutput = "all" Or MaximaOutput = "[" & variabel & "=k_1]" Then
        If LmSet Then
            MaximaOutput = "L=" & VBA.ChrW(8477)
        Else
            MaximaOutput = variabel & VBA.ChrW(8712) & VBA.ChrW(8477)
        End If
    ElseIf MaximaOutput = "false" Then
        MaximaOutput = ""
    Else

        If InStr(omax.KommentarOutput, "solving system of equations") > 0 Then
            ConvertOutputToSolved2 (variabel)
        ElseIf LmSet Then
            ' notation med L={}

            If MaximaOutput <> "" Then
                MaximaOutput = Replace(MaximaOutput, "[", "")
                MaximaOutput = Replace(MaximaOutput, "]", "")
                MaximaOutput = Replace(MaximaOutput, "=", "")
                MaximaOutput = Replace(MaximaOutput, variabel, "")
                MaximaOutput = "L={" & MaximaOutput & "}"
            End If

        Else    ' notation med x=...
            MaximaOutput = RemoveBrackets(MaximaOutput)
            MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8744) & "    ")    ' komma erstattes med eller
        End If

        '    MaximaOutput = Replace(MaximaOutput, "%i", "i")
    End If    ' slut p*aa* L=*oe*
    ResetMaxima

    If gemMaxExct >= 0 Then MaximaExact = gemMaxExct
Slut:
End Sub
Sub MaximaSolveInequality(variabel As String)
    Dim nr As Integer
    Dim Arr As Variant
    Dim i As Integer
    If variabel = "" Then Exit Sub

    Kommando = "SolveIneq(" & CodeForMaxima(Kommando) & "," & variabel & ")"
    '    MaximaInputStreng = MaximaInputStreng & "load(solve_rat_ineq)$"
    '    kommando = "solve_rat_ineq(" & CodeForMaxima(kommando) & ")"
    '    MaximaInputStreng = MaximaInputStreng & "load(solve_rat_ineq)$"
    '    kommando = "fourier_elim([" & CodeForMaxima(kommando) & "],[" & variabel & "])"
    '    MaximaInputStreng = MaximaInputStreng & "load(fourier_elim)$"
    MaximaInputStreng = MaximaInputStreng & "fullratsimp(" & Kommando & "),numer:false;%,numer;"

    AddStandardSecondary

    Call RunMaxima
    If StopNow Then GoTo Slut

    nr = UBound(Minput)
    If Not MaximaExact = 2 Then
        MaximaOutput = ConvertToWordSymbols(moutput(nr - 2))
    End If

    ConvertOutputToSolved2 (variabel)

    GoTo Slut




    nr = UBound(Minput)
    If Not MaximaExact = 2 Then
        If InStr(moutput(nr - 1), variabel & " = ") > 0 Then    ' hvis l*oe*st analytisk skal analytisk l*oe*sning bruges ellers numerisk
            MaximaOutput = ConvertToWordSymbols(moutput(nr - 1))
        Else
            MaximaOutput = ConvertToWordSymbols(MaximaOutput)
        End If
    End If


    If MaximaOutput = "[]" Then
        If LmSet Then
            MaximaOutput = "L=" & VBA.ChrW(8709)
        Else
            MaximaOutput = variabel & VBA.ChrW(8712) & VBA.ChrW(8709)
        End If
    ElseIf MaximaOutput = "all" Or MaximaOutput = "[" & variabel & "=k_1]" Then
        If LmSet Then
            MaximaOutput = "L=" & VBA.ChrW(8477)
        Else
            MaximaOutput = variabel & VBA.ChrW(8712) & VBA.ChrW(8477)
        End If
    ElseIf MaximaOutput = "false" Then
        MaximaOutput = ""
    Else


        If LmSet Then
            ' notation med x til [0,inf[
            MaximaOutput = Trim(MaximaOutput)
            moutput(nr - 1) = Trim(moutput(nr - 1))

            If moutput(nr - 1) <> "" Then
                MaximaOutput = Replace(MaximaOutput, "[", "")
                MaximaOutput = Replace(MaximaOutput, "]", "")
                MaximaOutput = Replace(MaximaOutput, "=", "")
                MaximaOutput = Replace(MaximaOutput, variabel, "")
                MaximaOutput = "L={" & MaximaOutput & "}"
            End If

        Else    ' notation med x=...
            MaximaOutput = RemoveBrackets(MaximaOutput)
            If Not MaximaComplex Then
                Arr = Split(MaximaOutput, ListSeparator)
                MaximaOutput = ""
                For i = 0 To UBound(Arr)
                    If InStr(Arr(i), "%i") < 1 Then     'kun reelle l*oe*sninger
                        If MaximaOutput <> "" Then MaximaOutput = MaximaOutput & "    " & VBA.ChrW(8743) & "    "
                        MaximaOutput = MaximaOutput & Arr(i)
                    End If
                Next
            Else
                MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8743) & "    ")    ' komma erstattes med og
            End If
            MaximaOutput = Replace(MaximaOutput, "or", "  " & VBA.ChrW(8744) & "  ")    ' or erstattes med v

        End If
    End If
    MaximaOutput = Replace(MaximaOutput, "%i", "i")

Slut:

End Sub
Sub MaximaSolveNumeric(ByVal variabel As String, ByVal guess As String, Optional ByVal eps As String)
    Dim Kommentar As String
    If variabel = "" Then Exit Sub
    If eps = "" Then
        eps = "10^-" & MaximaCifre
    Else
        eps = "10^-" & eps
    End If
    If InStr(Kommando, "=") < 1 Then
        Kommentar = "Der skal v*ae*re et ligmed tegn i en ligning."
        MaximaOutput = ""
        Exit Sub
    End If
    Kommando = Replace(Kommando, "=", "-(") & ")"    ' skal v*ae*re udtryk ikke ligning. Dette s*ae*ttes lig nul
    Kommando = "newton(" & CodeForMaxima(Kommando) & "," & variabel & "," & guess & "," & eps & ")"
    MaximaInputStreng = MaximaInputStreng & "[ratsimpexpons:false,load(""newton1"")]$" & Kommando & ";"
    ' af uransagelige *aa*rsager kan der komme eksakte l*oe*sninger ud

    If MaximaVidNotation Then     ' videnskabelig notation
        MaximaInputStrengSec = "scinotall(%),numer"  ' uden ,numer bliver der ikke reduceret i udtrykket sqrt(2) f.eks.
    Else
        MaximaInputStrengSec = "dectalall(%),numer"
    End If

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec & ";"

    Call RunMaxima
    If StopNow Then GoTo Slut

    If LmSet Then
        ' notation med L={}
        MaximaOutput = Trim(MaximaOutput)
        If MaximaOutput <> "" Then
            MaximaOutput = Replace(MaximaOutput, "[", "")
            MaximaOutput = Replace(MaximaOutput, "]", "")
            MaximaOutput = Replace(MaximaOutput, "=", "")
            MaximaOutput = Replace(MaximaOutput, variabel, "")
            MaximaOutput = "L={" & MaximaOutput & "}"
        End If
    Else    ' notation med x=...
        MaximaOutput = RemoveBrackets(MaximaOutput)
        If MaximaOutput <> "" Then
            MaximaOutput = variabel & VBA.ChrW(&H2248) & RemoveBrackets(MaximaOutput)
        End If
        MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8744) & "    " & variabel & VBA.ChrW(&H2248))    ' komma erstattes med eller
    End If

Slut:
End Sub
Sub MaximaFindRoot(ByVal variabel As String, ByVal xmin As String, ByVal xmax As String, Optional ByVal eps As String)
    If variabel = "" Then Exit Sub
    If eps = "" Then
        eps = "10^-" & MaximaCifre
    End If
    Kommando = Replace(Kommando, "=", "-(") & ")"    ' skal v*ae*re udtryk ikke ligning. Dette s*ae*ttes lig nul
    Kommando = "find_root(" & CodeForMaxima(Kommando) & "," & variabel & "," & xmin & "," & xmax & ")"
    MaximaInputStreng = MaximaInputStreng & "[ratsimpexpons:false]$" & Kommando & ";"

    If MaximaVidNotation Then     ' videnskabelig notation
        MaximaInputStrengSec = "scinotall(%),numer"  ' uden ,numer bliver der ikke reduceret i udtrykket sqrt(2) f.eks.
    Else
        MaximaInputStrengSec = "dectalall(%),numer"
    End If
    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec & ";"

    Call RunMaxima
    If StopNow Then GoTo Slut

    MaximaOutput = Trim(MaximaOutput)
    If InStr(MaximaOutput, "find_root") > 0 Then
        MaximaOutput = "[]"
    End If
    If LmSet Then
        ' notation med L={}
        If MaximaOutput <> "" Then
            MaximaOutput = Replace(MaximaOutput, "[", "")
            MaximaOutput = Replace(MaximaOutput, "]", "")
            MaximaOutput = Replace(MaximaOutput, "=", "")
            MaximaOutput = Replace(MaximaOutput, variabel, "")
            MaximaOutput = "L={" & MaximaOutput & "}"
        End If
    Else    ' notation med x=...
        MaximaOutput = RemoveBrackets(MaximaOutput)
        If MaximaOutput <> "" Then
            MaximaOutput = variabel & VBA.ChrW(&H2248) & RemoveBrackets(MaximaOutput)
        End If
        MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8744) & "    " & variabel & VBA.ChrW(&H2248))    ' komma erstattes med eller
    End If

Slut:
End Sub
Sub Nsolve(ByVal variabel As String, ByVal qn As Integer, ByVal qm As Integer, ByVal qtid As Integer, ByVal maxsol As Integer, Optional ByVal qnn As Integer = 30, Optional ByVal qmm As Integer = 30, Optional newt As Boolean = False)
    If variabel = "" Then Exit Sub
    '    eps = "10^-" & MaximaCifre
    If TempDefs <> "" Then
        MaximaInputStreng = MaximaInputStreng & AddDefinition("definer:" & TempDefs)
    End If

    Kommando = "nsolve(" & CodeForMaxima(Kommando) & "," & variabel & "," & qn & "," & qm & "," & qtid & "," & maxsol & "," & qnn & "," & qmm & "," & VBA.LCase(newt) & ")"
    MaximaInputStreng = MaximaInputStreng & "[ratsimpexpons:false]$" & Kommando & ";"

    If MaximaVidNotation Then     ' videnskabelig notation
        MaximaInputStrengSec = "scinotall(%),numer"  ' uden ,numer bliver der ikke reduceret i udtrykket sqrt(2) f.eks.
    Else
        MaximaInputStrengSec = "dectalall(%),numer"
    End If
    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec & ";"

    RunMaxima -150
    If StopNow Then GoTo Slut

    If LmSet Then
        ' notation med L={}
        MaximaOutput = Trim(MaximaOutput)
        If MaximaOutput <> "" Then
            MaximaOutput = Replace(MaximaOutput, "[", "")
            MaximaOutput = Replace(MaximaOutput, "]", "")
            MaximaOutput = Replace(MaximaOutput, "=", "")
            MaximaOutput = Replace(MaximaOutput, variabel, "")
            MaximaOutput = "L={" & MaximaOutput & "}"
        End If
    Else    ' notation med x=...
        MaximaOutput = RemoveBrackets(MaximaOutput)
        MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8744) & "    ")    ' komma erstattes med eller
        MaximaOutput = Replace(MaximaOutput, "=", VBA.ChrW(&H2248))    ' cirka ligmed
        '    If MaximaOutput <> "" Then
        '        MaximaOutput = variabel & VBA.ChrW(&H2248) & RemoveBrackets(MaximaOutput)
        '    End If
        '    MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8744) & "    " & variabel & VBA.ChrW(&H2248)) ' komma erstattes med eller
    End If

Slut:
End Sub

Sub SolveSystem(Variable As String)
    Dim Kommentar As String
    Dim i As Integer
    Dim rkommando As String, ikommando As String, pkommando As String, ckommando As String

    ikommando = Kommando    ' kommando gemmes. Denne pilles der ikke ved

    If Variable = "" Then Exit Sub
    If InStr(Kommandoer(0), "=") < 1 Then
        Kommentar = Sprog.A(139) & "Der skal v*ae*re et ligmed tegn i en ligning."
        MaximaOutput = ""
        Exit Sub
    End If
    If TempDefs <> "" Then
        MaximaInputStreng = MaximaInputStreng & AddDefinition("definer:" & TempDefs)
        '        MaximaInputStreng = MaximaInputStreng & "[" & CodeForMaxima(Replace(TempDefs, "=", ":")) & "]$"
        '        AddToKillDefs (TempDefs)
    End If

    rkommando = "ev(%)"    ' kommando hvor log og sin konverteres
    '    If Not Radians And Not MaximaExact = 2 Then ' kan ikke huske hvorfor  not num, men det er et problem n*aa*r der kommer udtryk med variabel
    If Not Radians Then
        rkommando = "ConvertToDegr(" & rkommando & ")"
    End If
    '    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
    If MaximaLogOutput = 0 Then
        If (InStr(ikommando, "log" & VBA.ChrW(8289)) > 0 Or InStr(ikommando, "10" & VBA.ChrW(12311) & "^") > 0) Then
            rkommando = "ConvertLog(" & rkommando & ")"
        End If
        '    ElseIf MaximaLogOutput = 2 And Not MaximaExact = 2 Then
    ElseIf MaximaLogOutput = 2 Then
        rkommando = "ConvertLog(" & rkommando & ")"
    End If


    MaximaInputStreng = MaximaInputStreng & "autonsolve:false$"
    Kommando = "[" & CodeForMaxima(Kommandoer(0))
    For i = 1 To UBound(Kommandoer)
        Kommando = Kommando & "," & CodeForMaxima(Kommandoer(i))
    Next

    ' ikke sikker p*aa* at dette er godt ved ligningssystemer
'    If MaximaUnits Then 'reducer f*oe*rst enheder p*aa* udtrykket inden kommandoen. Kan nogen gange give et mere kompliceret udtryk.
'        ckommando = Replace(ConvertMaximaUnits(ckommando), """", "")
'    End If
    
'    If MaximaUnits Then
'        MaximaInputStreng = MaximaInputStreng & Kommando & "];"
'        '        kommando = "[convert(lhs(%[1])-rhs(%[1]),[])=0"
'        Kommando = "[applyunitrule(lhs(%[1]))=applyunitrule(rhs(%[1]))"
'        For i = 1 To UBound(Kommandoer)
'            Kommando = Kommando & ",applyunitrule(lhs(%[" & i + 1 & "]))=applyunitrule(rhs(%[" & i + 1 & "]))"
'        Next
'    End If

    Kommando = Kommando & "]"
    Variable = "[" & Variable & "]"

    If MaximaExact = 2 Then    ' n*oe*dvendigt for at trig ligninger med sing ikke l*oe*ses eksakt
        Kommando = "Solve(ev(" & Kommando & ",numer)," & Variable & ")"
    Else
        Kommando = "Solve(" & Kommando & "," & Variable & ")"
    End If

    '    If Not Radians Then
    '        Kommando = "fullratsimp(ConvertToDegr(" & Kommando & "))"
    '    End If

    '    If TempDefs <> "" Then
    '    kommando = kommando & "," & TempDefs & ",numer"
    '    End If

    MaximaInputStreng = MaximaInputStreng & Kommando
    If MaximaExact = 1 Then    ' exact
        MaximaInputStreng = MaximaInputStreng & ";"
    Else
        MaximaInputStreng = MaximaInputStreng & ",autonsolve=true;"
    End If

    MaximaInputStrengSec = MaximaInputStrengSec & "fullratsimp(" & rkommando & ");"    's*aa* log og trig konverteringer

    AddStandardSecondary

    Call RunMaxima
    If StopNow Then GoTo Slut

    ConvertOutputToSolved2 (Variable)

Slut:
End Sub
Sub Eliminate(Variable As String)
    Dim Kommentar As String
    Dim i As Integer
    Dim Arr As Variant
    If Variable = "" Then Exit Sub
    If InStr(Kommandoer(0), "=") < 1 Then
        Kommentar = Sprog.A(139) '"Der skal v*ae*re et ligmed tegn i en ligning."
        MaximaOutput = ""
        Exit Sub
    End If

    If TempDefs <> "" Then
        MaximaInputStreng = MaximaInputStreng & "[" & CodeForMaxima(Replace(TempDefs, "=", ":")) & "]$"
        AddToKillDefs (TempDefs)
    End If

    Kommando = "[" & CodeForMaxima(Kommandoer(0))
    For i = 1 To UBound(Kommandoer)
        Kommando = Kommando & "," & CodeForMaxima(Kommandoer(i))
    Next

'    If MaximaUnits Then
'        MaximaInputStreng = MaximaInputStreng & Kommando & "];"
'        Kommando = "[applyunitrule(lhs(%[1])-rhs(%[1]))=0"
'        '        kommando = "[convert(lhs(%[1]),[])=convert(rhs(%[1]),[])"
'        For i = 1 To UBound(Kommandoer)
'            Kommando = Kommando & ",applyunitrule(lhs(%[" & i + 1 & "]))=applyunitrule(rhs(%[" & i + 1 & "]))"
'        Next
'    End If

    Kommando = Kommando & "]"
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"
    Variable = "[" & Variable & "]"
    Kommando = "eliminate(%," & Variable & ")"
    If Not Radians Then
        Kommando = "ratsimp(ConvertToDegr(" & Kommando & "))"
    End If
    '    If TempDefs <> "" Then
    '    kommando = kommando & "," & TempDefs & ",numer"
    '    End If

    MaximaInputStreng = MaximaInputStreng & Kommando & ";"

    AddStandardSecondary

    Call RunMaxima
    If StopNow Then GoTo Slut

    MaximaOutput = Mid(MaximaOutput, 2, Len(MaximaOutput) - 2)
    Arr = Split(MaximaOutput, ListSeparator)
    If Not (InStr(Arr(0), "=") > 0) Then
        MaximaOutput = Arr(0) & "=0"
    Else
        MaximaOutput = Arr(0)
    End If
    For i = 1 To UBound(Arr)
        If Not (InStr(Arr(i), "=") > 0) Then
            MaximaOutput = MaximaOutput & "    " & VBA.ChrW(8743) & "    " & Arr(i) & "=0"
        Else
            MaximaOutput = MaximaOutput & "    " & VBA.ChrW(8743) & "    " & Arr(i)
        End If
    Next

Slut:
End Sub
Sub SolveSystemNumeric(ByVal Variable As String, ByVal guess As String, Optional ByVal eps As String)
    Dim Kommentar As String
    Dim i As Integer
    If Variable = "" Then Exit Sub
    If InStr(Kommandoer(0), "=") < 1 Then
        Kommentar = Sprog.A(139) '"Der skal v*ae*re et ligmed tegn i en ligning."
        MaximaOutput = ""
        Exit Sub
    End If
    Dim Sep As String
    Sep = ListSeparator

    Kommando = "[" & Replace(CodeForMaxima(Kommandoer(0)), "=", "-(") & ")"
    For i = 1 To UBound(Kommandoer)
        Kommando = Kommando & "," & Replace(CodeForMaxima(Kommandoer(i)), "=", "-(") & ")"
    Next
    Kommando = Kommando & "]"
    Variable = "[" & Variable & "]"
    guess = "[" & guess & "]"
    Kommando = "mnewton(" & Kommando & "," & Variable & "," & guess & ")"

    MaximaInputStreng = MaximaInputStreng & "[load(""mnewton""),ratsimpexpons:false,newtonmaximeter:5000000,newtonepsilon:10^-" & MaximaCifre & "]$" & Kommando & ";"

    If MaximaVidNotation Then     ' videnskabelig notation
        MaximaInputStrengSec = "scinotall(%),numer;"
    Else
        MaximaInputStrengSec = "dectalall(%),numer;"
    End If

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec


    Call RunMaxima
    If StopNow Then GoTo Slut

    '    MaximaOutput = Mid(MaximaOutput, 2, Len(MaximaOutput) - 2)

    ConvertOutputToSolved2 (Variable)

Slut:
End Sub
Sub beregn()
    Dim inkom As String, pkommando As String, StartKommando2 As String, StartKommando3 As String
    Dim posca As Integer
    Dim posligmed As Integer
    Dim possumtegn As Integer
    Dim gemMaxExct As Integer
    gemMaxExct = -1
    Kommando = Replace(Kommando, vbLf, "")
    Kommando = Replace(Kommando, vbCrLf, "")
    Kommando = Replace(Kommando, vbCr, "")
    '    kommando = Replace(kommando, VBA.ChrW(12310), "") ' specielle usynlige paranteser fjernes, giver problemer at fjerne her
    '    kommando = Replace(kommando, VBA.ChrW(12311), "") ' specielle usynlige paranteser fjernes
    Kommando = Trim(Kommando)    ' fjern mellemrum duer ikke da 1/5 2 bliver til 1/52

    If right(Kommando, 1) = "=" Then    ' fjern ligmed hvis det er yderst til h*oe*jre
        Kommando = Left(Kommando, Len(Kommando) - 1)
    End If

    Do    ' g*aa* tilbage til n*ae*rmeste ligmed
        posligmed = InStr(Kommando, "=")
        possumtegn = InStr(Kommando, VBA.ChrW(8721))
        '    posprodtegn = InStr(Kommando, VBA.ChrW(8719))
        If possumtegn = 0 Then possumtegn = InStr(Kommando, VBA.ChrW(8719))    ' produkttegn
        If possumtegn > 0 And possumtegn < posligmed Then    ' hvis sumtegn er der =tegn som del deraf
            posligmed = 0
        End If
        posca = InStr(Kommando, VBA.ChrW(8776))
        If posca > posligmed Then posligmed = posca
        If posligmed > 0 Then
            Kommando = right(Kommando, Len(Kommando) - posligmed)
        End If
    Loop While posligmed > 0

start:
    StartKommando = Kommando    ' kommando f*oe*r der bliver pillet ved den
    StartKommando3 = Kommando
    
    pkommando = Kommando    ' kommando f*oe*r der bliver pillet ved den
    pkommando = Replace(pkommando, VBA.ChrW(12310), "")    ' specielle usynlige paranteser fjernes, giver problemer at fjerne her
    pkommando = Replace(pkommando, VBA.ChrW(12311), "")    ' specielle usynlige paranteser fjernes
    '    pkommando = Replace(pkommando, " ", "") ' ikke altid optimalt men nok det bedste. nej ikke efter underforst*aa*ede gangetegn

'    StartKommando2 = pkommando ' til at sammenligne om numerisk resultat er lig input
'    StartKommando2 = Replace(StartKommando2, "sin", "sing")
'    StartKommando2 = Replace(StartKommando2, VBA.ChrW(8289), "")
'    StartKommando2 = CodeForMaxima(StartKommando2)
'    StartKommando2 = Replace(StartKommando2, "log10", "lgog")


' *ae*ndret pga maximaimage
'    If MaximaUnits Then
'        '        kommando = CodeForMaxima(kommando) & ";applyunitrule(%)"
'        Kommando = "applyunitrule(" & CodeForMaxima(Kommando) & ")"
'    Else
        Kommando = CodeForMaxima(Kommando)
'    End If
    inkom = "%"    ' gem kommando

    If MaximaLogOutput = 0 Then
        If (InStr(pkommando, "log" & VBA.ChrW(8289)) > 0 Or InStr(pkommando, "10" & VBA.ChrW(12311) & "^") > 0) Then
            inkom = "ConvertLog(" & inkom & ")"
        End If
    ElseIf MaximaLogOutput = 2 Then
        inkom = "ConvertLog(" & inkom & ")"
    End If

    If Not Radians Then
        inkom = "ConvertToDegr(" & inkom & ")"
    End If
    '    If Not MaximaExact = 2 Then ' kan give m*ae*rkelig resultat #1,INF med ,numer   72*e^((0.619/0.22*(e^22-e^0.22*t )) )
    ''        inkom = "fullratsimp(" & inkom & ")"
    '        inkom = "simplify(" & inkom & ")"
    '    End If

    If MaximaComplex And PolarOutput Then
        Kommando = "ConvertToPolarAngleNotation(" & Kommando
        If Radians Then
            Kommando = Kommando & ",true)"
        Else
            Kommando = Kommando & ",false)"
        End If
    End If

    If MaximaVidNotation Then      ' videnskabelig notation og enheder
        MaximaInputStreng = MaximaInputStreng & Kommando & ";"
        MaximaInputStrengSec = "ev(" & inkom & ",numer);scinotall(%);ev(is(equal(ev(scanmap(strtonum,%th(1)),nouns),%th(2))),numer);"    ' ,numer fjernet fra scinot
        MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec
    ElseIf MaximaExact = 2 Or MaximaUnits Then    ' enheder overrider, overrider auto,exact,num
        '        MaximaInputStreng = MaximaInputStreng & kommando & ",numer;" & inkom & ";%,numer;" ' m*aa*ske bruge dectalall for at f*aa* mere pr*ae*cis antal bet cifre
        If MaximaBigFloat Then
            MaximaInputStreng = MaximaInputStreng & "fpprec:50$convertbfloat(" & Kommando & ");"
        Else
            MaximaInputStreng = MaximaInputStreng & Kommando & ",numer;"
        End If
'        MaximaInputStrengSec = "simplifynum(" & inkom & ");dectalallNum(%),numer;ev(is(equal(ev(scanmap(strtonum,%th(1)),nouns)," & StartKommando2 & ")),numer);"    ' dectal er n*oe*dvendig for enheder og korrekt antal bet cif. . fors*oe*g med ,numer
'        MaximaInputStrengSec = "simplifynum(" & inkom & ");dectalallNum(%),numer;ev(is(equal(ev(scanmap(strtonum,%th(1)),nouns),%th(2))),numer);"    ' dectal er n*oe*dvendig for enheder og korrekt antal bet cif. . fors*oe*g med ,numer
        MaximaInputStrengSec = "simplifynum(" & inkom & ");"    ' dectal er n*oe*dvendig for enheder og korrekt antal bet cif. . fors*oe*g med ,numer
#If Mac Then
        If MaximaUnits Then
'            MaximaInputStrengSec = "applyb1(%,unitrule),numer;" 'g*oe*r det ikke bedre
            If InStr(ConvertUnits(OutUnits), "%") > 0 Then ' hvis % i outunits giver det muligvis enhed der placeres f*oe*r tal
                MaximaInputStrengSec = MaximaInputStrengSec & "%,numer;" ' dectalall placerer enheder med %m prefix foran tal, s*aa* hellere forkert antal betcif?
            Else
                MaximaInputStrengSec = MaximaInputStrengSec & "dectalallNum(%),numer;"
            End If
        Else
            MaximaInputStrengSec = MaximaInputStrengSec & "dectalallNum(%),numer;"
'numer:true$uforget(append(globalbaseunitlisting,globalderivedunitlisting))$setunits(%%ms);2/3*s,numer;ConvertUnit(%);%,numer;
        End If
#Else
        MaximaInputStrengSec = MaximaInputStrengSec & "dectalallNum(%),numer;"
#End If
        MaximaInputStrengSec = MaximaInputStrengSec & "ev(is(equal(ev(scanmap(strtonum,%th(1)),nouns),%th(2))),numer);"     ' dectal er n*oe*dvendig for enheder og korrekt antal bet cif. . fors*oe*g med ,numer
        MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec
    ElseIf MaximaExact = 0 Then    ' auto
        '        MaximaInputStreng = MaximaInputStreng & inkom & ",numer:false;" & inkom & ",numer;is(%=%th(2));"
        MaximaInputStreng = MaximaInputStreng & Kommando & ",numer:false;dectalall(simplifynum(" & inkom & "));"
        If MaximaBigFloat Then
            MaximaInputStreng = MaximaInputStreng & "convertbfloat(" & Kommando & ");"
        Else
            MaximaInputStreng = MaximaInputStreng & Kommando & ",numer;"
        End If
        MaximaInputStreng = MaximaInputStreng & "dectalallNum(simplifynum(" & inkom & ")),numer;IsResultExact(%th(4),%th(3),%th(1));"
        MaximaInputStrengSec = inkom & ";simplifynum(%),numer;dectalall(" & inkom & "),numer;is(equal(%th(2),%th(3)))"    ' anderledes slutning da eller kan komme spr igen
    ElseIf MaximaExact = 1 Then    ' exact
        If TempDefs <> "" Then
            inkom = inkom & "," & TempDefs
        End If
        MaximaInputStreng = MaximaInputStreng & Kommando & ";"
        MaximaInputStrengSec = "simplify(" & inkom & ");"
        MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec
    End If

    Call RunMaxima
    '
    If StopNow Then
        If MaximaExact < 2 Then
            If MsgBox(Sprog.RetryNum, vbRetryCancel, "Stop") = vbRetry Then
                gemMaxExct = MaximaExact
                MaximaExact = 2
                StopNow = False
                PrepareNewCommand
                Kommando = StartKommando
                GoTo start
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If

    Dim nr As Integer, i As Integer
    Dim eql As String
    nr = UBound(Minput)

    If nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = moutput(1)
    ElseIf MaximaExact = 0 And Not MaximaVidNotation And Not MaximaUnits Then
        Dim exactres As String
        Dim numres As String
        If nr > 1 Then moutput(nr - 1) = Replace(moutput(nr - 1), vbLf, "")
        If nr > 2 Then moutput(nr - 2) = Replace(moutput(nr - 2), vbLf, "")
'        numres = Replace(moutput(nr - 1), """", "") ' for upr*ae*cist.
'        exactres = Replace(moutput(nr - 3), """", "")
        numres = ConvertToWordSymbols(moutput(nr - 1))
        If UBound(moutput) > 3 Then exactres = ConvertToWordSymbols(moutput(nr - 3))
        If exactres = "" And numres = "" Then exactres = MaximaOutput    ' hvis der har v*ae*ret sp*oe*rgsm*aa*l
        eql = MaximaOutput
        If eql = "true" Or StartKommando3 = numres Then
             eql = "="
        Else
             eql = VBA.ChrW(&H2248)
        End If
        MaximaOutput = ""
        pkommando = ConvertToWordSymbols(Replace(pkommando, DecSeparator, ".")) ' pkommando kan ikke sammenlignes 100% med output f*oe*r den ogs*aa* er k*oe*rt igennem
        If exactres <> "" Then
            '    MsgBox "pkommando:" & pkommando & vbCrLf & " exactres:" & ConvertToWordSymbols(exactres)
            If exactres = numres Or pkommando = exactres Or StartKommando3 = exactres Then ' hvis numerisk og eksakt ens eller eksakt lig input
                MaximaOutput = MaximaOutput & eql & numres
            Else
                MaximaOutput = MaximaOutput & "=" & exactres & eql & numres
            End If
            
            KommentarOutput = ""
            For i = 0 To nr - 1
                KommentarOutput = KommentarOutput & Minput(nr - i)
            Next
            KommentarOutput = Replace(KommentarOutput, vbCrLf, "")
'            KommentarOutput = Minput(nr - 4) & Minput(nr - 3) & Minput(nr - 2) & Minput(nr - 1) & Minput(nr)
        Else
            MaximaOutput = ""
        End If
    Else    ' vid not, num, exact
        Dim eqsg As String
        eql = MaximaOutput
        If MaximaExact = 1 Then
            eqsg = "="
        ElseIf eql = "true" Then
            eqsg = "="
        Else
            eqsg = VBA.ChrW(&H2248)
        End If
        If moutput(nr - 1) <> "" Then
            MaximaOutput = eqsg & ConvertToWordSymbols(moutput(nr - 1))
            KommentarOutput = ""
            For i = 0 To nr - 1
                KommentarOutput = KommentarOutput & Minput(nr - i)
            Next
            KommentarOutput = Replace(KommentarOutput, vbCrLf, "")
'            KommentarOutput = Minput(nr - 4) & Minput(nr - 3) & Minput(nr - 2) & Minput(nr - 1) & Minput(nr)
        Else
            MaximaOutput = ""
        End If
        'GoTo slut
    End If
    
    
    'If InStr(KommentarOutput, "rat: replaced") Then
    '    KommentarOutput = ""
    'End If

    'If InStr(KommentarOutput, "Division by 0") Then
    '    KommentarOutput = "Division med 0 er ikke tilladt!!"
    '    MaximaOutput = ""
    'End If
    'If InStr(KommentarOutput, "Lisp error") Then
    '    If InStr(MaximaOutput, "[") Then ' typisk kommer der en liste hvis der er fejl []
    '        KommentarOutput = "Der skete en fejl i afviklingen af Maxima."
    '        MaximaOutput = ""
    '    End If
    'End If
    'If MaximaOutput = "" Then
    '    KommentarOutput = "Der kom f*oe*lgende fejlmelding fra Maxima:" & vbCrLf & vbCrLf & KommentarOutput
    'End If
    
    If gemMaxExct >= 0 Then MaximaExact = gemMaxExct
Slut:
End Sub
Sub PolySolve(ByVal variabel As String)
' l*oe*ser polynomium numerisk vha. speciel metode der giver alle l*oe*sninger
    Kommando = "algsys([" & CodeForMaxima(Kommando) & "],[" & variabel & "])"
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"

    If MaximaVidNotation Then     ' videnskabelig notation
        MaximaInputStrengSec = "scinotall(%),numer;"
    Else
        MaximaInputStrengSec = "dectalall(%),numer;"
    End If

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec


    Call RunMaxima
    If StopNow Then GoTo Slut
    ConvertOutputToSolved

Slut:
End Sub
Sub ConvertOutputToSolved()
' konverterer output fra [1,3] til x=1 v x=3 eller L={1,3}
' kun beregnet til numerisk l*oe*sning
    Dim Variable As String
    Dim Arr As Variant
    Dim i As Integer
    Dim Sep As String
    If LmSet Then
        ' notation med L={}
        MaximaOutput = Trim(MaximaOutput)
        If MaximaOutput <> "" Then
            MaximaOutput = Mid(MaximaOutput, 2, Len(MaximaOutput) - 2)
            MaximaOutput = Replace(MaximaOutput, "[", "(")
            MaximaOutput = Replace(MaximaOutput, "]", ")")
            MaximaOutput = Replace(MaximaOutput, "=", "")
            Arr = Split(Mid(Variable, 2, Len(Variable) - 2), ",")
            For i = 0 To UBound(Arr)
                MaximaOutput = Replace(MaximaOutput, Arr(i), "")
            Next
            MaximaOutput = "L={" & MaximaOutput & "}"
        End If
    Else    ' notation med x=...

        Dim ea As New ExpressionAnalyser
        Dim o As String
        ea.Text = MaximaOutput
        MaximaOutput = ""
        ea.Text = ea.GetNextBracketContent
        ea.Pos = 1
        o = ea.GetNextBracketContent
        If o = "" Then
            MaximaOutput = ""
            Exit Sub
        End If
        MaximaOutput = o
        MaximaOutput = Replace(MaximaOutput, Sep, "    " & VBA.ChrW(8743) & "    ")
        o = ea.GetNextBracketContent
        If o = "" Then Exit Sub    ' kun en l*oe*sning

        MaximaOutput = "(" & MaximaOutput & ")"
        o = "(" & Replace(o, Sep, "    " & VBA.ChrW(8743) & "    ") & ")"
        MaximaOutput = MaximaOutput & "   " & VBA.ChrW(8744) & "   " & o

        Do
            o = ea.GetNextBracketContent
            If o = "" Then Exit Do
            o = "(" & Replace(o, Sep, "    " & VBA.ChrW(8743) & "    ") & ")"
            MaximaOutput = MaximaOutput & "   " & VBA.ChrW(8744) & "   " & o
        Loop
    End If

End Sub
Sub ConvertOutputToSolvedGGB()
    ' til geogebraoutput ved ligningssystemer
    Dim p As Integer, s As String, i As Integer, j As Integer
    Dim ea As New ExpressionAnalyser
    Dim Arr() As String, arr2() As String
    
    If MaximaOutput = "?" Or MaximaOutput = "" Then
        Exit Sub
    End If
    ea.SetNormalBrackets
    ea.Text = MaximaOutput
'    p = InStr(ea.text, "(_")
    p = 1
    s = ea.GetNextBracketContent(p + 1)
    Arr = Split(s, "@")
    MaximaOutput = ""
    
    For i = 0 To UBound(Arr)
        arr2 = Split(Arr(i), "&")
        If UBound(arr2) > 0 Then
            If UBound(Arr) > 0 Then MaximaOutput = MaximaOutput & "("
            For j = 0 To UBound(arr2)
                MaximaOutput = MaximaOutput & arr2(j) & "    " & VBA.ChrW(8743) & "    "
            Next
            MaximaOutput = Left(MaximaOutput, Len(MaximaOutput) - 9)
            If UBound(Arr) > 0 Then MaximaOutput = MaximaOutput & ")"
        End If
        MaximaOutput = MaximaOutput & "    " & VBA.ChrW(8744) & "    "
    Next
    If Len(MaximaOutput) > 9 Then MaximaOutput = Left(MaximaOutput, Len(MaximaOutput) - 9)
    
End Sub
Sub ConvertOutputToSolved2(Optional Variable As String)
' konverterer output fra [[x=1,y=3],[x=3,y=5]] til (x=1 og x=3) v ... eller L={(1,3),(3,5)}
' kun beregnet til symbolsk l*oe*sning og mnewton og SolveIneq
    Dim Arr As Variant
    Dim i As Integer



    If MaximaOutput = "[]" Then
        If LmSet Then
            MaximaOutput = "L=" & VBA.ChrW(8709)
        Else
            MaximaOutput = Variable & VBA.ChrW(8712) & VBA.ChrW(8709)
        End If
    ElseIf MaximaOutput = "all" Or MaximaOutput = "[[all]]" Or MaximaOutput = "[" & Variable & "=k_1]" Then
        If LmSet Then
            MaximaOutput = "L=" & VBA.ChrW(8477)
        Else
            MaximaOutput = Variable & VBA.ChrW(8712) & VBA.ChrW(8477)
        End If
    ElseIf MaximaOutput = "false" Then
        MaximaOutput = ""
    Else

        If LmSet Then
            ' notation med L={}
            MaximaOutput = Trim(MaximaOutput)
            If MaximaOutput <> "" Then
                MaximaOutput = Mid(MaximaOutput, 2, Len(MaximaOutput) - 2)
                MaximaOutput = Replace(MaximaOutput, "[", "(")
                MaximaOutput = Replace(MaximaOutput, "]", ")")
                MaximaOutput = Replace(MaximaOutput, "=", "")
                Arr = Split(Mid(Variable, 2, Len(Variable) - 2), ",")
                For i = 0 To UBound(Arr)
                    MaximaOutput = Replace(MaximaOutput, Arr(i), "")
                Next
                MaximaOutput = "L={" & MaximaOutput & "}"
            End If

        Else    ' notation med x=...

            Dim ea As New ExpressionAnalyser
            Dim o As String

            ea.Text = MaximaOutput
            MaximaOutput = ""
            ea.Text = ea.GetNextBracketContent
            ea.Pos = 1
            o = ea.GetNextBracketContent
            If o = "" Then
                MaximaOutput = ea.Text
                Exit Sub
            End If
            MaximaOutput = o
            MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8743) & "    ")

            o = ea.GetNextBracketContent
            If o = "" Then Exit Sub    ' kun en l*oe*sning

            If InStr(MaximaOutput, VBA.ChrW(8743)) > 0 Then    ' kun parentes hvis og tegn
                MaximaOutput = "(" & MaximaOutput & ")"
            End If

            o = Replace(o, ListSeparator, "    " & VBA.ChrW(8743) & "    ")
            If InStr(o, VBA.ChrW(8743)) > 0 Then    ' kun parentes hvis og tegn
                o = "(" & o & ")"
            End If

            MaximaOutput = MaximaOutput & "   " & VBA.ChrW(8744) & "   " & o

            Do
                o = ea.GetNextBracketContent
                If o = "" Then Exit Do
                o = Replace(o, ListSeparator, "    " & VBA.ChrW(8743) & "    ")
                If InStr(o, VBA.ChrW(8743)) > 0 Then    ' kun parentes hvis og tegn
                    o = "(" & o & ")"
                End If
                MaximaOutput = MaximaOutput & "   " & VBA.ChrW(8744) & "   " & o
            Loop

        End If
    End If

End Sub
Sub Omskriv(vidnot As Boolean, autosimp As Boolean, faktoriser As Boolean, udvid As Boolean, rationaliser As Boolean, trigsimp As Boolean)
    Dim pkommando As String
    pkommando = Kommando

    Kommando = CodeForMaxima(Kommando)

    If TempDefs <> "" Then
        MaximaInputStreng = MaximaInputStreng & AddDefinition("definer:" & TempDefs)
    End If

    If MaximaUnits Then 'reducer f*oe*rst enheder p*aa* udtrykket inden kommandoen. Kan nogen gange give et mere kompliceret udtryk.
        Kommando = Replace(ConvertMaximaUnits(Kommando), """", "")
    End If
'    If MaximaUnits Then
'        Kommando = "applyunitrule(" & Kommando & ")"
'    End If

    If autosimp Then
        Kommando = "simplify(" & Kommando & ")"
    End If
    If faktoriser Then
        Kommando = "factor(" & Kommando & ")"
    End If
    If udvid Then
        Kommando = "expand(" & Kommando & ")"
    End If
    If rationaliser Then
        Kommando = "ev(ratsimp(" & Kommando & "),keepfloat=false)"
    End If
    If trigsimp Then
        Kommando = "trigreduce(trigsimp(" & Kommando & "))"
    End If

    If Not Radians Then
        '        kommando = "simplify(ConvertToDegr(" & kommando & "))"
        Kommando = "ConvertToDegr(" & Kommando & ")"
    End If

    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
        If (InStr(pkommando, "log" & VBA.ChrW(8289)) > 0 Or InStr(pkommando, "10" & VBA.ChrW(12311) & "^") > 0) Then
            Kommando = "ConvertLog(" & Kommando & ")"
        End If
    ElseIf MaximaLogOutput = 2 And Not MaximaExact = 2 Then
        Kommando = "ConvertLog(" & Kommando & ")"
    End If

    If MaximaExact = 2 Then
        Kommando = Kommando & ",numer"
    End If
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"


    If vidnot Then     ' videnskabelig notation
        MaximaInputStrengSec = "scinotall(%),numer;"
    ElseIf MaximaExact = 2 Then
        MaximaInputStrengSec = "dectalall(%),numer;"
    Else
        MaximaInputStrengSec = ""
    End If

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec


    Call RunMaxima
    If StopNow Then GoTo Slut

Slut:
End Sub
Sub Reduce()
    Dim pkommando As String
    pkommando = Kommando

    Kommando = CodeForMaxima(Kommando)
    If MaximaUnits Then 'reducer f*oe*rst enheder p*aa* udtrykket inden kommandoen. Kan nogen gange give et mere kompliceret udtryk.
        Kommando = Replace(ConvertMaximaUnits(Kommando), """", "")
    End If
'    If MaximaUnits Then
'        Kommando = "applyunitrule(" & Kommando & ")"
'    End If
    Kommando = "simplify(" & Kommando & ")"
    If Not Radians Then
        Kommando = "simplify(ConvertToDegr(" & Kommando & "))"
    End If
    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
        If (InStr(pkommando, "log" & VBA.ChrW(8289)) > 0 Or InStr(pkommando, "10" & VBA.ChrW(12311) & "^") > 0) Then
            Kommando = "ConvertLog(" & Kommando & ")"
        End If
    ElseIf MaximaLogOutput = 2 And Not MaximaExact = 2 Then
        Kommando = "ConvertLog(" & Kommando & ")"
    End If

    If MaximaExact = 2 Then
        Kommando = Kommando & ",numer"
    End If
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"


    If MaximaVidNotation Then     ' videnskabelig notation
        MaximaInputStrengSec = "scinotall(%),numer;"
    ElseIf MaximaExact = 2 Then
        MaximaInputStrengSec = "dectalall(%),numer;"
    Else
        MaximaInputStrengSec = ""
    End If

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec


    Call RunMaxima
    If StopNow Then GoTo Slut

Slut:
End Sub
Sub CompareTest()
    Dim pkommando As String
    pkommando = Kommando
    Kommando = CodeForMaxima(Kommando)
'    If MaximaUnits Then
'        Kommando = "applyunitrule(" & Kommando & ")"
'    End If
    Kommando = "TestTF(" & Kommando & ")"
    If Not Radians Then
        Kommando = "ConvertToDegr(" & Kommando & ")"    ' Det er m*aa*ske problematisk at ratsimp ikke kan kaldes da der expandes igen. converttodegree s*aa* virker ikke?
    End If
    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
        If (InStr(pkommando, "log" & VBA.ChrW(8289)) > 0 Or InStr(pkommando, "10" & VBA.ChrW(12311) & "^") > 0) Then
            Kommando = "ConvertLog(" & Kommando & ")"
        End If
    ElseIf MaximaLogOutput = 2 And Not MaximaExact = 2 Then
        Kommando = "ConvertLog(" & Kommando & ")"
    End If
    If MaximaExact = 2 Then
        Kommando = Kommando & ",numer"
    End If
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"

    If MaximaVidNotation Then     ' videnskabelig notation
        MaximaInputStrengSec = "scinotall(%),numer;"
    ElseIf MaximaExact = 2 Then
        MaximaInputStrengSec = "dectalall(%),numer;"
    Else
        MaximaInputStrengSec = ""
    End If

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec

    Call RunMaxima
    If StopNow Then GoTo Slut


Slut:
End Sub
Sub Factor()
    Dim pkommando As String
    pkommando = Kommando
    Kommando = CodeForMaxima(Kommando)
'    If MaximaUnits Then
'        Kommando = "applyunitrule(" & Kommando & ")"
'    End If
    Kommando = "factor(" & Kommando & ")"
    If Not Radians Then
        Kommando = "ConvertToDegr(" & Kommando & ")"    ' Det er m*aa*ske problematisk at ratsimp ikke kan kaldes da der expandes igen. converttodegree s*aa* virker ikke?
    End If
    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
        If (InStr(pkommando, "log" & VBA.ChrW(8289)) > 0 Or InStr(pkommando, "10" & VBA.ChrW(12311) & "^") > 0) Then
            Kommando = "ConvertLog(" & Kommando & ")"
        End If
    ElseIf MaximaLogOutput = 2 And Not MaximaExact = 2 Then
        Kommando = "ConvertLog(" & Kommando & ")"
    End If
    If MaximaExact = 2 Then
        Kommando = Kommando & ",numer"
    End If
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"

    If MaximaVidNotation Then     ' videnskabelig notation
        MaximaInputStrengSec = "scinotall(%),numer;"
    ElseIf MaximaExact = 2 Then
        MaximaInputStrengSec = "dectalall(%),numer;"
    Else
        MaximaInputStrengSec = ""
    End If

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec

    Call RunMaxima
    If StopNow Then GoTo Slut


Slut:
End Sub
Sub Expand()
    Dim pkommando As String
    pkommando = Kommando

    Kommando = CodeForMaxima(Kommando)
'    If MaximaUnits Then
'        Kommando = "applyunitrule(" & Kommando & ")"
'    End If
    Kommando = "expand(" & Kommando & ")"
    If Not Radians And Not MaximaExact = 2 Then
        Kommando = "ConvertToDegr(" & Kommando & ")"
    End If
    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
        If (InStr(pkommando, "log" & VBA.ChrW(8289)) > 0 Or InStr(pkommando, "10" & VBA.ChrW(12311) & "^") > 0) Then
            Kommando = "ConvertLog(" & Kommando & ")"
        End If
    ElseIf MaximaLogOutput = 2 And Not MaximaExact = 2 Then
        Kommando = "ConvertLog(" & Kommando & ")"
    End If

    If MaximaExact = 2 Then
        Kommando = Kommando & ",numer"
    End If
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"

    If MaximaVidNotation Then     ' videnskabelig notation
        MaximaInputStrengSec = "scinotall(%),numer;"
    ElseIf MaximaExact = 2 Then
        MaximaInputStrengSec = "dectalall(%),numer;"
    Else
        MaximaInputStrengSec = ""
    End If

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec

    Call RunMaxima
    If StopNow Then GoTo Slut

    '    MaximaOutput = Replace(MaximaOutput, "%i", "i")

Slut:
End Sub
Sub Differentier(variabel As String)
    Dim pkommando As String
    pkommando = Kommando

    Kommando = CodeForMaxima(Kommando)
'    If MaximaUnits Then
'        Kommando = "applyunitrule(" & Kommando & ")"
'    End If

    Kommando = "diff(" & Kommando & "," & variabel & ",1)"
    If Not Radians Then
        Kommando = "ratsimp(ConvertToDegr(" & Kommando & "))"
    End If
    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
        If (InStr(pkommando, "log" & VBA.ChrW(8289)) > 0 Or InStr(pkommando, "10" & VBA.ChrW(12311) & "^") > 0) Then
            Kommando = "ConvertLog(" & Kommando & ")"
        End If
    ElseIf MaximaLogOutput = 2 And Not MaximaExact = 2 Then
        Kommando = "ConvertLog(" & Kommando & ")"
    End If

    Kommando = "dectalallNum(" & Kommando & ")"
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"

    AddStandardSecondary

    Call RunMaxima
Slut:
End Sub
Sub UserRegression(fkt As String, var As String, xlist As String, ylist As String, prec As Integer, VarGuess As String)
    Dim pkommando As String
    pkommando = fkt

    Kommando = "GeneralRegr(" & CodeForMaxima(fkt) & "," & var & "," & xlist & "," & ylist & ",1e-" & prec & "," & VarGuess & ")"
    If Not Radians Then
        Kommando = "ConvertToDegr(" & Kommando & ")"
    End If
    If MaximaLogOutput = 0 Then
        If (InStr(pkommando, "log") > 0 Or InStr(pkommando, "10^") > 0) Then
            Kommando = "ConvertLog(" & Kommando & ")"
        End If
    ElseIf MaximaLogOutput = 2 Then
        Kommando = "ConvertLog(" & Kommando & ")"
    End If

    Kommando = Kommando & ",numer"
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"

    If MaximaVidNotation Then     ' videnskabelig notation
        MaximaInputStrengSec = "scinotall(%),numer;"
    ElseIf MaximaExact = 2 Then
        MaximaInputStrengSec = "dectalall(%),numer;"
    Else
        MaximaInputStrengSec = ""
    End If


    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec

    Call RunMaxima
    If StopNow Then GoTo Slut

    '    MaximaOutput = Replace(MaximaOutput, "%i", "i")

Slut:

End Sub
Sub Integrer(variabel As String)
    Dim pkommando As String
    pkommando = Kommando

    Kommando = CodeForMaxima(Kommando)
'    If MaximaUnits Then
'        Kommando = "applyunitrule(" & Kommando & ")"
'    End If
    Kommando = "integrate(" & Kommando & "," & variabel & ")"
    If Not Radians Then
        Kommando = "ratsimp(ConvertToDegr(" & Kommando & "))"
    End If
    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
        If (InStr(pkommando, "log" & VBA.ChrW(8289)) > 0 Or InStr(pkommando, "10" & VBA.ChrW(12311) & "^") > 0) Then
            Kommando = "ConvertLog(" & Kommando & ")"
        End If
    ElseIf MaximaLogOutput = 2 And Not MaximaExact = 2 Then
        Kommando = "ConvertLog(" & Kommando & ")"
    End If

    If MaximaExact = 2 Then
        Kommando = Kommando & ",numer"
    End If
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"

    AddStandardSecondary

    Call RunMaxima

Slut:
End Sub
Sub SolveDENumeric(variabel As String, xmin As String, xmax As String, xstep As String, yvarlist As String, initguesslist As String, DElist As String)
    
    variabel = CodeForMaxima(variabel)
    xmin = CodeForMaxima(xmin)
    xmax = CodeForMaxima(xmax)
    xstep = CodeForMaxima(xstep)
    yvarlist = CodeForMaxima(yvarlist)
    initguesslist = CodeForMaxima(initguesslist)
    DElist = CodeForMaxima(DElist)
    
    If TempDefs <> "" Then
        MaximaInputStreng = MaximaInputStreng & AddDefinition("definer:" & TempDefs)
    End If
    
    MaximaInputStreng = MaximaInputStreng & "qDElist:rk(" & DElist & "," & yvarlist & "," & initguesslist & ",[" & variabel & "," & xmin & "," & xmax & "," & xstep & "])$qDElist"

    MaximaInputStreng = MaximaInputStreng & ";"
    
    AddStandardSecondary

    Call RunMaxima
    If StopNow Then GoTo Slut
    
Slut:
End Sub
Sub SolveDE(ByVal funktion As String, ByVal variabel As String, Optional startx As String, Optional starty As String, Optional startdiffy, Optional bcx, Optional bcy)
    Dim pkommando As String
    Dim ea As New ExpressionAnalyser
    Dim gemfkt As String
    ea.SetNormalBrackets
    DiffEqVar = variabel

    pkommando = Kommando

    Kommando = CodeForMaxima(Kommando)
    funktion = CodeForMaxima(funktion)
    gemfkt = funktion
    variabel = CodeForMaxima(variabel)
    startx = CodeForMaxima(startx)
    starty = CodeForMaxima(starty)
    startdiffy = CodeForMaxima(startdiffy)
    bcx = CodeForMaxima(bcx)
    bcy = CodeForMaxima(bcy)

    If TempDefs <> "" Then
        MaximaInputStreng = MaximaInputStreng & AddDefinition("definer:" & TempDefs)
    End If

'    If MaximaUnits Then
'        Kommando = "applyunitrule(" & Kommando & ")"
'    End If

    '    kommando = "ev(ode2(" & Replace(kommando, "diff(", "'diff(") & "," & funktion & "," & variabel & "),logabs=false)" ' logabs=false fjernet da logistisk ikke l*oe*ses korrekt s*aa*
    If InStr(funktion, "(") > 0 Then    ' hvis funktion
        Kommando = Replace(Kommando, funktion, "qvar")
        ea.Text = Kommando
        ea.ReplaceVar Split(funktion, "(")(0), "qvar"
        Kommando = ea.Text
        '        Kommando = Replace(Kommando, Split(funktion, "(")(0), "qvar") ' hvis b*aa*de N(t) og N indg*aa*r
        funktion = "qvar"
    End If

    If bcy <> "" And bcx <> "" Then
        Kommando = "desolvereal(" & Replace(Kommando, "diff(", "'diff(") & "," & variabel & "=" & startx & "," & funktion & "=" & starty & "," & variabel & "=" & bcx & "," & funktion & "=" & bcy & ")"
    ElseIf startdiffy <> "" Then
        Kommando = "desolvereal(" & Replace(Kommando, "diff(", "'diff(") & "," & variabel & "=" & startx & "," & funktion & "=" & starty & ",'diff(" & funktion & "," & variabel & ")=" & startdiffy & ")"
    ElseIf starty <> "" Then
        Kommando = "desolvereal(" & Replace(Kommando, "diff(", "'diff(") & "," & variabel & "=" & startx & "," & funktion & "=" & starty & ")"
    Else
        Kommando = "desolvereal(" & Replace(Kommando, "diff(", "'diff(") & "," & variabel & "," & funktion & ")"
    End If

    '    Kommando = "ode2(" & Replace(Kommando, "diff(", "'diff(") & "," & funktion & "," & variabel & ")"
    '    If bcy <> "" And bcx <> "" Then
    '        Kommando = "bc2(" & Kommando & "," & variabel & "=" & startx & "," & funktion & "=" & starty & "," & variabel & "=" & bcx & "," & funktion & "=" & bcy & ")"
    '    ElseIf startdiffy <> "" Then
    '        Kommando = "ic2(" & Kommando & "," & variabel & "=" & startx & "," & funktion & "=" & starty & ",'diff(" & funktion & "," & variabel & ")=" & startdiffy & ")"
    '    ElseIf starty <> "" Then
    '        Kommando = "ic1(" & Kommando & "," & variabel & "=" & startx & "," & funktion & "=" & starty & ")"
    '    End If


    If Not Radians Then
        Kommando = "simplify(ConvertToDegr(" & Kommando & "))"
    End If
    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
        If (InStr(pkommando, "log" & VBA.ChrW(8289)) > 0 Or InStr(pkommando, "10" & VBA.ChrW(12311) & "^") > 0) Then
            Kommando = "ConvertLog(" & Kommando & ")"
        End If
    ElseIf MaximaLogOutput = 2 And Not MaximaExact = 2 Then
        Kommando = "ConvertLog(" & Kommando & ")"
    End If

    If MaximaExact = 2 Then
        Kommando = Kommando & ",numer"
    End If
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"
    '    MaximaInputStreng = MaximaInputStreng & "if IsSolved(%," & funktion & ") then % else CSolve(%," & funktion & ");"

    AddStandardSecondary

    If MaximaUnits Then
        MaximaUnits = False
        Call RunMaxima
        MaximaUnits = True
    Else
        Call RunMaxima
    End If
    
    If StopNow Then GoTo Slut

    If funktion = "qvar" Then
        MaximaOutput = Replace(MaximaOutput, "qvar", gemfkt)
    End If

    If Len(MaximaOutput) > 1 Then
        MaximaOutput = RemoveBrackets(MaximaOutput)
        MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8744) & "    ")    ' komma erstattes med eller
    End If

Slut:
End Sub
Sub Plot2D(forskrifter As String, punkter As String, var As String, xmin As String, xmax As String, ymin As String, ymax As String, xaksetitel As String, yaksetitel As String, Optional SaveFile As Boolean = False)
' forskrifter adskilt ad komma
'    text = "load(draw)$draw2d("
    Dim Text As String
    Dim Arr As Variant
    Dim i As Integer
    Dim fkt As String
    colindex = 0
    Text = "draw2d("
    Text = Text & "font=""Arial"",font_size=8,"
    If SaveFile Then
        Text = Text & "terminal='gif,file_name=""" & Replace(Environ("TEMP"), "\", "/") & "/WordMatGraf"",pic_height=460,pic_width=660,"
    End If

    Text = Text & "nticks=100,axis_bottom = false,axis_top = false,axis_left = false,axis_right = false,grid=true,xtics_axis = true,ytics_axis = true"

    Text = Text & ",xaxis_width = 2,xaxis_color = black,xaxis_type  = solid,xaxis=true,yaxis_width = 2,yaxis_color = black,yaxis_type=solid,yaxis=true"
    '    text = text & "explicit(" & forskrifter & ",x," & xmin & "," & xmax & ",y," & ymin & "," & ymax & ")"
    If Not MaximaComplex Then
'            fkt = "'RealOnly(" & fkt & ")"
        Text = Text & ",draw_realpart = false"
    End If
    Arr = Split(forskrifter, ";")
    For i = 0 To UBound(Arr)
        fkt = omax.CodeForMaxima(Arr(i))
        Text = Text & ",key=""" & ConvertToAscii(Arr(i)) & """,color=" & GetNextColor() & ",explicit(" & fkt & "," & var & "," & xmin & "," & xmax & ")"
    Next
    If punkter <> "" Then
        Text = Text & ",key=""punkter"",point_type=filled_circle,point_size=1,points(" & punkter & ")"
    End If

    '    If titel <> "" Then
    '        text = text & ",title=" & titel
    '    End If
    If xaksetitel <> "" Then
        Text = Text & ",xlabel=" & xaksetitel
    End If
    If yaksetitel <> "" Then
        Text = Text & ",ylabel=" & yaksetitel
    End If

    Text = Text & ");"
    MaximaInputStreng = MaximaInputStreng & Text
    Call RunMaxima(-30)

End Sub
Sub Draw2D(graphobj As String, df As String, xaksetitel As String, yaksetitel As String, grid As Boolean, SaveFile As Boolean, ByVal resolution As Integer)
    Dim Text As String
    Dim waittime As Integer
    If df = "" Then
#If Mac Then
'        text = "gnuplot_command:""/Applications/WordMat/gnuplot/gnuplot""$set_plot_option([gnuplot_term, aqua])$draw2d("
    Text = "gnuplot_command:""/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/MaximaWM/maxima/bin/gnuplot""$set_plot_option([gnuplot_term, aqua])$draw2d("
'    text = "gnuplot_command:""/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/gnuplot/gnuplot""$set_plot_option([gnuplot_term, aqua])$draw2d("
'        waittime = -20
#Else
        Text = "draw2d("
#End If
        waittime = -400
    Else
#If Mac Then
    Text = "gnuplot_command:""/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/MaximaWM/maxima/bin/gnuplot""$set_plot_option([gnuplot_term, aqua])$load(drawdf)$drawdf(" & df & ","
'text = "gnuplot_command:""/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/gnuplot/gnuplot""$set_plot_option([gnuplot_term, aqua])$load(drawdf)$drawdf(" & df & ","
'        waittime = -25
#Else
        Text = "load(drawdf)$drawdf(" & df & ","
#End If
        waittime = -600
    End If


    If SaveFile Then
#If Mac Then
        Text = Text & "terminal=pdf"
        Text = Text & ",file_name=""" & Replace(GetTempDir(), "\", "/") & "WordMatGraf""," & MacDrawDims(500, 350) & ","
#Else
        Text = Text & "terminal='gif"
        Text = Text & ",file_name=""" & Replace(GetTempDir(), "\", "/") & "WordMatGraf"",dimensions=[" & resolution * 660 & "," & resolution * 460 & "],"
#End If
'        text = text & "terminal='gif,file_name=""" & Replace(Environ("TEMP"), "\", "/") & "/WordMatGraf"",dimensions=[" & resolution * 660 & "," & resolution * 460 & "],"
        '        text = text & "terminal='gif,file_name=""" & Replace(Environ("TEMP"), "\", "/") & "/WordMatGraf"",dimensions=[1980,1380],font_size=20,"
        '        text = text & "terminal='gif,file_name=""" & Replace(Environ("TEMP"), "\", "/") & "/WordMatGraf"",dimensions=[660,460],font_size=20,"
    End If
    Text = Text & "axis_bottom = false,axis_top = false,axis_left = false,axis_right = false"
    Text = Text & ",xaxis_width = 2,xaxis_color = black,xaxis_type = solid,xaxis=true,yaxis_width = 2,yaxis_color = black,yaxis_type=solid,yaxis=true,"

    Text = Text & graphobj

    Text = Text & ",font_size=" & resolution * 12
    '    text = text & ",font_size=60"

    If grid Then
        If resolution = 1 Then
            Text = Text & ",grid=true"
        Else
            '    text = text & ",user_preamble = ""set arrow from graph 0.95, first 0 to graph 1, first 0""" ' pil p*aa* x-akse ikke p*ae*n
            '            text = text & ",user_preamble = ""set grid lt -1 """ ' sorte linjer solid
            Text = Text & ",user_preamble = ""set grid lt 0 lc -1 lw 1"""    ' sorte dottede linjer
            '           text = text & ",user_preamble = ""set grid lt 0 lw " & resolution & """" ' linewidth>1 vises ikke korrekt p*aa* vandrette linjer
            ' lw 2 er linewidth 2, lt er linetype, lc er linecolor
        End If
    End If
    If xaksetitel <> "" Then
        Text = Text & ",xlabel=""" & xaksetitel & """"
    End If
    If yaksetitel <> "" Then
        Text = Text & ",ylabel=""" & yaksetitel & """"
    End If

    Text = Text & "),numer;" ' numer indf*oe*rt fordi der var et udtryk med numerisk integrale, hvor den bare ventede uden. Der er ingen grund til eksakt udtryk n*aa*r der skal plottes.
    MaximaInputStreng = MaximaInputStreng & Text
    
#If Mac Then
    If MaximaUnits Then
        MaximaUnits = False
        Call RunMaxima(waittime)
        MaximaUnits = True
    Else
        Call RunMaxima(waittime)
    End If
#Else
        Call RunMaxima(waittime)
#End If

End Sub
Sub PlotDF(lign As String, IndepVar As String, DepVar As String, xmin As String, xmax As String, ymin As String, ymax As String, xtrajec As String, ytrajec As String, parametre As String)
'plot af kurveintegraler

    Dim Text As String

    Text = "plotdf("

    Text = Text & lign

    If xmin <> "" And xmax <> "" Then
        Text = Text & ",[" & IndepVar & "," & DepVar & "],[" & IndepVar & "," & xmin & "," & xmax & "]"
    End If
    If ymin <> "" And ymax <> "" Then
        Text = Text & ",[" & DepVar & "," & ymin & "," & ymax & "]"
    End If
    If Len(xtrajec) > 0 And Len(ytrajec) > 0 Then
        Text = Text & ",[trajectory_at," & xtrajec & "," & ytrajec & "]"
    End If
    If Len(parametre) > 0 Then
        Text = Text & "," & parametre
    End If

    Text = Text & ");"
    MaximaInputStreng = MaximaInputStreng & Text
#If Mac Then
    If MaximaUnits Then
        MaximaUnits = False
        Call RunMaxima(-500)
        MaximaUnits = True
    Else
        Call RunMaxima(-500)
    End If
#Else
        Call RunMaxima(-500)
#End If
End Sub
Sub Plot2Dold(forskrifter As String, xmin As String, xmax As String, ymin As String, ymax As String, xaksetitel As String, yaksetitel As String)
' forskrifter adskilt ad komma
    Dim Text As String
    Text = "plot2d(["
    Text = Text & forskrifter
    Text = Text & "]"

    Text = Text & ",[x," & xmin & "," & xmax & "]"
    If xaksetitel <> "" Then
        Text = Text & ",[xlabel," & xaksetitel & "]"
    End If
    If yaksetitel <> "" Then
        Text = Text & ",[ylabel," & yaksetitel & "]"
    End If
    '    text = text & ",[box,false]"
    If InStr(forskrifter, "discrete") > 0 Then
        Text = Text & ",[style,points,lines]"
    Else
        Text = Text & ",[style,lines]"
    End If

    Text = Text & ",[point_type,diamond],[axes,true],[gnuplot_out_file, false],[plot_format,xmaxima]"
    Text = Text & ");"
    MaximaInputStreng = Text
    Call RunMaxima

End Sub
Sub Plot3D(forskrifter As String, xvar As String, xmin As String, xmax As String, yvar As String, ymin As String, ymax As String, xaksetitel As String, yaksetitel As String, kunlinjer As Boolean, gridlines As Integer)
' forskrifter adskilt ad komma
    Dim Text As String
    Dim tid As Integer
    tid = 20
    tid = tid + Len(forskrifter) / 2
    Text = "plot3d(["
    Text = Text & forskrifter

    Text = Text & ",[" & xvar & "," & xmin & "," & xmax & "]"
    Text = Text & ",[" & yvar & "," & ymin & "," & ymax & "]"

    Text = Text & "]"

    If xaksetitel <> "" Then
        Text = Text & ",[xlabel," & xaksetitel & "]"
    End If
    If yaksetitel <> "" Then
        Text = Text & ",[ylabel," & yaksetitel & "]"
    End If
    '    text = text & ",[box,false]"
    '    If InStr(forskrifter, "discrete") > 0 Then
    '        text = text & ",[style,points,lines]"
    '    Else
    '        text = text & ",[style,lines]"
    '    End If
    If kunlinjer = True Then
        Text = Text & ",[palette,false]"
    End If
    If gridlines <> 30 Then
        Text = Text & ",[grid," & gridlines & "," & gridlines & "]"
    End If

    '    text = text & ",[point_type,diamond],[axes,true],[gnuplot_out_file, false],[plot_format,xmaxima]"
    Text = Text & ");"
    MaximaInputStreng = MaximaInputStreng & Text
    Call RunMaxima(-1 * tid)
#If Mac Then
    If MaximaUnits Then
        MaximaUnits = False
        Call RunMaxima(-1 * tid)
        MaximaUnits = True
    Else
        Call RunMaxima(-1 * tid)
    End If
#Else
        Call RunMaxima(-1 * tid)
#End If
End Sub
Sub Draw3D(graphobj As String, Optional antalobj As Integer = 1)
'    text = "load(draw)$draw3d("
    Dim Text As String
    Dim tid As Integer

    tid = 20 + antalobj * 10

#If Mac Then
        Text = "gnuplot_command:""/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/gnuplot/gnuplot""$set_plot_option([gnuplot_term, aqua])$draw3d("
        Text = Text & "terminal=pdf"
        Text = Text & ",file_name=""" & Replace(GetTempDir(), "\", "/") & "WordMatGraf""," & MacDrawDims(500, 350) & ","
'        waittime = -20
#Else
    Text = "draw3d("
#End If
    Text = Text & graphobj

    Text = Text & ");"
    MaximaInputStreng = MaximaInputStreng & Text

#If Mac Then
    If MaximaUnits Then
        MaximaUnits = False
        Call RunMaxima(-1 * tid)
        MaximaUnits = True
    Else
        Call RunMaxima(-1 * tid)
    End If
#Else
        Call RunMaxima(-1 * tid)
#End If
    If InStr(MaxProc.LastMaximaOutput, "no surface within these ranges") Then
        ' hvis bare een ligning ikke eksisterer indenfor det valgte vindue er der fejl
        MsgBox Sprog.A(420), vbOKOnly, Sprog.Error
    End If

End Sub
Sub Wait(pausetime As Variant)
    Dim start
    start = Timer    ' Set start time.
    Do While Timer < start + pausetime
        DoEvents    ' Yield to other processes.
    Loop

End Sub

Function CodeForMaxima(ByVal Text As String, Optional CasType As Integer = 0) As String
' Castype=0 is maxima, CasType=1/2 is GeoGebra (geogebra friendly. still needs cenvertToGeogebraSyntax)
    Dim ea As New ExpressionAnalyser
    Dim ea2 As New ExpressionAnalyser
    Dim Pos As Integer
    Dim posb As Integer
    Dim rod As Integer
    Dim pos2 As Integer
    Dim pos3 As Integer
    Dim pos4 As Integer
    Dim pos5 As Integer
    Dim posog As Integer
    Dim Arr As Variant, arr2() As String, arr3() As String
    Dim var As String
    Dim startn As String
    Dim slutn As String
    Dim Expr As String
    Dim sluts As String
    Dim app As String
    Dim f As String
    Dim udtryk As String
    Dim varl2 As String
    Dim i As Integer, j As Integer
    Dim c As String
    Dim posstart As Integer
    Dim posslut As Integer
    Dim pos2d As Integer
    Dim poss As Integer
    Dim posn As Integer
    Dim poso As Integer
    Dim nedre As String
    Dim upper As String
    Dim stpos As Integer
    Dim slpos As Integer
    Dim matrix As String
    Dim mtext As String
    Dim pm As String
    Dim s As String    ' universal string
    Dim ib As Boolean
    Dim subs As Integer, sups As Integer, ends As Integer
    Dim antalm As Integer, varl As Integer, di As Integer, antalmS As String

'    On Error GoTo fejl
    ea.SetNormalBrackets
    ea2.StartBracket = VBA.ChrW(12310)
    ea2.EndBracket = VBA.ChrW(12311)

    Text = Trim(Text)

    '    For j = 1 To Len(text)
    '    i = AscW(Mid(text, j, 1))
    '    s = s & Mid(text, j, 1) & " - " & i & vbCrLf
    '    Next
    '    MsgBox s


    Text = Replace(Text, VBA.ChrW(9633), "")    ' tegn foran paranteser der g*oe*r dem usynlige. Muligvis problematisk da paranteserne ikke fjernes. g*oe*r ogs*aa* br*oe*ker sm*aa*
    Text = Replace(Text, VBA.ChrW(9508), "")    ' tegn for tom plads ved ihvertfald [ ]
    Text = Replace(Text, VBA.ChrW(11), "")    ' shift-enter
    Text = Replace(Text, vbLf, "")    ' shift-enter og enter
    Text = Replace(Text, vbCrLf, "")
    Text = Replace(Text, vbCr, "")
    Text = Replace(Text, """", "")    ' apostrof fjernes
    Text = Replace(Text, "+-", "-")    ' plus-minus giver minus
    Text = Replace(Text, VBA.ChrW(8212), "+")    'dobbelt minustegn giver plus
    Text = Replace(Text, ChrW(8711) & VBA.ChrW(183), "divq ")    ' nabla gange
    Text = Replace(Text, ChrW(8711) & VBA.ChrW(215), "curlq ")    ' nabla kryds - curl
    Text = Replace(Text, ChrW(8711) & "^2", "laplaceq ")    ' nabla
    Text = Replace(Text, ChrW(8711), "gradq ")    ' nabla
    Text = Replace(Text, VBA.ChrW(183), "*")    ' prik erstattes med gange
    Text = Replace(Text, VBA.ChrW(8901), "*")    ' \cdot , ved ikke hvorfor der tilsyneladende er to
    Text = Replace(Text, VBA.ChrW(8729), "*")    ' prik \cdot erstattes med gange
    Text = Replace(Text, VBA.ChrW(8226), "*")    ' tyk prik erstattes med gange
    Text = Replace(Text, VBA.ChrW(247), "/")    '
    Text = Replace(Text, VBA.ChrW(180) & VBA.ChrW(180), "^''")    ' der er tre forskellige '‰` apostrof lignende de forst*aa*es alle som diff
    Text = Replace(Text, VBA.ChrW(96) & VBA.ChrW(96), "^''")    ' der er tre forskellige '‰`
    Text = Replace(Text, VBA.ChrW(180), "^'")    ' der er tre forskellige '‰` apostrof lignende de forst*aa*es alle som diff
    Text = Replace(Text, VBA.ChrW(96), "^'")    ' der er tre forskellige '‰`
    Text = Replace(Text, VBA.ChrW(&H393) & "_incomplete ", "gamma_incomplete")
    Text = Replace(Text, VBA.ChrW(915), "gamma")    ' stort gammategn
    Text = Replace(Text, VBA.ChrW(8800), "#")    ' ikkeligmed
    Text = Replace(Text, VBA.ChrW(8804), "<=")    '
    Text = Replace(Text, VBA.ChrW(8805), ">=")    '
    Text = Replace(Text, VBA.ChrW(8710), "increment")    ' specielt delta increment
    Text = Replace(Text, "^ " & VBA.ChrW(176) & "C", "degC")    ' specielt gradtegn(\circ) og C
    Text = Replace(Text, "^ " & VBA.ChrW(8728) & "C", "degC")    ' specielt gradtegn(\circ) og C
    Text = Replace(Text, VBA.ChrW(8451), "degC")    ' specielt oC tegn
    Text = Replace(Text, VBA.ChrW(8728), "tgrader")    ' gradtegn
    Text = Replace(Text, "%", "/100")    ' procenttegn
    Text = Replace(Text, VBA.ChrW(160), "")    ' foran accent f.eks. vector pil 160 8407
    If Not MaximaComplex Then
        Text = Replace(Text, VBA.ChrW(8736), "Symangle")    '
    End If
    If CasType = 0 Then
        Text = Replace(Text, VBA.ChrW(176), "tgrader")    ' gradtegn
    End If

    ' som test flyttes disse fra slutningen til starten
    ' *oe**ae**aa* *AE**OE* ' *ae*ndres ikke tilbagee
    Text = Replace(Text, "*ae*", "ae")
    Text = Replace(Text, "*oe*", "oe")
    Text = Replace(Text, "*aa*", "aa")
    Text = Replace(Text, "*AE*", "AE")
    Text = Replace(Text, "*OE*", "OE")
    Text = Replace(Text, "*AA*", "AA")

    'gr*ae*ske bogstaver
    Text = Replace(Text, VBA.ChrW(916), "Delta")
    Text = Replace(Text, VBA.ChrW(948), "delta")
    Text = Replace(Text, VBA.ChrW(945), "alpha")
    Text = Replace(Text, VBA.ChrW(946), "beta")
    Text = Replace(Text, VBA.ChrW(947), "gammaLB")
    Text = Replace(Text, VBA.ChrW(952), "theta")
    Text = Replace(Text, VBA.ChrW(920), "Theta")
    Text = Replace(Text, VBA.ChrW(955), "lambda")
    Text = Replace(Text, VBA.ChrW(923), "Lambda")
    Text = Replace(Text, VBA.ChrW(956), "mu")
    Text = Replace(Text, VBA.ChrW(961), "rho")
    Text = Replace(Text, VBA.ChrW(963), "sigma")
    Text = Replace(Text, VBA.ChrW(931), "Sigma")
    Text = Replace(Text, VBA.ChrW(966), "varphi")
    Text = Replace(Text, VBA.ChrW(981), "phi")
    Text = Replace(Text, VBA.ChrW(934), "Phi")
    Text = Replace(Text, VBA.ChrW(949), "varepsilon")
    Text = Replace(Text, VBA.ChrW(1013), "epsilon")
    Text = Replace(Text, VBA.ChrW(968), "psi")
    Text = Replace(Text, VBA.ChrW(936), "Psi")
    Text = Replace(Text, VBA.ChrW(926), "Xi")
    Text = Replace(Text, VBA.ChrW(958), "xi")
    Text = Replace(Text, VBA.ChrW(935), "Chi")
    Text = Replace(Text, VBA.ChrW(967), "chi")
    Text = Replace(Text, VBA.ChrW(928), "Pi")
    Text = Replace(Text, VBA.ChrW(964), "tau")
    Text = Replace(Text, VBA.ChrW(957), "greek-nu")
    Text = Replace(Text, VBA.ChrW(954), "kappa")
    Text = Replace(Text, VBA.ChrW(951), "eta")
    Text = Replace(Text, VBA.ChrW(950), "zeta")
    Text = Replace(Text, VBA.ChrW(969), "omega")    ' lille omega


    ' _ symbol bliver til 1/2 eller tekst Symhalf afh. af kontekst. Symhalf hvis bogstavforan
    '    text = Replace(text, VBA.ChrW(189), "(1/2)") '
    '    text = Replace(text, VBA.ChrW(189), "Symhalf") ' 189 er symbolet for en halv
    
    ea.Text = Text
    ' fjern dobbelt mellemrum ' indført her 1.24. Er også lige før replacespacebymult. Indført for måske at fange nogle fejl med mange mellemrum
    Do While InStr(ea.Text, "  ") > 0
        ea.Text = Replace(ea.Text, "  ", " ")
    Loop

    Pos = 0
    Do
        Pos = InStr(Pos + 1, ea.Text, VBA.ChrW(189))
        If Pos > 0 Then
            If ea.IsLetterPos(Pos - 1, False) Then
                ea.RemoveChar Pos
                ea.InsertBeforePos "Symhalf", Pos
            Else
                ea.RemoveChar Pos
                ea.InsertBeforePos "(1/2)", Pos
            End If
        End If
    Loop While Pos > 0
    Text = ea.Text

    'lambertW0
    Do
        Pos = InStr(Text, "W_0")
        If Pos > 0 Then
            ea.Text = Text
            Text = Left(Text, Pos - 1) & "generalized_lambert_w(0," & right(Text, Len(Text) - Pos - 4)
        End If
    Loop While Pos > 0
    'lambertW-1
    Do
        Pos = InStr(Text, "W_(-1)")
        If Pos > 0 Then
            ea.Text = Text
            Text = Left(Text, Pos - 1) & "generalized_lambert_w(-1," & right(Text, Len(Text) - Pos - 7)
        End If
    Loop While Pos > 0

    ' autoregistrering af matrixparenteser
    If InStr(Text, "[") > 0 Then
        matrixstartbracket = "["
        matrixendbracket = "]"
    End If

    ' nu overfl*oe*dig: virker ikke med (ab)\vec
    '    text = Replace(text, VBA.ChrW(8407), "SymVecta") ' vektorpil
    '    text = Replace(text, VBA.ChrW(8401), "SymVectb") ' vektorpil men ikke hel pil
    '    text = Replace(text, VBA.ChrW(773), "SymVectc") ' matrixtegn
    '    text = Replace(text, VBA.ChrW(175), "SymVectc") ' matrixtegn \overbar

    ' vektortegn til SymVecta
    ea.Text = Text
    Pos = 1
    Do
        Pos = InStr(Pos, ea.Text, VBA.ChrW(8407))
        If Pos > 0 Then
            matvekt = True
            If ea.ChrByIndex(Pos - 1) = ")" Then
                s = ea.GetPrevBracketContent(Pos)
                s = Replace(s, " ", "SPC")
                MaximaInputStreng = MaximaInputStreng & "declare(" & s & "SymVecta,nonscalar)$"
                KillDef = KillDef & s & "SymVecta,"
                'AddDefinition ("definer: declare(" & s & "SymVecta,nonscalar)")
                ea.Text = Left(ea.Text, ea.Pos - 2) & s & "SymVecta" & right(ea.Text, Len(ea.Text) - Pos)
            Else
                '            AddDefinition ("definer: declare(" & ea.ChrByIndex(pos - 1) & "SymVecta,nonscalar)")
                s = ea.ChrByIndex(Pos - 1)
                MaximaInputStreng = MaximaInputStreng & "declare(" & s & "SymVecta,nonscalar)$"
                KillDef = KillDef & s & "SymVecta,"
                ea.Text = Left(ea.Text, Pos - 1) & "SymVecta" & right(ea.Text, Len(ea.Text) - Pos)
            End If
            Pos = Pos + 7
        End If
    Loop While Pos > 0

    ' vektortegn 2 ikke hel pil til SymVectb
    Pos = 1
    Do
        Pos = InStr(Pos, ea.Text, VBA.ChrW(8401))
        If Pos > 0 Then
            matvekt = True
            If ea.ChrByIndex(Pos - 1) = ")" Then
                s = ea.GetPrevBracketContent(Pos)
                ea.Text = Left(ea.Text, ea.Pos - 2) & s & "SymVectb" & right(ea.Text, Len(ea.Text) - Pos)
            Else
                ea.Text = Left(ea.Text, Pos - 1) & "SymVectb" & right(ea.Text, Len(ea.Text) - Pos)
            End If
            Pos = Pos + 7
        End If
    Loop While Pos > 0

    ' matrixtegn til SymVectc
    Pos = 1
    Do
        Pos = InStr(Pos, ea.Text, VBA.ChrW(773))
        If Pos > 0 Then
            matvekt = True
            If ea.ChrByIndex(Pos - 1) = ")" Then
                s = ea.GetPrevBracketContent(Pos)
                ea.Text = Left(ea.Text, ea.Pos - 2) & s & "SymVectc" & right(ea.Text, Len(ea.Text) - Pos)
            Else
                ea.Text = Left(ea.Text, Pos - 1) & "SymVectc" & right(ea.Text, Len(ea.Text) - Pos)
            End If
            Pos = Pos + 7
        End If
    Loop While Pos > 0


    ' matrixtegn \overbar til SymVectc
    
    Pos = 1
    Do
        Pos = InStr(Pos, ea.Text, VBA.ChrW(175))
        If Pos > 0 Then
            matvekt = True
            s = ea.GetNextVar(Pos)
            If Len(ea.Text) = ea.Pos Then
                ea.Text = Left(ea.Text, Pos - 1) & " " & s & "SymVectc"
            Else
                ea.Text = Left(ea.Text, Pos - 1) & " " & s & "SymVectc" & right(ea.Text, Len(ea.Text) - ea.Pos)
            End If
            Pos = Pos + 7
        End If
    Loop While Pos > 0

    Text = ea.Text

    If InStr(Text, VBA.ChrW(9608)) > 0 Then matvekt = True    '9608=matrixtegn
    If InStr(Text, VBA.ChrW(9632)) > 0 Then matvekt = True    '9632= andet vektortegn

    If matvekt Then  '9632?
        Text = Replace(Text, VBA.ChrW(215), "~")    ' vektorprodukt
    Else
        Text = Replace(Text, VBA.ChrW(215), "*")    ' kryds erstattes med gange
    End If

    'tv*ae*rvektor \hat
    Pos = 1
    Do
        Pos = InStr(Pos, Text, VBA.ChrW(770))
        If Pos > 0 Then
            ea.Text = Text
            If ea.ChrByIndex(Pos - 1) = ")" Then
                s = ea.GetPrevBracketContent(Pos)
                Text = Left(Text, ea.Pos - 2) & "tvvec(" & s & ")" & right(Text, Len(Text) - Pos)
                Pos = Pos + 4
            Else
                s = ea.GetPrevVar(Pos)
                Text = Left(Text, Pos - Len(s) - 2) & "tvvec(" & Mid(Text, Pos - Len(s) - 1, Len(s)) & ")" & right(Text, Len(Text) - Pos)
                Pos = Pos + 5
            End If
        End If
    Loop While Pos > 0

    '    text = Replace(text, VBA.ChrW(&H393) & " ", "Gamma")

    '   Fjern problematiske usynlige paranteser foran / ' l*oe*ses af den efterf*oe*lgende ogs*aa*
    '    Do
    '    pos = InStr(text, VBA.ChrW(12311) & "/")
    '    If pos > 0 Then
    '        ea.text = text
    '        ea.StartBracket = VBA.ChrW(12310)
    '        ea.EndBracket = VBA.ChrW(12311)
    '        ea.InsertTextAt ")", pos + 1
    '        ea.GetPrevBracketContent pos
    '        ea.pos = ea.pos - 1
    '        ea.FindPrevBreak
    '        ea.InsertTextAt "("
    '    End If
    '    Loop While pos > 0




    '   Fjern problematiske usynlige paranteser efter og foran /
    '    pos = 1
    '    Do
    '    pos = InStr(pos, text, VBA.ChrW(12310))
    '    If pos > 0 Then
    '        ib = False
    '        ea2.text = text
    '        ea2.pos = pos
    '        ea2.FindPrevBreak
    '        pos2 = ea2.pos
    '        ea.text = ea2.GetNextBracketContent(pos)
    '        If ea.ValidateBrackets And InStr(ea.text, "=") < 1 Then ' fors*oe*ger at l*oe*se problem hvor skjulte og alm paranteser er blandet som f.eks.[( ])
    '            If ea2.ChrByIndex(ea2.pos) = "/" Or ea2.ChrByIndex(pos2 - 1) = "/" Then
    '            ea2.InsertTextAt "(", pos2
    '            ea2.FindNextBreak
    '            ea2.RemoveChar
    '            ea2.InsertTextAt ")", ea2.pos
    '            ea2.RemoveChar pos + 1
    '            text = ea2.text
    '            End If
    '        End If
    'skipfb:
    '        pos = pos + 1
    '    End If
    '    Loop While pos > 0



    ' trig med store bogstaver
    If InStr(Text, "Sin") > 0 Then
        ea.Text = Text
        ea.ReplaceVar "Sin", "sin"
        Text = ea.Text
    End If
    If InStr(Text, "Cos") > 0 Then
        ea.Text = Text
        ea.ReplaceVar "Cos", "cos"
        Text = ea.Text
    End If
    If InStr(Text, "Tan") > 0 Then
        ea.Text = Text
        ea.ReplaceVar "Tan", "tan"
        Text = ea.Text
    End If
    If InStr(Text, "Log") > 0 Then
        ea.Text = Text
        ea.ReplaceVar "Log", "log"
        Text = ea.Text
    End If
    If InStr(Text, "Ln") > 0 Then
        ea.Text = Text
        ea.ReplaceVar "Ln", "ln"
        Text = ea.Text
    End If

    ''  her var disse f*oe*r men nu flyttet efter 8289 fjernes  text = Replace(text, "cosh^(-1)", "acosh")


    ' fjern +- symbol
    Pos = InStr(Text, VBA.ChrW(177))
    If Pos > 0 Then
        pm = InputBox(Sprog.A(421), Sprog.A(422), "-")
        If pm = "+" Or VBA.LCase(pm) = "plus" Then
            Text = Replace(Text, VBA.ChrW(177), "+")
        ElseIf pm = "-" Or VBA.LCase(pm) = "minus" Then
            Text = Replace(Text, VBA.ChrW(177), "-")
        Else
            MsgBox Sprog.A(423)
            Text = Replace(Text, VBA.ChrW(177), "+")
        End If
    End If

    ' kommaer og listeseparator
    If Not MaximaSeparator Then
        ea.Text = Text
        ea.ConvertDecSeparator
        Text = ea.Text
        '        text = Replace(text, ",", ".")
    End If
    Text = Replace(Text, ";", ",")


    ' d/dx notation gammel
    '    Do
    '    pos = InStr(text, VBA.ChrW(8518) & "/")
    '    If pos > 0 Then
    '        ea.text = text
    '        If Mid(ea.text, pos + 2, 3) = "(" & VBA.ChrW(8518) & "^" Then
    '            antalm = ea.ChrByIndex(pos + 5)
    '            varl = 3
    '            varl2 = 1
    '        Else
    '            antalm = 1
    '            varl = 0
    '            varl2 = 0
    '        End If
    '        ea.pos = pos + 3 + varl
    '        var = ea.GetNextVar
    '
    ''        var = Mid(text, pos + 3, 1)
    '        If ea.ChrByIndex(ea.pos + 1 + varl2) <> "(" And ea.ChrByIndex(ea.pos + 2 + varl2) <> "(" Then
    '            Do
    '                ea.pos = ea.pos + 1
    '                c = ea.ChrByIndex(ea.pos)
    '            Loop Until c <> " " And c <> ")"
    '            posstart = ea.pos
    '            posslut = ea.FindNextBreak(False)
    '            udtryk = ea.GetTextInterval(posstart, posslut)
    '            text = Left(ea.text, pos - 1) & "diff(" & udtryk & "," & var & "," & antalm & ")" & Right(ea.text, Len(ea.text) - posslut)
    '        Else
    '            udtryk = ea.GetNextBracketContent(ea.pos + varl2 + 1)
    '            posslut = ea.pos - 1
    '            text = Left(ea.text, pos - 1) & "diff(" & udtryk & "," & var & "," & antalm & ")" & Right(ea.text, Len(ea.text) - posslut)
    '        End If
    '    End If
    '    Loop While pos > 0


'    ' stykkevise funktioner defineret med en tuborg parantes. Ikke f*ae*rdig. Et problem at der inds*ae*ttes gangetegn i koden senere hen
'    ea.Text = Text
'    ea.SetNormalBrackets
'    i = 0
'    Do
'        pos = InStr(ea.Text, "{" & VBA.ChrW(9608))
'        If pos > 0 Then
'
''            pos2 = InStr(pos, ea.Text, VBA.ChrW(9508))
''            expr = Mid(ea.Text, pos + 2, pos2 - pos - 1)
'            expr = ea.GetNextBracketContent(pos)
'            arr2 = Split(expr, "@")
'            If InStr(expr, ", ") > 0 Then
'                c = ", "
'            Else
'                c = " ,"
'            End If
'            s = "(if "
'            For j = 0 To UBound(arr2)
'                arr3 = Split(arr2(j), c)
'                If UBound(arr3) > 0 Then
'                    s = s & arr3(1) & " then " & arr3(0)
'                End If
'                If j < UBound(arr2) Then s = s & " elseif "
'            Next
'            ea.Text = Left(ea.Text, pos - 1) & s & ")" & Right(ea.Text, Len(ea.Text) - pos - Len(expr) - 3)
'            i = i + 1
'        End If
'    Loop While pos > 0 And i < 1000
'    Text = ea.Text
    

    ' |x| -> abs(x) '
    ea.Text = Text
    i = 0
    Do
        Pos = InStr(ea.Text, VBA.ChrW(124))
        If Pos > 0 Then
            posb = Pos
            Do
                i = i + 1    ' forkert input kan ellers skabe uendelig l*oe*kke
                posb = posb + 1
                c = ea.ChrByIndex(posb)
                If c = "|" Then
                    If Not (ea.IsOperator(ea.ChrByIndex(posb - 1), 4)) Then
                        If matvekt Or InStr(ea.Text, VBA.ChrW(9632)) > 0 Then
                            If CasType = 0 Then
                                ea.Text = Left(ea.Text, Pos - 1) & "vecmag(" & Mid(ea.Text, Pos + 1, posb - Pos - 1) & ")" & right(ea.Text, Len(ea.Text) - posb)
                            Else
                                ea.Text = Left(ea.Text, Pos - 1) & "length(" & Mid(ea.Text, Pos + 1, posb - Pos - 1) & ")" & right(ea.Text, Len(ea.Text) - posb)
                            End If
                        Else
                            ea.Text = Left(ea.Text, Pos - 1) & "abs(" & Mid(ea.Text, Pos + 1, posb - Pos - 1) & ")" & right(ea.Text, Len(ea.Text) - posb)
                        End If
                        Exit Do
                    Else
                        Pos = posb
                    End If
                ElseIf c = "(" Then
                    ea.Pos = posb
                    posb = ea.FindEndOfBracket() - 1
                End If
            Loop While posb < Len(ea.Text) And i < 1000
            '        Loop While c <> "|" And c <> "(" And c <> "["

            posb = InStr(ea.Text, VBA.ChrW(124))
        End If
    Loop While Pos > 0 And i < 1000
    If i = 1000 Then
        MsgBox Sprog.A(424), vbOKOnly, Sprog.A(105)
    End If
    i = 0
    Text = ea.Text

    ' |x| -> abs(x) ' men der tages ikke h*oe*jde for paranteser
    '    ea.text = text
    '    Do
    '    pos = InStr(ea.text, VBA.ChrW(124))
    '    If pos > 0 Then
    '        posb = InStr(pos + 1, ea.text, VBA.ChrW(124))
    '        If ea.IsOperator(ea.ChrByIndex(posb - 1), 4) Then
    '        End If
    '        If matvekt Or InStr(ea.text, VBA.ChrW(9632)) > 0 Then
    '            ea.text = Left(ea.text, pos - 1) & "vecmag(" & Mid(ea.text, pos + 1, posb - pos - 1) & ")" & Right(ea.text, Len(ea.text) - posb)
    '        Else
    '            ea.text = Left(ea.text, pos - 1) & "abs(" & Mid(ea.text, pos + 1, posb - pos - 1) & ")" & Right(ea.text, Len(ea.text) - posb)
    '        End If
    '    End If
    '    Loop While pos > 0
    '    text = ea.text


    ' find underforst*aa*ede paranteser efter ^ og / ' (skal v*ae*re efter diff og andre funktioner med komma)
    ea.Text = Text
    ea.InsertBracketAfter ("/")
    ea.InsertBracketBefore ("/")
    Text = ea.Text
    
    ' inds*ae*t underforst*aa*ede gangetegn ' den køres også til slut efter alle mellemrum er fjernet, da der så kan opstå flere.
    'Indført her også for at omgå problem med fx at underforstået gangetegn ved 3(x+1) ikke bliver håndteret af findhierarkibreak
    ' Det giver dog problemer for 2^2x da
'    ea.text = text
'    ea.Pos = 1
'    ea.InsertMultSigns
'    text = ea.text
    
    ' d/dx notation
    Text = ConvertDifferentials(Text, VBA.ChrW(8518)) ' \dd
    If Left(Text, 5) = "Fejl!" Then
        Text = ea.Text
        ConvertErrorText = Text
        GoTo Slut
    End If
    Text = ConvertDifferentials(Text, VBA.ChrW(8706)) ' \partial
    If Left(Text, 5) = "Fejl!" Then
        Text = ea.Text
        ConvertErrorText = Text
        GoTo Slut
    End If
    If dAsDiffChr Then
        Text = ConvertDifferentials(Text, "d") ' normal d
        If Left(Text, 5) = "Fejl!" Then
            Text = ea.Text
            ConvertErrorText = Text
            GoTo Slut
        End If
    End If
    
    ' inds*ae*t parentes ved f.eks. log 2 -> log(2)
    ea.Text = Text
    ea.InsertFunctionBrackets
    ' {rod(5&3)}^2  hvor rod er rodtegn og {} er usynlig paranteser. Indf*oe*rt med geogebra, men nu l*oe*st i potensdelen
'    pos = InStr(ea.text, VBA.ChrW(12311) & "^")
'    Do While pos > 0
'        ea.SetInvisibleBrackets
'        ea.GetPrevBracketContent (pos + 1)
'        pos2 = ea.pos
'        ea.RemoveChar (pos)
'        ea.InsertAfterPos ")", pos - 1
'        ea.RemoveChar (pos2 - 1)
'        ea.InsertAfterPos "(", pos2 - 2
'        pos = InStr(ea.text, VBA.ChrW(12311) & "^")
'        ea.SetNormalBrackets
'    Loop
    
    ea.Text = Replace(ea.Text, VBA.ChrW(12310), "")    ' specielle usynlige paranteser fjernes   flyttet fra f*oe*r trig med store bogstaver for at f*aa* insertfunctionbrackets til at virke
    ea.Text = Replace(ea.Text, VBA.ChrW(12311), "")    ' specielle usynlige paranteser fjernes

    ea.Text = Replace(ea.Text, "cosh^(-1)", "acosh")
    ea.Text = Replace(ea.Text, "sinh^(-1)", "asinh")
    ea.Text = Replace(ea.Text, "tanh^(-1)", "atanh")
    ea.Text = Replace(ea.Text, "cos^(-1) ", "acos")    ' under sj*ae*ldne omst*ae*ndigheder kan der v*ae*re et mellemrum istedet for et 8289 tegn efter
    ea.Text = Replace(ea.Text, "sin^(-1) ", "asin")    ' det blive forst*aa*et som et gangetegn hvis ikke fjernes
    ea.Text = Replace(ea.Text, "tan^(-1) ", "atan")
    ea.Text = Replace(ea.Text, "cos^(-1)", "acos")
    ea.Text = Replace(ea.Text, "sin^(-1)", "asin")
    ea.Text = Replace(ea.Text, "tan^(-1)", "atan")

    ea.ConvertTrigPow ("sin")    ' skal v*ae*re efter inverse trig
    ea.ConvertTrigPow ("cos")
    ea.ConvertTrigPow ("tan")
    Text = ea.Text

    ' logaritmer med base forskellig fra 10 og e
    Dim base As String
    Dim spos As Integer
    ea.Text = Text
    spos = 1
    Do
        '    ea.text = text
        Pos = ea.FindChr("log_", spos)
        If Pos > 0 Then
            pos2 = InStr(Pos + 2, ea.Text, VBA.ChrW(8289))
            If Not pos2 > 0 Then
                base = ea.GetNextVar(Pos)
                base = right(base, Len(base) - 4)
            Else
                base = Mid(ea.Text, Pos + 4, pos2 - Pos - 4)
            End If
            ea.Text = Left(ea.Text, Pos + 2) & "base" & right(ea.Text, Len(ea.Text) - Pos - Len(base) - 4)
            ea.Pos = Pos + 7
            '        ea.InsertBeforePos ("(")
            '        ea.pos = ea.pos + 3
            ea.GetNextBracketContent
            ea.InsertBeforePos "," & base, ea.Pos - 1
            '        text = ea.text
            '        spos = ea.pos + 10
        End If
    Loop While Pos > 0
    Text = ea.Text


    Text = Replace(Text, VBA.ChrW(8289), "")    ' symbol der definerer funktion fjernes fra word syntaks

    'logaritmer
    If ConvertLnLog Then
        ea.Text = Text
        ea.ReplaceVar "log", "log10"
        Text = ea.Text
        '    spos = 1
        '    Do
        '    ea.text = text
        '    pos = ea.FindChr("log", spos)
        '    If pos > 0 Then
        '        ea.pos = pos
        '        ea.InsertBeforePos ("(")
        '        ea.pos = ea.pos + 3
        '        ea.GetNextBracketContent
        '        ea.InsertBeforePos ("/log(10))")
        '        text = ea.text
        '        spos = ea.pos + 10
        '    End If
        '    Loop While pos > 0
    End If

    ' 3 og 4 rod
    For rod = 3 To 4
        Do
            Pos = InStr(Text, VBA.ChrW(8728 + rod))
            If Pos > 0 Or pos4 > 0 Or pos5 > 0 Then
                ea.Text = Text
                ea.Pos = Pos + 1
                If Mid(Text, Pos + 1, 1) <> "(" Then
                    ea.InsertUnderstoodBracketPair
                End If
                ea.Pos = Pos
                Call ea.GetNextBracketContent    ' bare for at finde slut parantes
                Call ea.InsertBeforePos("^(1/" & rod & ")") ' metoden her kan faktisk give problemer ved fx \cbrt(4)^2 da denne skal fortolkes cbrt(4^2) og ikke cbrt(4)^2
                ea.Text = Replace(ea.Text, VBA.ChrW(8728 + rod), "", 1, 1)
                Text = ea.Text
            End If
        Loop While Pos > 0
    Next

    'kvadratrod
    Do
        Pos = InStr(Text, VBA.ChrW(8730))
        If Pos > 0 Then
            If Mid(Text, Pos + 1, 1) <> "(" Then
                Text = Replace(Text, VBA.ChrW(8730), "sqrt", 1, 1)
                Pos = Pos + 4
                ea.Text = Text
                ea.Pos = Pos
                ea.InsertUnderstoodBracketPair
                Text = ea.Text
            Else
                ea.Text = Text
                ea.Pos = Pos
                Arr = Split(ea.GetNextBracketContent, "&")
                pos2 = ea.Pos
                If UBound(Arr) = 0 Then
                    Text = Replace(Text, VBA.ChrW(8730), "sqrt", 1, 1)
                ElseIf UBound(Arr) = 1 Then
                    Call ea.InsertBeforePos("^(1/(" & Arr(0) & ")))")
                    ea.Text = Replace(ea.Text, VBA.ChrW(8730), "(", 1, 1)
                    posog = ea.FindChr("&", 1)
                    Text = Left(ea.Text, Pos + 1) & right(ea.Text, Len(ea.Text) - posog)
                ElseIf UBound(Arr) = -1 Then ' tilf*oe*jet v1.09 pga uendelig l*oe*kke ved mangl parentes
                    MsgBox Sprog.A(169), vbOKOnly, Sprog.Error
                    GoTo Fejl
                End If
            End If
        End If
    Loop While Pos > 0

    ' sumtegn
    Do
        Pos = InStr(Text, VBA.ChrW(8721))
        If Pos > 0 Then
            ea.Text = Text
            ea.Pos = Pos
            subs = ea.FindChr("_")
            sups = ea.FindChr("^")
            ends = ea.FindChr(VBA.ChrW(9618))
            Arr = Split(Mid(Text, subs + 2, sups - subs - 3), "=")
            If UBound(Arr) < 1 Then    ' hvis definition er = oversat til :
                Arr = Split(Mid(Text, subs + 2, sups - subs - 3), ":")
            End If
            If UBound(Arr) < 1 Then
                MsgBox Sprog.A(425), vbOKOnly, Sprog.Error '"Notationen for sumtegnet er ikke korrekt. Den nedre gr*ae*nse skal skrives f.eks. i=1"
                GoTo videre
            End If
            var = Arr(0)
            startn = Arr(1)
            slutn = Mid(Text, sups + 1, ends - sups - 1)
            sluts = ea.FindHierakiBreak(1)
            If ea.ChrByIndex(ea.Pos) = "+" Or ea.ChrByIndex(ea.Pos) = "-" Or ea.ChrByIndex(ea.Pos) = "*" Or ea.ChrByIndex(ea.Pos) = "/" Then sluts = sluts - 1
            Expr = Mid(Text, ends + 1, sluts - ends)
            '        If ea.ChrByIndex(ends + 1) = "(" Then
            '            expr = ea.GetNextBracketContent
            '            If ea.pos = Len(ea.text) Then
            '                sluts = ea.pos
            '            Else
            '                sluts = ea.pos - 1
            '            End If
            '        Else
            '            sluts = InStr(ends, text, "=") - 1
            '            If sluts < 1 Then sluts = Len(text)
            '            expr = Mid(text, ends + 1, sluts - ends)
            '        End If
            Text = Left(Text, Pos - 1) & "sum(" & Expr & "," & var & "," & startn & "," & slutn & ")" & right(Text, Len(Text) - sluts)
        End If
    Loop While Pos > 0

    ' produkttegn
    Do
        Pos = InStr(Text, VBA.ChrW(8719))
        If Pos > 0 Then
            ea.Text = Text
            ea.Pos = Pos
            subs = ea.FindChr("_")
            sups = ea.FindChr("^")
            ends = ea.FindChr(VBA.ChrW(9618))
            Arr = Split(Mid(Text, subs + 2, sups - subs - 3), "=")
            If UBound(Arr) < 1 Then    ' hvis definition er = oversat til :
                Arr = Split(Mid(Text, subs + 2, sups - subs - 3), ":")
            End If
            If UBound(Arr) < 1 Then
                MsgBox Sprog.A(426)
                GoTo videre
            End If
            var = Arr(0)
            startn = Arr(1)
            slutn = Mid(Text, sups + 1, ends - sups - 1)
            sluts = ea.FindHierakiBreak(1)
            Expr = Mid(Text, ends + 1, sluts - ends)
            Text = Left(Text, Pos - 1) & "product(" & Expr & "," & var & "," & startn & "," & slutn & ")" & right(Text, Len(Text) - sluts)
        End If
    Loop While Pos > 0

videre:

    'gr*ae*nsev*ae*rdier
    Pos = 1
    Do
        Pos = InStr(Pos, Text, "lim")
        '    pos = InStr(text, "lim" & VBA.ChrW(9516))
        '    pos = InStr(text, "lim" & VBA.ChrW(9516))
        If Pos > 0 Then
            ea.Text = Text
            If ea.ChrByIndex(Pos + 3) = VBA.ChrW(9516) Or ea.ChrByIndex(Pos + 3) = "_" Then
                ea.Pos = Pos
                var = ea.GetNextBracketContent
                ends = ea.Pos

                '        If ea.ChrByIndex(ends) = VBA.ChrW(12310) Then
                '            ea.StartBracket = VBA.ChrW(12310)
                '            ea.EndBracket = VBA.ChrW(12311)
                '            expr = ea.GetNextBracketContent
                '            If ea.pos = Len(ea.text) Then
                '                sluts = ea.pos
                '            Else
                '                sluts = ea.pos - 1
                '            End If
                '        ElseIf ea.ChrByIndex(ends) = "(" Then
                '            expr = ea.GetNextBracketContent
                '            If ea.pos = Len(ea.text) Then
                '                sluts = ea.pos
                '            Else
                '                sluts = ea.pos - 1
                '            End If
                '        Else
                '            sluts = InStr(ends, text, "=") - 1
                '            If sluts < 1 Then sluts = Len(text)
                '            expr = Mid(text, ends, sluts - ends + 1)
                '        End If
                sluts = ea.FindHierakiBreak(2)
                Expr = Mid(Text, ends, sluts - ends + 1)

                Arr = Split(var, VBA.ChrW(8594))
                var = Trim(Arr(0))
                DiffEqVar = var
                app = Trim(Arr(1))
                If right(app, 2) = "^+" Then
                    app = Left(app, Len(app) - 2) & ",plus"
                ElseIf right(app, 2) = "^-" Then
                    app = Left(app, Len(app) - 2) & ",minus"
                End If
                Text = Left(Text, Pos - 1) & "limit(" & Expr & "," & var & "," & app & ")" & right(Text, Len(Text) - sluts)
            Else
                Pos = Pos + 3
            End If
        End If
    Loop While Pos > 0


    'differentiering
    Dim subt As String, vararr() As String, s2 As String, f2 As String
    di = 0    ' hvis der er flere f'(2) skal der bruges flere variable
#If Mac Then
    Text = Replace(Text, "^('')", "^''")
    Text = Replace(Text, "^(''')", "^'''")
    Text = Replace(Text, "^('''')", "^''''")
#End If
    Do
        Pos = InStr(Text, "^'")
        If Pos > 0 Then
            ea.Text = Text
            ea.Pos = Pos
            '        f = ea.ChrByIndex(pos - 1)
            f = ea.GetPrevVar(Pos - 1)
            antalm = 0
            Do
                antalm = antalm + 1
            Loop Until ea.ChrByIndex(Pos + 1 + antalm) <> "'"
            If CasType = 0 Then
                antalmS = "round(" & antalm & ")"
            Else
                antalmS = CStr(antalm)
            End If
            If ea.ChrByIndex(Pos + antalm + 1) = " " Then    ' fjern evt. mellemrum efter '
                Text = Left(Text, Pos + 1) & right(Text, Len(Text) - Pos - 2)
                ea.Text = Text
            End If

            If ea.ChrByIndex(Pos + antalm + 1) = "(" Then
                '            var = ea.ChrByIndex(pos + 3 + antalm)
                var = ea.GetNextBracketContent(Pos + 1 + antalm)
                varl = Len(var) + 2
            Else
                If DiffEqVar = "" Then
                    var = "x"    ' hvis ingen variabel antages x
                Else
                    var = DiffEqVar    ' hvis diffligning s*aa* antages diffeqvar
                End If
                varl = 0
            End If
            
            pos2 = InStr(f, "_")
            subt = ""
            If pos2 > 0 Then
               subt = right(f, Len(f) - pos2)
               For i = 1 To Len(subt)
                  If Not ea.IsLetter(Mid(subt, i, 1)) Then
                     subt = ""
                     Exit For
                  End If
               Next
            End If

            If ea.ChrByIndex(Pos - 1) = ")" Then    ' hvis (...)'(x)
                ' find startparantes?
                udtryk = ea.GetPrevBracketContent(Pos)
                If ea.ChrByIndex(ea.Pos - 2) = "_" Then    '  f_(...)'(x)
                    f = ea.GetPrevVar(ea.Pos - 1) & "(" & udtryk & ")"
                    GoTo msat
                    '                MsgBox "Der er s*ae*nket skrift som ikke kan fortolkes", vbOKOnly, "Fejl"
                    '                GoTo fejl
                ElseIf ea.IsLetterPos(ea.Pos - 2) Then    ' sin(x)'
                    f = ea.GetPrevVar(ea.Pos - 1)
                    Text = Left(Text, ea.Pos - Len(f)) & "subst(" & var & ",qkqg" & di & ",diff(" & f & "(qkqg" & di & "),qkqg" & di & "," & antalmS & "))" & right(Text, Len(Text) - Pos - varl - antalm)
                    di = di + 1
                Else    ' (x^2)'
                    Text = Left(ea.Text, ea.Pos - 2) & "diff(" & udtryk & "," & var & "," & antalmS & ")" & right(ea.Text, Len(ea.Text) - Pos - varl - antalm)
                End If
            ElseIf ea.ChrByIndex(Pos - 2 - Len(subt)) = ")" And Len(subt) > 0 And Len(subt) = antalm Then ' partiel (...)_xy''
                  s = ea.GetPrevBracketContent(Pos - 2)
                  For i = 1 To antalm
                     s = "diff(" & s & "," & Mid(subt, i, 1) & ",1)"
                  Next
                 Text = Left(ea.Text, ea.Pos - 2) & s & right(ea.Text, Len(ea.Text) - Pos - varl - antalm)
            
            ElseIf Not ea.ChrByIndex(Pos + antalm + 1) = "(" Then    ' hvis y'  variabel ikke funktion
               If Len(subt) > 0 And Len(subt) = antalm Then ' hvis partiel y'
                  s = Left(f, pos2 - 1)
                  For i = 1 To antalm
                     s = "diff(" & s & "," & Mid(subt, i, 1) & ",1)"
                  Next
                 Text = Left(Text, Pos - 1 - Len(f)) & s & right(Text, Len(Text) - Pos - varl - antalm)
               Else
                Text = Left(Text, Pos - 1 - Len(f)) & "diff(" & f & "," & var & "," & antalmS & ")" & right(Text, Len(Text) - Pos - varl - antalm)
               End If
            ElseIf ea.ChrByIndex(Pos + antalm + 1) = "(" Then    ' hvis f'(x)
msat:
                If DiffEqVar = var Or CasType = 1 Or CasType = 2 Then  'hvis diffligning s*aa* ingen subst
                  If Len(subt) > 0 And Len(subt) = antalm Then ' hvis partiel
                        Text = Left(Text, Pos - 1 - Len(f)) & "diff(" & f & "(" & var & ")," & var & "," & antalmS & ")" & right(Text, Len(Text) - Pos - varl - antalm)
                    Else
                        Text = Left(Text, Pos - 1 - Len(f)) & "substitute(diff(" & f & "(qkqg" & di & "),qkqg" & di & "," & antalmS & "),qkqg" & di & "," & var & ")" & right(Text, Len(Text) - Pos - varl - antalm)
                        di = di + 1
                    End If
                Else    ' ikke diff ligning
                  If Len(subt) > 0 And Len(subt) = antalm Then ' hvis partiel
'                     s = Left(f, pos2 - 1) & "(" & var & ")"
                     f2 = Left(f, pos2 - 1)
                     s = "last(fundef(" & f2 & "))"
                     vararr = Split(var, ",") ' ;?
                     For i = 1 To antalm
'                        s = "diff(" & s & "," & Mid(subt, i, 1) & ",1)"
                        s = "diff(" & s & "," & Mid(subt, i, 1) & ",1)"
                     Next
                     s2 = "subst(["
                     For i = 0 To UBound(vararr)
                        s2 = s2 & "part(first(fundef(" & f2 & "))," & i + 1 & ")=" & vararr(i) & ","
                     Next
                     s = Left(s2, Len(s2) - 1) & "]," & s & ")"
                    Text = Left(Text, Pos - 1 - Len(f)) & s & right(Text, Len(Text) - Pos - varl - antalm)
                  Else ' alm ikke partiel f'(x)
                    Text = Left(Text, Pos - 1 - Len(f)) & "subst(" & var & ",qkqg" & di & ",diff(" & f & "(qkqg" & di & "),qkqg" & di & "," & antalmS & "))" & right(Text, Len(Text) - Pos - varl - antalm)
                    di = di + 1
                  End If
                End If
            Else
                GoTo Fejl
            End If
        End If
    Loop While Pos > 0

   ' nabla operator, grad og div
    Do
        Pos = InStr(Text, ChrW(8711) & ChrW(13))
        If Pos > 0 Then
            ea.Text = Text
            ea.Pos = Pos
        End If
    Loop While Pos > 0


    'integraler
    Dim integrand As String
    Pos = InStr(Text, VBA.ChrW(8747))
    If Pos > 0 Then
        ea.Text = Text
        ea.Text = Replace(ea.Text, "diff", "qiff")    ' d'et vil give problemer da det vil blive opfattet som afslutning af integrale
        ea.Text = Replace(ea.Text, "round", "rounq")    ' d'et vil give problemer da det vil blive opfattet som afslutning af integrale
        i = 0
        Do
            ea.Pos = 0
            ea.StartBracket = VBA.ChrW(8747)
            ea.EndBracket = VBA.ChrW(8518)    'differential d
            integrand = ea.GetNextBracketContent
            Pos = ea.Pos - Len(integrand)    'pos s*ae*ttes til start-integraltegnet
            pos2 = ea.Pos
            If Len(integrand) = 0 Then
                ea.EndBracket = "d"
                integrand = ea.GetNextBracketContent(Pos)
                Pos = ea.Pos - Len(integrand) - 2    'pos s*ae*ttes til start-integraltegnet
                pos2 = ea.Pos
                If Not (ea.IsLetterPos(ea.Pos, False) And Not (ea.IsLetterPos(ea.Pos - 2, False))) Then ' hvis enligt d eller som af variabel/funktion der ikke starter med d, s*aa* er det ikke slutningen af integralet. Derfor fors*oe*ges med simplere metode til at finde d
                    j = 0
                    Do
                        j = j + 1
                        pos3 = InStr(ea.Pos, ea.Text, "d")
                        ea.Pos = pos3 + 1
                    Loop While j < 100 And Not (ea.IsLetterPos(ea.Pos, False) And Not (ea.IsLetterPos(ea.Pos - 2, False)))
                    integrand = Mid(ea.Text, Pos + 1, pos3 - Pos - 1)
                    pos2 = ea.Pos
                End If
            End If
            var = ea.GetNextVar(ea.Pos - 1)
            If Len(integrand) > 0 Then
                If Len(integrand) < 3 And Mid(ea.Text, pos2 - Len(var), 2 + Len(var)) = "(" & var & ")" Then    ' hvis speciel notation med \int (dx)/(bla bla)
                    ea.SetNormalBrackets
                    integrand = ea.GetNextBracketContent(pos2 + 1)
                    pos2 = ea.Pos - 1
                End If
                If Len(var) < 2 Then
                    ConvertErrorText = Sprog.A(427)
                    GoTo Slut
                End If
                var = right(var, Len(var) - 1)
                integrand = Trim(integrand)
                pos2 = pos2 + Len(var)
                If Len(integrand) > 0 And Len(var) > 0 Then
                    IntegrationVars = IntegrationVars & var & ";"
                    poss = InStr(integrand, VBA.ChrW(9618))
                    posn = InStr(integrand, "_")
                    If Mid(integrand, posn + 1, 1) = "(" Then
                        ea2.Text = integrand
                        ea2.SetNormalBrackets
                        ea2.GetNextBracketContent (posn)
                        poso = InStr(ea2.Pos, integrand, "^")
                    Else
                        poso = InStr(integrand, "^")
                    End If
                    If posn > 0 And poso > 0 And posn < poss And poso < poss Then ' bestemt integrale
                        nedre = Mid(integrand, posn + 1, poso - posn - 1)
                        upper = Mid(integrand, poso + 1, poss - poso - 1)
                        '        text = Left(text, posn - 1) & Mid(text, poss + 1, pos2 - poss - 1) & "," & Mid(text, pos2 + 1, 1) & "," & nedre & "," & upper & ")" & Mid(text, pos2 + 2, Len(text) - pos2 - 1)
                        integrand = right(integrand, Len(integrand) - poss) & "," & var & "," & nedre & "," & upper & ")"
                        If Left(integrand, 1) = VBA.ChrW(9618) Then    'fjern funktiontegn
                            integrand = right(integrand, Len(integrand) - 1)
                        End If
                        If MaximaExact = 2 Then
                            '                        ea.text = Replace(ea.text, VBA.ChrW(8747), "NIntegrate(", 1, 1)
                            ea.Text = Left(ea.Text, Pos - 1) & "NIntegrate(" & integrand & right(ea.Text, ea.Length - pos2 - Len(var) + 2)
                        Else
                            ea.Text = Left(ea.Text, Pos - 1) & "Integrate(" & integrand & right(ea.Text, ea.Length - pos2 - Len(var) + 2)
                            '                        ea.text = Replace(ea.text, VBA.ChrW(8747), "Integrate(", 1, 1)
                        End If
                    Else    ' ubestemt integrale
                        integrand = integrand & "," & var & ")"
                        If Left(integrand, 1) = VBA.ChrW(9618) Then    'fjern funktionstegn
                            integrand = right(integrand, Len(integrand) - 1)
                        End If
                        ea.Text = Left(ea.Text, Pos - 1) & "integrate(" & integrand & right(ea.Text, ea.Length - pos2 - Len(var) + 2)
                    End If
                    i = i + 1
                End If
            End If
        Loop While i < 100 And InStr(ea.Text, VBA.ChrW(8747)) > 0 'Len(integrand) > 0
        ea.Text = Replace(ea.Text, "qiff", "diff")
        ea.Text = Replace(ea.Text, "rounq", "round")
        ea.SetNormalBrackets
        Text = ea.Text
    End If

    'dobbeltintegraler
    Dim var2 As String
    Pos = InStr(Text, VBA.ChrW(8748))
    If Pos > 0 Then
        ea.Text = Text
        ea.Text = Replace(ea.Text, "diff", "qiff")    ' d'et vil give problemer da det vil blive opfattet som afslutning af integrale
        i = 0
        Do
            ea.Pos = 0
            ea.StartBracket = VBA.ChrW(8748)
            ea.EndBracket = VBA.ChrW(8518)    'differential d
            integrand = ea.GetNextBracketContent
            Pos = ea.Pos - Len(integrand)    'pos s*ae*ttes til start-integraltegnet
            pos2 = ea.Pos
            If Len(integrand) = 0 Then
                ea.EndBracket = "d"
                integrand = ea.GetNextBracketContent
                Pos = ea.Pos - Len(integrand) - 2    'pos s*ae*ttes til start-integraltegnet
                pos2 = ea.Pos
            End If
            If Len(integrand) > 0 Then
                var = ea.GetNextVar(ea.Pos - 1)
                If Len(var) < 2 Then
                    ConvertErrorText = Sprog.A(427)
                    GoTo Slut
                End If
                var = right(var, Len(var) - 1)
                var2 = ea.GetNextVar(ea.Pos + 1)
                If Len(var2) < 2 Or (Left(var2, 1) <> "d" And Left(var2, 1) <> VBA.ChrW(8518)) Then
                    ConvertErrorText = Sprog.A(428)
                    GoTo Slut
                End If
                var2 = right(var2, Len(var2) - 1)
                integrand = Trim(integrand)
                pos2 = pos2 + Len(var)
                If Len(integrand) > 0 And Len(var) > 0 And Len(var2) > 0 Then
                    IntegrationVars = IntegrationVars & var & ";"
                    poss = InStr(integrand, VBA.ChrW(9618))
                    posn = InStr(integrand, "_")
                    If Mid(integrand, posn + 1, 1) = "(" Then
                        ea2.Text = integrand
                        ea2.SetNormalBrackets
                        ea2.GetNextBracketContent (posn)
                        poso = InStr(ea2.Pos, integrand, "^")
                    Else
                        poso = InStr(integrand, "^")
                    End If
                    If posn > 0 And poso > 0 And posn < poss And poso < poss Then
                        ConvertErrorText = Sprog.A(429)
                        GoTo Slut
                    Else    ' ubestemt integrale
                        integrand = integrand & "," & var & ")," & var2 & ")"
                        If Left(integrand, 1) = VBA.ChrW(9618) Then    'fjern funktionstegn
                            integrand = right(integrand, Len(integrand) - 1)
                        End If
                        ea.Text = Left(ea.Text, Pos - 1) & "integrate(integrate(" & integrand & right(ea.Text, ea.Length - pos2 - Len(var) - Len(var2))
                    End If
                    i = i + 1
                End If
            End If
        Loop While i < 100 And Len(integrand) > 0
        ea.Text = Replace(ea.Text, "qiff", "diff")
        ea.SetNormalBrackets
        Text = ea.Text
    End If

    'trippelintegraler
    Dim var3 As String
    Pos = InStr(Text, VBA.ChrW(8749))
    If Pos > 0 Then
        ea.Text = Text
        ea.Text = Replace(ea.Text, "diff", "qiff")    ' d'et vil give problemer da det vil blive opfattet som afslutning af integrale
        i = 0
        Do
            ea.Pos = 0
            ea.StartBracket = VBA.ChrW(8749)
            ea.EndBracket = VBA.ChrW(8518)    'differential d
            integrand = ea.GetNextBracketContent()
            Pos = ea.Pos - Len(integrand)    'pos s*ae*ttes til start-integraltegnet
            pos2 = ea.Pos
            If Len(integrand) = 0 Then
                ea.EndBracket = "d"
                integrand = ea.GetNextBracketContent
                Pos = ea.Pos - Len(integrand) - 2    'pos s*ae*ttes til start-integraltegnet
                pos2 = ea.Pos
            End If
            If Len(integrand) > 0 Then
                var = ea.GetNextVar(ea.Pos - 1)
                If Len(var) < 2 Then
                    ConvertErrorText = Sprog.A(427)
                    GoTo Slut
                End If
                var = right(var, Len(var) - 1)
                var2 = ea.GetNextVar(ea.Pos + 1)
                If Len(var2) < 2 Or (Left(var2, 1) <> "d" And Left(var2, 1) <> VBA.ChrW(8518)) Then
                    ConvertErrorText = Sprog.A(428) '"Der er problemer med syntaxen. Der skal bruges to dx (eller lign) som afslutning af et dobbletintegrale. M*aa*ske har du brugt d i anden sammenh*ae*ng. Brug da differential d \dd"
                    GoTo Slut
                End If
                var2 = right(var2, Len(var2) - 1)
                var3 = ea.GetNextVar(ea.Pos + 1)
                If Len(var3) < 2 Or (Left(var3, 1) <> "d" And Left(var3, 1) <> VBA.ChrW(8518)) Then
                    ConvertErrorText = Sprog.A(428) '"Der er problemer med syntaxen. Der skal bruges to dx (eller lign) som afslutning af et dobbletintegrale. M*aa*ske har du brugt d i anden sammenh*ae*ng. Brug da differential d \dd"
                    GoTo Slut
                End If
                var3 = right(var3, Len(var3) - 1)
                integrand = Trim(integrand)
                pos2 = pos2 + Len(var)
                If Len(integrand) > 0 And Len(var) > 0 And Len(var2) > 0 Then
                    IntegrationVars = IntegrationVars & var & ";"
                    poss = InStr(integrand, VBA.ChrW(9618))
                    posn = InStr(integrand, "_")
                    If Mid(integrand, posn + 1, 1) = "(" Then
                        ea2.Text = integrand
                        ea2.SetNormalBrackets
                        ea2.GetNextBracketContent (posn)
                        poso = InStr(ea2.Pos, integrand, "^")
                    Else
                        poso = InStr(integrand, "^")
                    End If
                    If posn > 0 And poso > 0 And posn < poss And poso < poss Then
                        ConvertErrorText = Sprog.A(429) '"Du kan ikke bruge dobbeltintegraltegnet til bestemte integraler. Brug to enkeltintegraler."
                        GoTo Slut
                    Else    ' ubestemt integrale
                        integrand = integrand & "," & var & ")," & var2 & ")," & var3 & ")"
                        If Left(integrand, 1) = VBA.ChrW(9618) Then    'fjern funktionstegn
                            integrand = right(integrand, Len(integrand) - 1)
                        End If
                        ea.Text = Left(ea.Text, Pos - 1) & "integrate(integrate(integrate(" & integrand & right(ea.Text, ea.Length - pos2 - Len(var) - Len(var2) - Len(var3) - 2)
                    End If
                    i = i + 1
                End If
            End If
        Loop While i < 100 And Len(integrand) > 0
        ea.Text = Replace(ea.Text, "qiff", "diff")
        ea.SetNormalBrackets
        Text = ea.Text
    End If


    'integraler (gammel)
    '        Do
    '            pos = InStr(text, VBA.ChrW(8747))
    '            If pos > 0 Then
    '                pos2 = pos
    'findslut:
    '                pos2d = InStr(pos2, text, VBA.ChrW(8518))  ' differential d
    '                pos2 = InStr(pos2, text, "d")
    '                ea.text = text
    '                If pos2 <= 0 Then
    '                    If pos2d <= 0 Then
    '                        ConvertErrorText = "Der mangler et d.. som afslutning af et integrale"
    '                        GoTo slut
    '                    Else
    '                        pos2 = pos2d
    '                    End If
    '                ElseIf pos2d > 0 And pos2d < pos2 Then
    '                    pos2 = pos2d
    '                End If
    '                ea.pos = pos2
    '                If ea.IsLetterPos(pos2 - 1, False) Then    ' hvis bogstav foran d, s*aa* indg*aa*r det i variabel
    '                    pos2 = pos2 + i
    '                    GoTo findslut
    '                End If
    '                i = 0
    '                Do
    '                    i = i + 1
    '                Loop While ea.IsLetterPos(pos2 + i, False) And i < 50
    '                var = Mid(text, pos2 + 1, i - 1)
    '                If var = "" Or var = "iff" Then    ' der vil v*ae*re d hvis diff under integraltegn
    '                    pos2 = pos2 + i
    '                    GoTo findslut
    '                End If
    '                IntegrationVars = IntegrationVars & var & ";"
    '
    '                poss = InStr(pos, text, VBA.ChrW(9618))
    '                posn = InStr(pos, text, "_")
    '                poso = InStr(pos, text, "^")
    '                If posn > 0 And poso > 0 And posn < poss And poso < poss Then
    '                    nedre = Mid(text, posn + 1, poso - posn - 1)
    '                    upper = Mid(text, poso + 1, poss - poso - 1)
    '                    '        text = Left(text, posn - 1) & Mid(text, poss + 1, pos2 - poss - 1) & "," & Mid(text, pos2 + 1, 1) & "," & nedre & "," & upper & ")" & Mid(text, pos2 + 2, Len(text) - pos2 - 1)
    '                    text = Left(text, posn - 1) & Mid(text, poss + 1, pos2 - poss - 1) & "," & var & "," & nedre & "," & upper & ")" & Mid(text, pos2 + 1 + Len(var), Len(text) - pos2 - Len(var))
    '                    If MaximaExact = 2 Then
    '                        text = Replace(text, VBA.ChrW(8747), "NIntegrate(", 1, 1)
    '                    Else
    '                        text = Replace(text, VBA.ChrW(8747), "Integrate(", 1, 1)
    '                    End If
    '                Else    ' ubestemt integrale
    '                    text = Left(text, pos2 - 1) & "," & var & ")" & Mid(text, pos2 + 1 + Len(var), Len(text) - pos2 - Len(var))
    '                    If Mid(text, pos + 1, 1) = VBA.ChrW(9618) Then    'fjern funktiontegn
    '                        text = Left(text, pos) & Right(text, Len(text) - pos - 1)
    '                    End If
    '                    text = Replace(text, VBA.ChrW(8747), "integrate(", 1, 1)
    '                End If
    '                '        text = Replace(text, VBA.ChrW(9618), "", 1, 1) ' tegn efter integrale, men bliver allerede fjernet
    '            End If
    '        Loop While pos > 0


    ' find underforst*aa*ede paranteser efter ^ og / ' (skal v*ae*re efter diff og andre funktioner med komma). Skal være før ea.InsertMultSigns
    ' flyttet op for at tage h*oe*jde for skjulte parenteser
    ea.Text = Text
    ea.SetNormalBrackets
    ea.InsertBracketAfter ("^")
    Text = ea.Text

    ' konverter a_i - > a[i] a_i,j -> a[i,j]  hvis sat i indstillinger
    If MaximaIndex Then
        Pos = 1
        Do
            Pos = InStr(Pos, ea.Text, "_")
            If Pos > 0 Then
                If ea.IsLetterPos(Pos - 1, False) Or ea.ChrByIndex(Pos - 1) = ")" Then
                    ea.RemoveChar Pos
                    ea.Pos = Pos
                    If ea.ChrByIndex(Pos) = "(" Then
                        Expr = ea.GetNextBracketContent()
                        ea.Text = Left(ea.Text, Pos - 1) & "[" & Expr & "]" & right(ea.Text, Len(ea.Text) - Pos - Len(Expr) - 1)
                    Else
                        posb = ea.FindHierakiBreak(5)    ' *ae*ndret 5 pga f_1^'(x)
                        Expr = ea.GetTextInterval(Pos, posb)
                        ea.Text = Left(ea.Text, Pos - 1) & "[" & Expr & "]" & right(ea.Text, Len(ea.Text) - Pos - Len(Expr) + 1)
                    End If
                Else
                    Pos = Pos + 1
                End If
            End If
        Loop While Pos > 0
    End If
spring:
    ea.RemoveSpaceAfterSubSuperscript ("_")
    '    ea.RemoveSpaceAfterSubSuperscript ("^")
    ' fjern dobbelt mellemrum ' dette også indført i starten af codeformaxima, så nok overflødigt her
    Do While InStr(ea.Text, "  ") > 0
        ea.Text = Replace(ea.Text, "  ", " ")
    Loop
    ea.ReplaceSpaceByMult    ' erstatter mellemrum med * men ikke hvis ulovlig
    Text = ea.Text

    Text = Replace(Text, " ", "")     ' evt. resterende mellemrum fjernes
    '    text = Replace(text, " ", "*")     ' skal v*ae*re efter diff og integral ' problematisk ved (x^2 )' -> diff(x^2*,x,1)


    Text = Replace(Text, "det(", "determinant2(")    ' lidt usikker da der kan v*ae*re funktioner der slutter p*aa* "det"
    ' Matricer
    Do
        Pos = InStr(Text, VBA.ChrW(9632))
        If Pos > 0 Then
            matvekt = True
            stpos = Pos - 1
            ea.Text = Text
            ea.Pos = Pos
            matrix = ea.GetNextBracketContent
            slpos = ea.Pos
            If ea.ChrByIndex(slpos + 1) = "*" Then
                Text = Left(Text, slpos) & " ## " & right(Text, Len(Text) - slpos - 1)
            End If
            Arr = Split(matrix, "@")
            If CasType = 1 Or CasType = 2 Then
                mtext = "{"
                For i = 0 To UBound(Arr)
                    mtext = mtext & "{" & Replace(Arr(i), "&", ",") & "},"
                Next
                mtext = Left(mtext, Len(mtext) - 1) & "}"
            Else
                mtext = "matrix("
                For i = 0 To UBound(Arr)
                    mtext = mtext & "[" & Replace(Arr(i), "&", ",") & "],"
                Next
                mtext = Left(mtext, Len(mtext) - 1) & ")"
            End If
            Text = Left(Text, stpos - 1) & mtext & right(Text, Len(Text) - slpos)
        End If
    Loop While Pos > 0

    'vektorer
    Do
        Pos = InStr(Text, VBA.ChrW(9608))
        If Pos > 0 Then
            matvekt = True
            stpos = Pos - 1
            ea.Text = Text
            ea.Pos = Pos
            matrix = ea.GetNextBracketContent
            slpos = ea.Pos
            Arr = Split(matrix, "@")
            If CasType = 1 Or CasType = 2 Then
                mtext = "{"
                For i = 0 To UBound(Arr)
                    mtext = mtext & "{" & Arr(i) & "},"
                Next
                mtext = Left(mtext, Len(mtext) - 1) & "}"
            Else
                mtext = "matrix("
                For i = 0 To UBound(Arr)
                    mtext = mtext & "[" & Arr(i) & "],"
                Next
                mtext = Left(mtext, Len(mtext) - 1) & ")"
            End If
            
            If Len(Text) - slpos >= 0 Then
                Text = Left(Text, stpos - 1) & mtext & right(Text, Len(Text) - slpos)
            Else
                Text = Left(Text, stpos - 1) & mtext
            End If
        End If
    Loop While Pos > 0




    ' inds*ae*t underforst*aa*ede gangetegn mellem bogstav mellemrum og bogstav
    ' virker men s*aa* kan man ikke lave funktioner med s*ae*nket skrift da der da er mellemrum
    '    ea.text = text
    '    pos = 1
    '    Do
    '    pos = InStr(pos, ea.text, " ")
    '    If pos > 0 Then
    '        If ea.IsLetterPos(pos - 1) And ea.IsLetterPos(pos + 1) Then
    '            ea.text = Left(ea.text, pos - 1) & "*" & Right(ea.text, Len(ea.text) - pos)
    '        End If
    '        pos = pos + 2
    '    End If
    '    Loop While pos > 0
    '    text = ea.text


    ' inds*ae*t underforst*aa*ede gangetegn ' skal v*ae*re efter fjern mellemrum
    ea.Text = Text
    ea.Pos = 1
    ea.InsertMultSigns
    Text = ea.Text
    If Radians Then
        Text = Replace(Text, VBA.ChrW(8736), " CVinkel ")    ' hvis complex s*aa* er vinkel pol*ae*r notation
    Else
        Text = Replace(Text, VBA.ChrW(8736), " CVinkelgrad ")    ' hvis complex s*aa* er vinkel pol*ae*r notation
    End If

    Text = Replace(Text, "gradq*", "gradq ")
    Text = Replace(Text, "divq*", "divq ")
    Text = Replace(Text, "curlq*", "curlq ")
    Text = Replace(Text, "laplaceq*", "laplaceq ")


    If matvekt Then
        If CASengine = 0 Then
        '        text = Replace(text, "])^", "])^^")  ' matrix ^ enkelt ^ opl*oe*fter kun hvert enkelt element
        Text = Replace(Text, "^", "^^")  ' definerede matricer blev ikke h*aa*ndteret korrekt med ovenst*aa*ende. Fors*oe*ger nu at erstatte alle og h*aa*ber det ikke giver andre problemer
'#If Mac Then ' from 1.18  windows is also sbcl
        Text = Replace(Text, "*", " ## ")    ' egentlig kun vigtig ved matrix gangetegn og problematisk n*aa*r * ikke er matrix multiplikation, da der s*aa* er problemer med hieraki 2/3 . matrix([2],[3]) her ganges 3 ind f*oe*rst. Dog l*oe*st med egen _ operator istedet for .
        Text = Replace(Text, "_", " ## ")    ' egentlig kun vigtig ved matrix gangetegn og problematisk n*aa*r * ikke er matrix multiplikation, da der s*aa* er problemer med hieraki 2/3 . matrix([2],[3]) her ganges 3 ind f*oe*rst. Dog l*oe*st med egen _ operator istedet for .
'#Else
'        text = Replace(text, "*", " _ ")    ' egentlig kun vigtig ved matrix gangetegn og problematisk n*aa*r * ikke er matrix multiplikation, da der s*aa* er problemer med hieraki 2/3 . matrix([2],[3]) her ganges 3 ind f*oe*rst. Dog l*oe*st med egen _ operator istedet for .
'#End If
        '        text = Replace(text, "*matrix", " . matrix")  ' matrix gangetegn og prikprodukt
        '        text = Replace(text, "*vect", " . vect")
        Else
            Text = Replace(Text, ChrW(215), ChrW(8855))
            Text = Replace(Text, "~", ChrW(8855)) ' vektorprodukt
        End If
    End If

    'erstat e med %e og i med %i      er flyttet til efter underforst*aa*ede gangetegn da % ikke g*ae*lder som bogstav
    Text = CodeConstantPercentSign(Text, "e")
    If MaximaComplex Then
        If CASengine = 0 Then
            Text = CodeConstantPercentSign(Text, "i")
        Else
            ea.Text = Text
            ea.ReplaceVar "i", VBA.ChrW(943)
            Text = ea.Text
        End If
    End If


    '    Text = Replace(Text, "e^", "exp")

    '    Text = Replace(Text, "{", "{{}")
    '    Text = Replace(Text, "}", "{}}")
    '    Text = Replace(Text, "+", "{+}")
    '    Text = Replace(Text, "^", "{^} ")
    '    Text = Replace(Text, "~", "{~}")
    '    Text = Replace(Text, "%", "{%}")
    '    Text = Replace(Text, "(", "{(}")
    '    Text = Replace(Text, ")", "{)}")
    '    Text = Replace(Text, "[", "{[}")
    '    Text = Replace(Text, "]", "{]}")



    '       text = Replace(text, "log", "1/log(10)*log") 'virker ikke hvis der bliver divideret med log
    '    text = Replace(text, "ln", "log")
    ea.Text = Text
    If ConvertLnLog Then
        ea.ReplaceVar "ln", "log"
    End If
    ea.ReplaceVar "arg", "carg"
    Text = ea.Text


    ' trig funktioner konverterer til radianer
    If Not Radians Then
        Text = ConvertDegreeToRad(Text, "sin")
        Text = ConvertDegreeToRad(Text, "cos")
        Text = ConvertDegreeToRad(Text, "tan")
        Text = ConvertDegreeToRad(Text, "sec")
        Text = ConvertDegreeToRad(Text, "cot")
        Text = ConvertDegreeToRad(Text, "csc")
    End If

    ' her var *ae**oe**aa* og gr*ae*ske f*oe*r
    If CasType = 0 Then
        Text = Replace(Text, VBA.ChrW(960), "%pi")
        Text = Replace(Text, "-" & VBA.ChrW(8734), "minf")
        Text = Replace(Text, VBA.ChrW(8734), "inf")
    Else
'        text = Replace(text, "-" & VBA.ChrW(8734), "minf")
        Text = Replace(Text, VBA.ChrW(8734), "infinity")
    End If


    'enheder
    If MaximaUnits Then
        Text = ConvertUnits(Text)
    Else
        Text = Replace(Text, VBA.ChrW(937), "Omega")    ' omega
    End If


    '    text = Replace(text, VBA.ChrW(8744), ",")  og
    '    text = Replace(text, VBA.ChrW(8743), ",") eller

    GoTo Slut:
Fejl:
    MsgBox Sprog.A(99), vbOKOnly, Sprog.Error
Slut:
    If Len(ConvertErrorText) > 0 Then
        MsgBox ConvertErrorText, vbOKOnly, Sprog.Error
        CodeForMaxima = ""
    Else
        CodeForMaxima = Text
    End If
End Function
Function ConvertDifferentials(Text As String, DiffChr As String) As String
    ' d/dx notation
    ' konverterer d/dx til maxima diff(...,x) notation
    ' DiffChr er symbolet for d. Det kan fx v*ae*re VBA.ChrW(8518) (\dd) eller VBA.ChrW(...)(\partial) eller m*aa*ske et almindeligt d
    
    Dim num As String, dem As String, vararr(9) As String, varind As Integer, antalm2 As Integer
    Dim udtryk1 As String, udtryk2 As String, Pos As Integer, pos3 As Integer, antalm As Integer
    Dim pos2 As Integer, var As String, i As Integer, udtryk As String, posstart As Integer, posslut As Integer
    Dim ea As New ExpressionAnalyser
    ea.SetNormalBrackets
    Do
        Pos = InStr(Text, "/(" & DiffChr)
        If Pos > 0 Then
            ea.Text = Text
            dem = ea.GetNextBracketContent(Pos + 1)
            pos3 = ea.Pos
            num = ea.GetPrevBracketContent(Pos)
            If Left(num, 1) <> DiffChr Then
                GoTo Slut
            End If
            Pos = ea.Pos
            varind = 0
            If Left(num, 2) = DiffChr & "^" Then
                antalm = Mid(num, 3, 1)
                pos2 = InStr(dem, DiffChr)
                Do
                    If Mid(dem, pos2 + 2, 1) = "^" Then
                        var = Mid(dem, pos2 + 1, 1)
                        antalm2 = Mid(dem, pos2 + 3, 1)
                        For i = varind To varind + antalm2 - 1
                            vararr(varind) = var
                            varind = varind + 1
                        Next
                    Else
                        vararr(varind) = Mid(dem, pos2 + 1, 1)
                        varind = varind + 1
                    End If
                    pos2 = InStr(pos2 + 1, dem, DiffChr)
                Loop While pos2 > 0
                If antalm <> varind Then
                    Text = Sprog.A(430)
                    GoTo Slut
                End If
                ElseIf Left(num, 1) = DiffChr Then
                    antalm = 1
                    vararr(0) = Mid(dem, 2, 1)
                Else
                    Text = Sprog.A(430)
                    GoTo Slut
            End If

            For i = 0 To antalm - 1
                udtryk1 = udtryk1 & "diff("
                udtryk2 = udtryk2 & "," & vararr(i) & ",1)"
            Next
            If Len(num) > 1 And Mid(num, 2, 1) <> "^" Then ' dy/dx eller df(x)/dx
                udtryk = right(num, Len(num) - 1)
                If Left(udtryk, 1) = "^" Then
                    udtryk = right(udtryk, Len(udtryk) - 1)
                End If
                Text = Left(Text, Pos - 2) & udtryk1 & udtryk & udtryk2 & right(Text, Len(Text) - pos3 + 1)
            ElseIf Mid(Text, pos3 + 1, 1) = "(" Then ' parentes lige efter d/dx operator
                udtryk = ea.GetNextBracketContent(pos3)
                Text = Left(Text, Pos - 2) & udtryk1 & udtryk & udtryk2 & right(Text, Len(Text) - ea.Pos + 1)
                ' Her mangler notation for df(x)/dx
            Else ' ingen parentes efter d/dx operator s*aa* udtrykket m*aa* findes via hieraki
                posstart = pos3
                ea.Pos = pos3 + 1
                posslut = ea.FindHierakiBreak(2)
                udtryk = Trim(Mid(Text, pos3, posslut - posstart + 1))
                Text = Left(Text, Pos - 2) & udtryk1 & udtryk & udtryk2 & right(Text, Len(Text) - posslut)
            End If
            '        var = Mid(text, pos + 3, 1)
            '            If di = 8 Then
            '            ElseIf di > 1 Then    ' hvis df(x)/dx
            '                udtryk = ea.GetTextInterval(pos - di + 1, pos - 1)
            '                text = Left(ea.text, pos - di - 1) & "diff(" & udtryk & "," & var & "," & antalm & ")" & Right(ea.text, Len(ea.text) - ea.pos)
            '            ElseIf ea.ChrByIndex(ea.pos + 1 + varl2) <> "(" And ea.ChrByIndex(ea.pos + 2 + varl2) <> "(" Then
            '                Do
            '                    ea.pos = ea.pos + 1
            '                    c = ea.ChrByIndex(ea.pos)
            '                Loop Until c <> " " And c <> ")"
            '                posstart = ea.pos
            '                posslut = ea.FindNextBreak(False)
            '                udtryk = ea.GetTextInterval(posstart, posslut)
            '                text = Left(ea.text, pos - 2) & "diff(" & udtryk & "," & var & "," & antalm & ")" & Right(ea.text, Len(ea.text) - posslut)
            '            Else    ' hvis paranteser efter d/dx
            '                udtryk = ea.GetNextBracketContent(ea.pos + varl2 + 1)
            '                posslut = ea.pos - 1
            '                text = Left(ea.text, pos - 2) & "diff(" & udtryk & "," & var & "," & antalm & ")" & Right(ea.text, Len(ea.text) - posslut)
            '            End If
        End If
    Loop While Pos > 0
Slut:
    ConvertDifferentials = Text
End Function
Function ConvertUnits(ByVal Text As String) As String
    Dim ea As New ExpressionAnalyser
    Text = Replace(Text, VBA.ChrW(937), "Ohm")    ' Omega er s*aa* ohm
    ea.Text = Text
    ea.ReplaceVar VBA.ChrW(176) & "C", "degC"  ' grad tegn og C
    ea.ReplaceVar "timer", "%timer"
    ea.ReplaceVar "min", "%min"
    ea.ReplaceVar "mm", "%%mm"
    ea.ReplaceVar "mum", "%mum"
    ea.ReplaceVar "mg", "%%mg"
    ea.ReplaceVar "mug", "%mug"
    ea.ReplaceVar "ms", "%%ms"
    ea.ReplaceVar "mus", "%mus"
    ea.ReplaceVar "mV", "%%mV"
    ea.ReplaceVar "muV", "%muV"
    ea.ReplaceVar "mA", "%%mA"
    ea.ReplaceVar "muA", "%muA"
    ea.ReplaceVar "mW", "%%mW"
    ea.ReplaceVar "muW", "%muW"
    ea.ReplaceVar "mJ", "%%mJ"
    ea.ReplaceVar "muJ", "%muJ"
    ea.ReplaceVar "mN", "%%mN"
    ea.ReplaceVar "muN", "%muN"
    ea.ReplaceVar "mC", "%%mC"
    ea.ReplaceVar "muC", "%muC"
    ea.ReplaceVar "mPa", "%%mPa"
    ea.ReplaceVar "muPa", "%muPa"
    ea.ReplaceVar "mF", "%%mF"
    ea.ReplaceVar "muF", "%muF"
    ea.ReplaceVar "mT", "%%mT"
    ea.ReplaceVar "muT", "%muT"
    ea.ReplaceVar "mK", "%%mK"
    ea.ReplaceVar "muK", "%muK"
    ea.ReplaceVar "mOhm", "%%mOhm"
    ea.ReplaceVar "muOhm", "%muOhm"
    ea.ReplaceVar "mpc", "%%mpc"
    ea.ReplaceVar "mupc", "%mupc"
    ConvertUnits = ea.Text

End Function
Function ConvertDegreeToRad(Text As String, trigfunc As String) As String
    Dim Pos, spos As Integer, s As String, v As String, f As Boolean
    Dim ea As New ExpressionAnalyser ', ea2 As New ExpressionAnalyser
    ea.SetNormalBrackets
'    ea2.SetNormalBrackets
    ea.Text = Text
    spos = 1

    Do
        Pos = ea.FindChr("a" & trigfunc, spos)
        If Pos > 0 Then
            If (Pos = 1 Or Not (ea.IsLetterPos(Pos - 1, False))) And Not (ea.IsLetterPos(Pos + Len(trigfunc) + 1, False)) Then
                ea.GetNextBracketContent
                ea.InsertBeforePos (")")
                ea.Pos = Pos
                ea.InsertBeforePos ("180/%pi*(")
                spos = Pos + 13
            Else
                spos = Pos + 3
            End If
        End If
    Loop While Pos > 0

    spos = 1
    Do
        Pos = ea.FindChr(trigfunc, spos)
        If Pos > 0 Then
            If (Pos = 1 Or Not (ea.IsLetterPos(Pos - 1, False))) And Not (ea.IsLetterPos(Pos + Len(trigfunc), False)) Then
                'If Not (ea.ChrByIndex(pos - 1) = "a") Then
                s = ea.GetNextBracketContent()
                If CASengine = 0 Then
                    ea.InsertBeforePos (")")
                    ea.Pos = Pos + Len(trigfunc)
                    ea.InsertAfterPos ("%pi/180*(")
                Else ' geogebra. Her fors*oe*ges at inds*ae*tte gradtegn, hvis der ikke indg*aa*r variabel indenfor parentes
                    If InStr(s, VBA.ChrW(176)) <= 0 Then
'                        ea2.text = s
'                        f = False
'                        Do
'                            v = ea2.GetNextVar
'                            If v <> "" Then
'                                If Not ea2.IsFunction(v) Then
'                                    f = True
'                                    Exit Do
'                                End If
'                            End If
'                        Loop While v <> ""
'                        If Not f Then
                            ea.Pos = ea.Pos - 1
                            ea.InsertBeforePos (VBA.ChrW(176))
'                        End If
                    End If
                End If
                spos = Pos + 13
            Else
                spos = Pos + 3
            End If
        End If
    Loop While Pos > 0

    ConvertDegreeToRad = ea.Text

End Function

Function CodeConstantPercentSign(Text As String, tegn As String) As String
' inds*ae*tter % foran konstant i udtryk hvis der ikke er bogstav foran eller bag
    Dim Pos As Integer
    Dim ea As New ExpressionAnalyser

    ea.Text = Text
    Pos = 1
    Do
        Pos = InStr(Pos, ea.Text, tegn)
        If Pos > 0 Then
            If Not (ea.IsLetterPos(Pos + 1, False) Or ea.ChrByIndex(Pos + 1) = "[" Or ea.IsLetterPos(Pos - 1, False)) Then
                ea.Text = Left(ea.Text, Pos - 1) & "%" & right(ea.Text, Len(ea.Text) - Pos + 1)
                Pos = Pos + 2
            Else
                Pos = Pos + 1
            End If
        End If
    Loop While Pos > 0

    CodeConstantPercentSign = ea.Text
End Function
Function ConvertToAscii(ByVal Text As String) As String
' konverterer fra word input til ascii

    Text = Replace(Text, VBA.ChrW(12310), "")    ' specielle usynlige paranteser fjernes
    Text = Replace(Text, VBA.ChrW(12311), "")    ' specielle usynlige paranteser fjernes
    Text = Replace(Text, VBA.ChrW(9633), "")    ' tegn foran paranteser der g*oe*r dem usynlige. Muligvis problematisk da paranteserne ikke fjernes. g*oe*r ogs*aa* br*oe*ker sm*aa*
    Text = Replace(Text, VBA.ChrW(9508), "")    ' tegn for tom plads ved ihvertfald [ ]
    Text = Replace(Text, VBA.ChrW(11), "")    ' shift-enter
    Text = Replace(Text, vbLf, "")    ' shift-enter og enter
    Text = Replace(Text, vbCrLf, "")
    Text = Replace(Text, vbCr, "")
    Text = Replace(Text, """", "")    ' apostrof fjernes
'#If Mac Then
'#Else
    Text = Replace(Text, "^'", "‰")    ' alm apostrof kan ikke bruges
    Text = Replace(Text, "'", "‰")    ' alm apostrof kan ikke bruges
'#End If
    Text = Replace(Text, "+-", "-")    ' plus-minus giver minus
    Text = Replace(Text, VBA.ChrW(8289), "")    ' symbol der definerer funktion fjernes fra word syntaks
    Text = Replace(Text, VBA.ChrW(8212), "+")    'dobbbelt minustegn giver plus
    '    text = Replace(text, VBA.ChrW(183), "*") ' prik erstattes med gange
    '    text = Replace(text, VBA.ChrW(215), "*") ' kryds erstattes med gange
    '    text = Replace(text, VBA.ChrW(8901), "*") ' \cdot , ved ikke hvorfor der tilsyneladende er to
    '    text = Replace(text, VBA.ChrW(8729), "*") ' prik \cdot erstattes med gange
    '    text = Replace(text, VBA.ChrW(8226), "*") ' tyk prik erstattes med gange
    '    text = Replace(text, VBA.ChrW(180), "^'") ' der er tre forskellige '‰` apostrof lignende de forst*aa*es alle som diff
    '    text = Replace(text, VBA.ChrW(96), "^'") ' der er tre forskellige '‰`

    Text = Replace(Text, ChrW(183), "*")  ' Siden 1.18 (SBCL) kan alm gangetegn ikke vises
    Text = Replace(Text, ChrW(215), "*")  '
    Text = Replace(Text, "  ", " ")
    '    text = Replace(text, " ", "") ' mellemrum fjernes ' problematisk n*aa*r , bruges som ; separator f(0 , 1) -> f(0,1)
    Text = Replace(Text, """", "")    ' apostrof fjernes
    Text = Replace(Text, "+-", "-")    ' plus-minus giver minus
    Text = Replace(Text, "%e", "e")    '
    '    text = Replace(text, "*ae*", "ae") '
    '    text = Replace(text, "*oe*", "oe") '
    '    text = Replace(text, "*aa*", "aa") '


    ConvertToAscii = Text
End Function
Function ConvertToWordSymbols(ByVal Text As String) As String
    On Error GoTo Fejl
    Dim gtegn As String
    Dim ea As New ExpressionAnalyser
    Dim Pos As Integer, pos2 As Integer
    Dim udtryk As String
    Dim possep As Integer
    Dim Arr As Variant
    Dim mtext As String
    Dim slutp As Integer
    Dim s As String, s2 As String

    ea.SetNormalBrackets
    gtegn = MaximaGangeTegn

    If CASengine = 0 Then
        Text = Replace(Text, " ", "")    ' mellemrum fjernes
    Else
        Text = Replace(Text, " + ", "+")
        Text = Replace(Text, " - ", "-")
        Text = Replace(Text, " * ", "*")
        Text = Replace(Text, " / ", "/")
    End If
    Text = Replace(Text, " . ", "*")    ' prikprodukt
    Text = Replace(Text, """", "")    ' apostrof fjernes
    Text = Replace(Text, "+-", "-")    ' plus-minus giver minus
    Text = Replace(Text, VBA.ChrW(11), "")    ' shift-enter
    Text = Replace(Text, vbCrLf, "")    ' enter
    Text = Replace(Text, VBA.ChrW(10), "")    ' shift-enter
    Text = Replace(Text, VBA.ChrW(13), "")    ' shift-enter
    Text = Replace(Text, "'lgog", "lgog")    '


    ea.Text = Text

    ea.RemoveFaultySeparator


    ' fx e^x^-1 rettes
    Pos = InStr(ea.Text, "^-")
    Do While Pos > 0
        ea.InsertTextAt "(", Pos + 1
        ea.Pos = Pos + 3
        ea.FindHierakiBreak 3
        ea.InsertTextAt ")", ea.Pos + 1
        Pos = InStr(Pos + 1, ea.Text, "^-")
    Loop


    ea.InsertBracketAfter ("^")


    ' eksponentiel notation positiv   ' f*oe*r insertbracketbefore da ellers 4E-13/x -> 4E-(13)/x men c^5.0E-1 -> c^0.5*10^-1 duer ikke
    Pos = 0
    Do
'#If Mac Then ' SBCL
        Pos = InStr(Pos + 1, ea.Text, "e+")
'#Else 'GCL
'        pos = InStr(pos + 1, ea.text, "E+")
'#End If
        If Pos > 1 Then
            If ea.IsNumber(Pos - 1) Then
                ea.Pos = Pos + 1
                Do
                    ea.Pos = ea.Pos + 1
                Loop While ea.IsNumber(ea.Pos)
                ' mellemrum fjernet for at f*aa* insertbracket til at virke
                ea.Text = Left(ea.Text, Pos - 1) & gtegn & "10^" & Mid(ea.Text, Pos + 2, ea.Pos - Pos - 2) & " " & right(ea.Text, Len(ea.Text) - ea.Pos + 1)
                'ea.text = Left(ea.text, pos - 1) & gtegn & "10^" & Mid(ea.text, pos + 2, ea.pos - pos - 2) & Right(text, Len(ea.text) - ea.pos + 1)
            End If
        End If
    Loop While Pos > 0

    ' eksponentiel notation negativ
    Pos = 0
    Do
'#If Mac Then ' SBCL
        Pos = InStr(Pos + 1, ea.Text, "e-")
'#Else 'GCL
'        pos = InStr(pos + 1, ea.text, "E-")
'#End If
        If Pos > 1 Then
            If ea.IsNumber(Pos - 1) Then
                ea.Pos = Pos + 1
                Do
                    ea.Pos = ea.Pos + 1
                Loop While ea.IsNumber(ea.Pos)
                ea.Text = Left(ea.Text, Pos - 1) & gtegn & "10^" & Mid(ea.Text, Pos + 1, ea.Pos - Pos - 1) & " " & right(ea.Text, Len(ea.Text) - ea.Pos + 1)
                '    ea.text = Left(ea.text, pos - 1) & gtegn & "10^" & Mid(ea.text, pos + 1, ea.pos - pos - 1) & Right(ea.text, Len(ea.text) - ea.pos + 1)
            End If
        End If
    Loop While Pos > 1

    ' efter eksponentiel da 1E-1*x g*aa*r galt
    ' 1* fjernes
    Pos = InStr(ea.Text, "1*")
    Do While Pos > 0
        If Pos = 1 Then
            ea.Text = Left(ea.Text, Pos - 1) & right(ea.Text, Len(ea.Text) - Pos - 1)    ' fjern 1*
        ElseIf Mid(ea.Text, Pos - 1, 1) = "*" Then  ' gange foran s*aa* fjern altid
            ea.Text = Left(ea.Text, Pos - 1) & right(ea.Text, Len(ea.Text) - Pos - 1)    ' fjern 1*
        ElseIf Mid(ea.Text, Pos - 1, 1) = "(" Then  ' parentes foran s*aa* fjern altid
            ea.Text = Left(ea.Text, Pos - 1) & right(ea.Text, Len(ea.Text) - Pos - 1)    ' fjern 1*
        ElseIf Mid(ea.Text, Pos - 1, 1) = ")" Then  ' parentes foran s*aa* fjern altid
            ea.Text = Left(ea.Text, Pos - 1) & right(ea.Text, Len(ea.Text) - Pos - 1)    ' fjern 1*
        ElseIf Mid(ea.Text, Pos - 1, 1) = "=" Then  ' ligmed foran s*aa* fjern altid
            ea.Text = Left(ea.Text, Pos - 1) & right(ea.Text, Len(ea.Text) - Pos - 1)    ' fjern 1*
        ElseIf Mid(ea.Text, Pos - 1, 1) = VBA.ChrW(8776) Then  ' ca. ligmed foran s*aa* fjern altid
            ea.Text = Left(ea.Text, Pos - 1) & right(ea.Text, Len(ea.Text) - Pos - 1)    ' fjern 1*
        ElseIf ea.IsNumber(Pos - 1) Then    ' tal foran s*aa* ingenting

        ElseIf ea.IsLetterPos(Pos - 1, True) Then    ' bogstav foran s*aa* ingenting

        ElseIf Mid(ea.Text, Pos - 1, 1) = "." Then    ' .1* s*aa* ingenting

        ElseIf Pos > 2 Then
            If Mid(ea.Text, Pos - 2, 2) = "^-" Then    ' ^-1 fjern ikke
            Else
                ea.Text = Left(ea.Text, Pos - 1) & right(ea.Text, Len(ea.Text) - Pos - 1)    ' fjern 1*
            End If
        Else
            ea.Text = Left(ea.Text, Pos - 1) & right(ea.Text, Len(ea.Text) - Pos - 1)    ' fjern 1*
        End If
        Pos = InStr(Pos + 1, ea.Text, "1*")
    Loop

    ' 1,0* fjernes og ,0* ->*
    Pos = InStr(ea.Text, ".0*")
    Do While Pos > 0
        If ea.ChrByIndex(Pos - 1) = "1" Then
            If Not (ea.IsNumber(Pos - 2)) Or Pos = 2 Then
                ea.Text = Left(ea.Text, Pos - 2) & right(ea.Text, Len(ea.Text) - Pos - 2)
            Else
                ea.Text = Left(ea.Text, Pos - 1) & right(ea.Text, Len(ea.Text) - Pos - 1)    ',0* ->*
            End If
        Else
            ea.Text = Left(ea.Text, Pos - 1) & right(ea.Text, Len(ea.Text) - Pos - 1)    ',0* ->*
        End If
        Pos = InStr(Pos + 1, ea.Text, ".0*")
    Loop
    ' tal,ikketal fjernes  f.eks. 240, -> 240  samt x*,45  ->x*0,45
    ' test: 2,0*a-*10^3*a-*g
    
'#If Mac Then
'** MAC   .0 p*aa* alle tal, fordi der bruges float istedet for dectalallnum
    Pos = InStr(ea.Text, ".0")
    Do While Pos > 0
            If Not (ea.IsNumber(Pos + 2)) Then
                ea.Text = Left(ea.Text, Pos - 1) & right(ea.Text, Len(ea.Text) - Pos - 1)
            End If
        Pos = InStr(Pos + 1, ea.Text, ".0")
    Loop
'#End If

    ' konverter a[i,j] -> a_(i,j)
    If MaximaIndex Then
        Pos = 1
        Do
            Pos = InStr(Pos + 1, ea.Text, "[")
            If Pos > 1 And ea.IsLetterPos(Pos - 1, False) Then
                ea.Pos = Pos
                ea.SetSquareBrackets
                udtryk = ea.GetNextBracketContent(Pos)
                '       if ea.ChrByIndex(ea.pos)=
                ea.Text = Left(ea.Text, Pos - 1) & "_(" & udtryk & ")" & right(ea.Text, Len(ea.Text) - ea.Pos + 1)

            End If
        Loop While Pos > 1
        ea.SetNormalBrackets
    End If


    ' fjern gangetegn der er underforst*aa*ede
    ' giver desv*ae*rre problemer bla. foran inverse trig  180sin^-1(12)/\pi og (1+x)=2 sin^(-1)(1/2)/\pi konvertes forkert til prof
    ' virker hverken hvis * erstattes med mellemrum eller fjernes
    ' r*ae*kkef*oe*lgen p*aa* disse to skal muligvis byttes om da det giver problemer ved identifikation af identitiske udtryk for beregn auto

    '    ea.InsertBracketBefore ("/") ' skal bruges for at removemultsign virker, men problem ved f.eks. matrix([5/6];[13/10]) -> matrix([5/6];([13)/10])
    If MaximaUnits Then ' version 1.07 fors*oe*ges n*aa*r enheder er sl*aa*et til
        ea.RemoveMultSign ' stadig problematisk ved f.eks. e^-x^2*(e^(x^2)+2*a-*c)/2
    End If
    If CASengine = 0 Then
        ea.SetNormalBrackets
    Else
        ea.SetTuborgBrackets
    End If
    ea.AddSpaceAfterSubSuperscript ("_")
    ea.SetNormalBrackets
    Text = ea.Text

    If Not MaximaSeparator Then
        Text = Replace(Text, ",", ";")
        Text = Replace(Text, ".", ",")
    End If
    Text = Replace(Text, "%pi", VBA.ChrW(960))
    Text = Replace(Text, "%e", "e")
    Text = Replace(Text, "%c", "c")
    Text = Replace(Text, "%k", "c_")
    Text = Replace(Text, "%r", "k_")
    Text = Replace(Text, "%z", VBA.ChrW(8484) & "_")
    Text = Replace(Text, "%Z", VBA.ChrW(8484))
    Text = Replace(Text, "#", VBA.ChrW(8800))    ' ikke ligmed

    If MaximaUnits Then
        Text = Replace(Text, VBA.ChrW(937), "Ohm")    ' Omega er s*aa* ohm
        Text = Replace(Text, "%%m", "m")
        Text = Replace(Text, "%mu", VBA.ChrW(956))
        ea.Text = Text
        ea.ReplaceVar "degC", VBA.ChrW(176) & "C"  ' grad tegn og C
        ea.ReplaceVar "%timer", "timer"
        ea.ReplaceVar "%min", "min"
        ea.ReplaceVar "aar", "*aa*r"
        Text = ea.Text
    End If


    Text = Replace(Text, "*", gtegn)    ' VBA.ChrW(183)


    '    Text = Replace(Text, "e", VBA.ChrW(183) & "10^ ")
    Text = Replace(Text, "exp(", "e^(")
    Text = Replace(Text, "log", "ln" & VBA.ChrW(8289))
    Text = Replace(Text, "lgog(", "log" & VBA.ChrW(8289) & "(")
    Text = Replace(Text, "sqrt", VBA.ChrW(8730))
    Text = Replace(Text, "cbrt", VBA.ChrW(8731)) ' fra geogebra

    '    text = Replace(text, "E+", gtegn & "10^")
    '    text = Replace(text, "E-", gtegn & "10^-")
        
    'rod af h*oe*jere grad.
    If CASengine > 0 Then 'Kun til GeoGebra
    Do
        Pos = InStr(Text, "nroot(")
        If Pos > 0 Then
            ea.Text = Text
            udtryk = ea.GetNextBracketContent(Pos + 4)
            pos2 = ea.Pos
'            Arr = Split(udtryk, ";")
            ea.Text = udtryk
            ea.SplitNonBracket Arr, ";"
            If UBound(Arr) > 0 Then
                Text = Left(Text, Pos - 1) & "" & VBA.ChrW(8730) & "(" & Arr(1) & "&" & Arr(0) & ")" & right(Text, Len(Text) - pos2 + 1)
            Else
                GoTo Fejl
            End If
        End If
    Loop While Pos > 0
    End If

    ' notequal(x,y) -> x=y
    Do
        Pos = InStr(Text, "notequal(")
        If Pos > 0 Then
            ea.Text = Text
            udtryk = ea.GetNextBracketContent(Pos + 7)
            Text = Left(Text, Pos - 1) & " " & Replace(udtryk, ListSeparator, VBA.ChrW(8800)) & " " & right(Text, Len(Text) - ea.Pos + 1)
        End If
    Loop While Pos > 0

    ' equal(x,y) -> x=y
    Do
        Pos = InStr(Text, "equal(")
        If Pos > 0 Then
            ea.Text = Text
            udtryk = ea.GetNextBracketContent(Pos + 4)
            Text = Left(Text, Pos - 1) & " " & Replace(udtryk, ListSeparator, "=") & " " & right(Text, Len(Text) - ea.Pos + 1)
        End If
    Loop While Pos > 0

    '    text = Replace(text, "if", vbCrLf & "hvis  ")
    '    text = Replace(text, "then", "  s*aa*  ")
    '    text = Replace(text, "else", vbCrLf & "ellers  ")

    ' abs(x) -> |x|
    Do
        Pos = InStr(Text, "abs(")
        If Pos > 0 Then
            ea.Text = Text
            udtryk = ea.GetNextBracketContent(Pos)
            Text = Left(Text, Pos - 1) & VBA.ChrW(124) & udtryk & VBA.ChrW(124) & right(Text, Len(Text) - ea.Pos + 1)
        End If
    Loop While Pos > 0

    ' 'vecmag(x) -> |x|
    Do
        Pos = InStr(Text, "'vecmag(")
        If Pos > 0 Then
            ea.Text = Text
            udtryk = ea.GetNextBracketContent(Pos + 6)
            Text = Left(Text, Pos - 1) & VBA.ChrW(124) & udtryk & VBA.ChrW(124) & right(Text, Len(Text) - ea.Pos + 1)
        End If
    Loop While Pos > 0

    'atan2(x,y) til atan(x/y)
    Pos = InStr(Text, "atan2")
    Do While Pos > 0
        ea.Text = Text
        ea.GetNextBracketContent (Pos)
        Text = Left(Text, ea.Pos - 1) & ")" & right(Text, Len(Text) - ea.Pos + 1)    ' slutparantes i n*ae*vner
        Text = Left(Text, Pos + 3) & "(" & right(Text, Len(Text) - Pos - 4)    ' , atan2 til atan(
        possep = InStr(Pos + 4, Text, ListSeparator)
        Text = Left(Text, possep - 1) & ")/(" & right(Text, Len(Text) - possep)    ' , til /

        Pos = InStr(Pos, Text, "atan2")
    Loop

    ' integrate til integraltegn
    Pos = InStr(Text, "integrate")
    Do While Pos > 0
        If Mid(Text, Pos - 1, 1) = "'" Then
            Text = Left(Text, Pos - 2) & right(Text, Len(Text) - Pos + 1)
            Pos = Pos - 1
        End If
        Dim udt As String
        ea.Text = Text
        ea.Pos = Pos + 8
        udt = ea.GetNextBracketContent()
        Arr = Split(udt, ListSeparator)
        If UBound(Arr) > 1 Then
            Text = Left(Text, Pos - 1) & VBA.ChrW(8747) & "_" & Arr(2) & "^" & Arr(3) & VBA.ChrW(9618) & Arr(0) & "  d" & Arr(1) & " " & right(Text, Len(Text) - Pos - Len(udt) - 10)
        Else
            Text = Left(Text, Pos - 1) & VBA.ChrW(8747) & VBA.ChrW(9618) & Arr(0) & "  d" & Arr(1) & " " & right(Text, Len(Text) - Pos - Len(udt) - 10)
        End If
        Pos = InStr(Pos, Text, "integrate")
    Loop

    ' differentier
    Pos = InStr(Text, "diff")
    Do While Pos > 0
        If Mid(Text, Pos - 1, 1) = "'" Then
            Text = Left(Text, Pos - 2) & right(Text, Len(Text) - Pos + 1)
            Pos = Pos - 1
        End If
        ea.Text = Text
        ea.Pos = Pos + 4
        udt = ea.GetNextBracketContent()
        Arr = Split(udt, ListSeparator)
        If UBound(Arr) > 1 Then
            If Arr(2) = "1" Then
                Text = Left(Text, Pos - 1) & VBA.ChrW(8518) & "/(" & VBA.ChrW(8518) & Arr(1) & ") (" & Arr(0) & ")" & " " & right(Text, Len(Text) - Len(udt) - 5 - Pos)
            Else
                Text = Left(Text, Pos - 1) & VBA.ChrW(8518) & "/(" & VBA.ChrW(8518) & "^" & Arr(2) & " " & Arr(1) & ") (" & Arr(0) & ")" & " " & right(Text, Len(Text) - Len(udt) - 6)
            End If
        Else
            Pos = Pos + 4
        End If
        Pos = InStr(Pos, Text, "diff")
    Loop

    ' Limit
    Pos = InStr(Text, "limit")
    Do While Pos > 0
        If Mid(Text, Pos - 1, 1) = "'" Then
            Text = Left(Text, Pos - 2) & right(Text, Len(Text) - Pos + 1)
            Pos = Pos - 1
        End If
        ea.Text = Text
        ea.Pos = Pos + 4
        udt = ea.GetNextBracketContent()
        Arr = Split(udt, ListSeparator)
        'lim" & VBA.ChrW(9516) & "(x" & VBA.ChrW(8594) & "1)" & VBA.ChrW(8289) & "(x+1)
        Text = Left(Text, Pos - 1) & "lim" & VBA.ChrW(9516) & "(" & Arr(1) & VBA.ChrW(8594) & Arr(2) & ")" & VBA.ChrW(8289) & "(" & Arr(0) & ") " & right(Text, Len(Text) - Pos - Len(udt) - 6)
        Pos = InStr(Pos, Text, "limit")
    Loop


    If Not Radians Then
        Text = Replace(Text, "sing(", "sin(")
        Text = Replace(Text, "asing(", "asin(")
        Text = Replace(Text, "cosg(", "cos(")
        Text = Replace(Text, "acosg(", "acos(")
        Text = Replace(Text, "tang(", "tan(")
        Text = Replace(Text, "atang(", "atan(")

        '        text = ConvertRadToDegree(text, "sin")
        '        text = ConvertRadToDegree(text, "cos")
        '        text = ConvertRadToDegree(text, "tan")
        Text = ConvertRadToDegree(Text, "sec")
        Text = ConvertRadToDegree(Text, "cot")
        Text = ConvertRadToDegree(Text, "csc")
    End If

    '    text = Replace(text, "acos", "cos^(-1) ")
    '    text = Replace(text, "asin", "sin^(-1) ")
    '    text = Replace(text, "atan", "tan^(-1) ")
    ea.Text = Text
    ea.ReplaceVar "cos", "cos" & VBA.ChrW(8289)
    ea.ReplaceVar "sin", "sin" & VBA.ChrW(8289)
    ea.ReplaceVar "tan", "tan" & VBA.ChrW(8289)
    ea.ReplaceVar "acos", "cos^(-1)" & VBA.ChrW(8289)
    ea.ReplaceVar "asin", "sin^(-1)" & VBA.ChrW(8289)
    ea.ReplaceVar "atan", "tan^(-1)" & VBA.ChrW(8289)
    ea.ReplaceVar "acosh", "cosh^(-1)" & VBA.ChrW(8289)
    ea.ReplaceVar "asinh", "sinh^(-1)" & VBA.ChrW(8289)
    ea.ReplaceVar "atanh", "tanh^(-1)" & VBA.ChrW(8289)
    ea.ReplaceVar "cosh", "cosh" & VBA.ChrW(8289)
    ea.ReplaceVar "sinh", "sinh" & VBA.ChrW(8289)
    ea.ReplaceVar "tanh", "tanh" & VBA.ChrW(8289)
    Text = ea.Text

    '        text = Replace(text, "acos(", "cos^(-1)" & VBA.ChrW(8289) & "(") ' parentes fors*oe*g ved invers ogs*aa*
    '        text = Replace(text, "asin(", "sin^(-1)" & VBA.ChrW(8289) & "(")
    '        text = Replace(text, "atan(", "tan^(-1)" & VBA.ChrW(8289) & "(")
    '        text = Replace(text, "cos(", "cos" & VBA.ChrW(8289) & "(")
    '        text = Replace(text, "sin(", "sin" & VBA.ChrW(8289) & "(")
    '        text = Replace(text, "tan(", "tan" & VBA.ChrW(8289) & "(")



    ' matricer
    Pos = InStr(Text, "matrix")
    Do While Pos > 0

        ea.Text = Text
        mtext = ea.GetNextBracketContent(Pos + 5)
        mtext = Replace(mtext, "]" & ListSeparator & "[", "@")
        mtext = Replace(mtext, ListSeparator, "&")

        mtext = matrixstartbracket & VBA.ChrW(9632) & "(" & Mid(mtext, 2, Len(mtext) - 2) & ")" & matrixendbracket
        Text = Left(Text, Pos - 1) & mtext & right(Text, Len(Text) - ea.Pos + 1)

        Pos = InStr(Pos + 1, Text, "matrix")
    Loop

    ' 'and i %if
    Pos = InStr(Text, "'and(")
    If Pos > 0 Then
        ea.Text = Text
        ea.Pos = Pos + 3
        udt = ea.GetNextBracketContent()
        udt = Replace(udt, ListSeparator, "  " & VBA.ChrW(8743) & "  ")
        Text = Left(Text, Pos - 1) & udt & right(Text, Len(Text) - ea.Pos + 1)
    End If

    ' %if
    Pos = InStr(Text, "%if")
    Do While Pos > 0
        ea.Text = Text
        mtext = ea.GetNextBracketContent(Pos + 2)
        slutp = ea.Pos
        Arr = Split(mtext, ListSeparator)
        Text = Left(Text, Pos - 1) & Arr(1) & "   for   " & Arr(0) & right(Text, Len(Text) - slutp + 1)
        Pos = InStr(Text, "%if")
    Loop

    'vecta_ -> vektorpil
    Pos = InStr(Text, "vecta_")
    Do While Pos > 0
        ea.Text = Text
        s = ea.GetNextVar(Pos + 6)
        Text = Left(Text, Pos - 1) & "(" & s & ")" & VBA.ChrW(160) & VBA.ChrW(8407) & right(Text, Len(Text) - ea.Pos)
        Pos = InStr(Text, "vecta_")
    Loop
    'vectb_ -> ikke hel vektorpil
    Pos = InStr(Text, "vectb_")
    Do While Pos > 0
        ea.Text = Text
        s = ea.GetNextVar(Pos + 6)
        Text = Left(Text, Pos - 1) & "(" & s & ")" & VBA.ChrW(160) & VBA.ChrW(8401) & right(Text, Len(Text) - ea.Pos)
        Pos = InStr(Text, "vectb_")
    Loop
    'vectc_ -> matrixstreg
    Pos = InStr(Text, "vectc_")
    Do While Pos > 0
        ea.Text = Text
        s = ea.GetNextVar(Pos + 6)
        Text = Left(Text, Pos - 1) & "(" & s & ")" & VBA.ChrW(160) & VBA.ChrW(773) & right(Text, Len(Text) - ea.Pos)
        Pos = InStr(Text, "vectc_")
    Loop

    'gr*ae*ske bogstaver
    Text = Replace(Text, "Delta", VBA.ChrW(916))
    Text = Replace(Text, "delta", VBA.ChrW(948))
    Text = Replace(Text, "alpha", VBA.ChrW(945))
    Text = Replace(Text, "beta", VBA.ChrW(946))
    Text = Replace(Text, "gamma_incomplete", "\Gamma_incomplete ")
    Text = Replace(Text, "gamma", VBA.ChrW(915))
    Text = Replace(Text, "gammaLB", VBA.ChrW(947))
    Text = Replace(Text, "theta", VBA.ChrW(952))
    Text = Replace(Text, "Theta", VBA.ChrW(920))
    Text = Replace(Text, "lambda", VBA.ChrW(955))
    Text = Replace(Text, "Lambda", VBA.ChrW(923))
    Text = Replace(Text, "rho", VBA.ChrW(961))
    Text = Replace(Text, "varphi", VBA.ChrW(966))
    Text = Replace(Text, "phi", VBA.ChrW(981))
    Text = Replace(Text, "Phi", VBA.ChrW(934))
    Text = Replace(Text, "varepsilon", VBA.ChrW(949))
    Text = Replace(Text, "epsilon", VBA.ChrW(1013))
    Text = Replace(Text, "psi", VBA.ChrW(968))
    Text = Replace(Text, "Psi", VBA.ChrW(936))
    Text = Replace(Text, "sigma", VBA.ChrW(963))
    Text = Replace(Text, "Sigma", VBA.ChrW(931))
    Text = Replace(Text, "mu", VBA.ChrW(956))
    Text = Replace(Text, "Ohm", VBA.ChrW(937))
    Text = Replace(Text, "Omega", VBA.ChrW(937))
    Text = Replace(Text, "omega", VBA.ChrW(969))
    Text = Replace(Text, "Xi", VBA.ChrW(926))
    Text = Replace(Text, "xi", VBA.ChrW(958))
    Text = Replace(Text, "Chi", VBA.ChrW(935))
    Text = Replace(Text, "chi", VBA.ChrW(967))
    Text = Replace(Text, "tau", VBA.ChrW(964))
    Text = Replace(Text, "Pi", VBA.ChrW(928))
    Text = Replace(Text, "greek-nu", VBA.ChrW(957))
    Text = Replace(Text, "kappa", VBA.ChrW(954))
    Text = Replace(Text, "zeta", VBA.ChrW(950))
    Text = Replace(Text, "eta", VBA.ChrW(951))    ' skal v*ae*re sidst da eta indg*aa*r i andre
    Text = Replace(Text, "increment", VBA.ChrW(8710))  ' specielt delta increment
    Text = Replace(Text, "Symhalf", VBA.ChrW(189))    ' _
    text = Replace(text, "degC", VBA.ChrW(8451))   ' specielt oC tegn
    Text = Replace(Text, "<=", VBA.ChrW(8804))  '
    Text = Replace(Text, ">=", VBA.ChrW(8805))    '
    Text = Replace(Text, "CVinkelO", VBA.ChrW(8736))    '
    Text = Replace(Text, "CVinkel", VBA.ChrW(8736))    '
    Text = Replace(Text, "Symangle", VBA.ChrW(8736))  '

    ' konverter vektor og matrixpile
    '    text = Replace(text, "SymVecta", VBA.ChrW(8407))   ' vektorpil
    '    text = Replace(text, "SymVectb", VBA.ChrW(8401)) ' vektorpil men ikke hel pil
    '    text = Replace(text, "SymVectc", VBA.ChrW(773)) ' matrixtegn
    ''    text = Replace(text, "SymVectc", VBA.ChrW(175))  ' matrixtegn \overbar

    ea.Text = Text
    ' SymVecta til vektorpil
    Pos = 1
    Do
        Pos = InStr(Pos, ea.Text, "SymVecta")
        If Pos > 0 Then
            ea.Pos = Pos
            s = ea.GetPrevVar(Pos)
            s = Left(s, Len(s) - (ea.Pos - Pos) - 1)
            If Pos + 7 = Len(ea.Text) Then
                ea.Text = Left(ea.Text, Pos - 1 - Len(s)) & "(" & s & ")" & VBA.ChrW(8407)
            Else
                ea.Text = Left(ea.Text, Pos - 1 - Len(s)) & "(" & s & ")" & VBA.ChrW(8407) & right(ea.Text, Len(ea.Text) - Pos - 7)
            End If
            Pos = Pos + 1
        End If
    Loop While Pos > 0

    ' SymVectb til halv vektorpil
    Pos = 1
    Do
        Pos = InStr(Pos, ea.Text, "SymVectb")
        If Pos > 0 Then
            ea.Pos = Pos
            s = ea.GetPrevVar(Pos)
            s = Left(s, Len(s) - (ea.Pos - Pos) - 1)
            If Pos + 7 = Len(ea.Text) Then
                ea.Text = Left(ea.Text, Pos - 1 - Len(s)) & "(" & s & ")" & VBA.ChrW(8401)
            Else
                ea.Text = Left(ea.Text, Pos - 1 - Len(s)) & "(" & s & ")" & VBA.ChrW(8401) & right(ea.Text, Len(ea.Text) - Pos - 7)
            End If
            Pos = Pos + 1
        End If
    Loop While Pos > 0

    ' SymVectc til halv vektorpil
    Pos = 1
    Do
        Pos = InStr(Pos, ea.Text, "SymVectc")
        If Pos > 0 Then
            ea.Pos = Pos
            s = ea.GetPrevVar(Pos)
            s = Left(s, Len(s) - (ea.Pos - Pos) - 1)
            If Pos + 7 = Len(ea.Text) Then
                ea.Text = Left(ea.Text, Pos - 1 - Len(s)) & "(" & s & ")" & VBA.ChrW(773)
            Else
                ea.Text = Left(ea.Text, Pos - 1 - Len(s)) & "(" & s & ")" & VBA.ChrW(773) & right(ea.Text, Len(ea.Text) - Pos - 7)
            End If
            Pos = Pos + 1
        End If
    Loop While Pos > 0

    ' lambert w
    Do
        Pos = InStr(ea.Text, "generalized_lambert_w")
        If Pos > 0 Then
        s = ea.GetNextBracketContent(Pos + 20)
        Arr = Split(s, ListSeparator)
        If UBound(Arr) = 0 Then
            Arr = Split(s, ",")
        End If
        s2 = Trim(Arr(0))
        If (s2 = "0" Or s2 = "-1") And UBound(Arr) = 1 Then
            ea.Text = Left(ea.Text, Pos - 1) & "W_(" & s2 & ") (" & Arr(1) & ")" & right(ea.Text, Len(ea.Text) - ea.Pos + 1)
        End If
        End If
    Loop While Pos > 0



    ' Disse kan v*ae*re resultat fra gr*ae*nsev*ae*rdi
    ea.ReplaceVar "inf", VBA.ChrW(8734)
    ea.ReplaceVar "minf", "-" & VBA.ChrW(8734)
    ea.ReplaceVar "infinity", VBA.ChrW(8734)
    ea.ReplaceVar "und", Sprog.A(431) '"ikke defineret"
    ea.ReplaceVar "ind", Sprog.A(432) ' "ikke defineret men begr*ae*nset"
    ea.ReplaceVar "tgrader", VBA.ChrW(176)     ' gradtegn
    Text = ea.Text

    '
    Text = Replace(Text, "SPC", " ")


    'ingen komplekse udtryk
    If Not MaximaComplex Then
        If InStr(Text, "%i") > 0 Or InStr(Text, VBA.ChrW(943)) > 0 Then
            Text = Sprog.A(433) & " " & VBA.ChrW(8477)
            GoTo Slut
        End If
    Else
        Text = Replace(Text, "%i", "i")
    End If

    '    text = Replace(text, "", "\")
    GoTo Slut
Fejl:
    MsgBox Sprog.A(100) & vbCrLf & vbCrLf & Text, vbOKOnly, Sprog.Error
Slut:
    ConvertToWordSymbols = Text
End Function
Function ConvertRadToDegree(Text As String, trigfunc As String) As String
    Dim Pos, spos As Integer
    Dim ea As New ExpressionAnalyser
    ea.StartBracket = "("
    ea.EndBracket = ")"
    ea.Text = Text
    spos = 1

    Do
        Pos = ea.FindChr("a" & trigfunc, spos)
        If Pos > 0 Then
            ea.GetNextBracketContent
            ea.InsertBeforePos (")")
            ea.Pos = Pos
            ea.InsertBeforePos (VBA.ChrW(960) & "/180" & MaximaGangeTegn & "(")
            spos = Pos + 11
        End If
    Loop While Pos > 0

    spos = 1
    Do
        Pos = ea.FindChr(trigfunc, spos)
        If Pos > 0 And Not (ea.ChrByIndex(Pos - 1) = "a") Then
            ea.GetNextBracketContent
            ea.InsertBeforePos (")")
            ea.Pos = Pos + Len(trigfunc)
            ea.InsertAfterPos ("180/" & VBA.ChrW(960) & MaximaGangeTegn & "(")
            spos = Pos + 11
        Else
            spos = spos + 3
        End If
    Loop While Pos > 0

    ConvertRadToDegree = ea.Text

End Function
Sub RemoveComplexSolutions()
'kun reelle l*oe*sninger ' problem at selvom der er i i kan det m*aa*ske reduceres til reelt
    Dim Arr As Variant
    Dim i As Integer
    If Not MaximaComplex Then
        '        If InStr(MaximaOutput, "%i") Then
        Arr = Split(MaximaOutput, ListSeparator)
        MaximaOutput = ""
        For i = 0 To UBound(Arr)
            If InStr(Arr(i), "%i") < 1 Then
                If MaximaOutput <> "" Then MaximaOutput = MaximaOutput & "    " & VBA.ChrW(8744) & "    "
                MaximaOutput = MaximaOutput & Arr(i)
            End If
        Next
        '        End If
    Else
        MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8744) & "    ")    ' komma erstattes med eller
    End If

End Sub
Function ConvertToLatex(Text As String) As String
    Dim Pos As Integer, pos1 As Integer, pos2 As Integer, pos3 As Integer, pos4 As Integer
    Dim ea As New ExpressionAnalyser
    Dim ea2 As New ExpressionAnalyser
    Dim s As String, t As String
    Dim Arr As Variant
    Dim ib As Boolean
    Dim rod As Integer
    Dim ends As Integer
    Dim subs As Integer, sups As Integer
    
    If ConvertTexWithMaxima Then
        ConvertToLatex = ConvertToLatexMaxima(Text)
        Exit Function
    End If
    ea.SetNormalBrackets
    ea2.SetTuborgBrackets


    Text = Replace(Text, "-" & VBA.ChrW(8734), "\minfty ")
    Text = Replace(Text, VBA.ChrW(8734), "\infty ")

    'diverse symboler
'    text = Replace(text, VBA.ChrW(183), "* ")    ' prik erstattes med gange
    Text = Replace(Text, VBA.ChrW(8901), VBA.ChrW(183))     ' \cdot , ved ikke hvorfor der tilsyneladende er to
    Text = Replace(Text, VBA.ChrW(8729), VBA.ChrW(183))    ' prik \cdot erstattes med gange (denne er \bullet)
    Text = Replace(Text, VBA.ChrW(8226), VBA.ChrW(183))    ' tyk prik erstattes med gange
    
    Text = Replace(Text, VBA.ChrW(247), "\div ")    ' divisionstegn med prikker over/under


    Text = Replace(Text, vbCrLf, "")
    Text = Replace(Text, vbCr, "")
    Text = Replace(Text, vbLf, "")

    Text = Replace(Text, VBA.ChrW(9633), "")    ' tegn foran paranteser der g*oe*r dem usynlige. Muligvis problematisk da paranteserne ikke fjernes. g*oe*r ogs*aa* br*oe*ker sm*aa*
    Text = Replace(Text, VBA.ChrW(9508), "")    ' tegn for tom plads ved ihvertfald [ ]
    Text = Replace(Text, VBA.ChrW(11), "")    ' shift-enter
    Text = Replace(Text, """", "")    ' apostrof fjernes
    Text = Replace(Text, VBA.ChrW(180) & VBA.ChrW(180), "^''")    ' der er tre forskellige '‰` apostrof lignende de forst*aa*es alle som diff
    Text = Replace(Text, VBA.ChrW(96) & VBA.ChrW(96), "^''")    ' der er tre forskellige '‰`
    Text = Replace(Text, VBA.ChrW(180), "^'")    ' der er tre forskellige '‰` apostrof lignende de forst*aa*es alle som diff
    Text = Replace(Text, VBA.ChrW(96), "^'")    ' der er tre forskellige '‰`
    '    text = Replace(text, VBA.ChrW(176), "\circ ") ' gradtegn
    Text = Replace(Text, VBA.ChrW(189), "\frac{1}{2} ")    ' _
    text = Replace(text, VBA.ChrW(8800), "\neq ")    ' ikkeligmed
    Text = Replace(Text, VBA.ChrW(177), "\pm ")    ' plus/minus
    Text = Replace(Text, VBA.ChrW(8723), "\mp ")    ' minus-plus
    Text = Replace(Text, VBA.ChrW(8804), "\leq ")    ' <=
    Text = Replace(Text, VBA.ChrW(8805), "\geq ")    ' >=
    Text = Replace(Text, VBA.ChrW(8776), "\approx ")    ' cirka lig
    Text = Replace(Text, VBA.ChrW(8773), "\cong ")    ' cirka lig
    Text = Replace(Text, VBA.ChrW(8656), "\Leftarrow ")    '
    Text = Replace(Text, VBA.ChrW(8658), "\Rightarrow ")    '
    Text = Replace(Text, VBA.ChrW(8660), "\Leftrightarrow ")    ' ensbetydende tegn
    Text = Replace(Text, VBA.ChrW(8661), "\Updownarrow ")    ' ensbetydende tegn op/ned
    Text = Replace(Text, VBA.ChrW(8657), "\Uparrow ")    ' dobbeltpil op
    Text = Replace(Text, VBA.ChrW(8659), "\Downarrow ")    ' dobbeltpil ned
    Text = Replace(Text, VBA.ChrW(8593), "\uparrow ")    '
    Text = Replace(Text, VBA.ChrW(8595), "\downarrow ")    '
    Text = Replace(Text, VBA.ChrW(8596), "\leftrightarrow ")    '
    Text = Replace(Text, VBA.ChrW(10187), "\partial ")    '
    Text = Replace(Text, VBA.ChrW(8469), "\mathbb{N} ")    ' symbol for naturlige tal
    Text = Replace(Text, VBA.ChrW(8484), "\mathbb{Z} ")    ' symbol for hele tal
    Text = Replace(Text, VBA.ChrW(8474), "\mathbb{Q} ")    ' symbol for rationelle tal
    Text = Replace(Text, VBA.ChrW(8477), "\mathbb{R} ")    ' symbol for reelle tal
    Text = Replace(Text, VBA.ChrW(8450), "\mathbb{C} ")    ' symbol for komplekse tal
    Text = Replace(Text, VBA.ChrW(8463), "\hbar ")    ' planck konstant/2pi
    Text = Replace(Text, VBA.ChrW(8704), "\forall ")    ' for alle tegn
    Text = Replace(Text, VBA.ChrW(8712), "\in ")    ' tilh*oe*rer tegn
    Text = Replace(Text, VBA.ChrW(8733), "\propto ")    ' propotionel tegn
    Text = Replace(Text, VBA.ChrW(8810), "\ll ")    ' <<
    Text = Replace(Text, VBA.ChrW(8811), "\gg ")    ' >>
    Text = Replace(Text, VBA.ChrW(8704), "\exists ")    ' eksisterer tegn
    Text = Replace(Text, VBA.ChrW(8708), "\exists ")    ' eksisterer tegn
    Text = Replace(Text, VBA.ChrW(8713), "\notin ")    ' tilh*oe*rer ikke
    Text = Replace(Text, VBA.ChrW(8709), "\emptyset ")    ' *OE* tomme m*ae*ngde
    Text = Replace(Text, VBA.ChrW(8711), "\nabla ")    ' nabla operator
    Text = Replace(Text, VBA.ChrW(8801), "\equiv ")    ' ligmed 3 streger
    Text = Replace(Text, VBA.ChrW(8710), "\inc ")    ' specielt delta increment
    Text = Replace(Text, VBA.ChrW(8943), "\ldots ")    ' ...
    Text = Replace(Text, VBA.ChrW(8942), "\vdots ")    ' 3 prikker lodret
    Text = Replace(Text, VBA.ChrW(8944), "\rddots ")    ' 3 prikker skr*aa*t op til h*oe*jre
    Text = Replace(Text, VBA.ChrW(8945), "\ddots ")    ' 3 prikker skr*aa*t ned til h*oe*jre
    Text = Replace(Text, VBA.ChrW(8756), "\therefore ")    ' 3 prikker i trekant
    Text = Replace(Text, VBA.ChrW(8757), "\because ")    ' 3 prikker i trekant
    Text = Replace(Text, VBA.ChrW(8834), "\subset ")    ' tilh*oe*rer m*ae*ngde (*ae*gte delm*ae*ngde)
    Text = Replace(Text, VBA.ChrW(8838), "\subseteq ")    ' tilh*oe*rer lig m*ae*ngde
    Text = Replace(Text, VBA.ChrW(8745), "\cap ")    ' foreningsm*ae*ngde
    Text = Replace(Text, VBA.ChrW(8746), "\cup ")    ' f*ae*llesm*ae*ngde
    Text = Replace(Text, VBA.ChrW(8764), "\sim ")    'ca ligmed en b*oe*lge
    Text = Replace(Text, VBA.ChrW(8771), "\simeq ")    ' ca ligmed en b*oe*lge en lige
    Text = Replace(Text, VBA.ChrW(8855), "\otimes ")    ' m*ae*ngdeprodukt
    Text = Replace(Text, VBA.ChrW(8614), "\mapsto ")    ' pil til funktionsdef
    Text = Replace(Text, VBA.ChrW(8640), "\rightharpoonup ")    ' pil h*oe*jre med kun halv pil
    Text = Replace(Text, VBA.ChrW(8651), "\lrhar ")    ' harpun h*oe*jre/venstre
    Text = Replace(Text, VBA.ChrW(8598), "\nwarrow ")    ' pil op til venstre
    Text = Replace(Text, VBA.ChrW(8599), "\nearrow ")    ' pil op til h*oe*jre
    Text = Replace(Text, VBA.ChrW(8600), "\nwarrow ")    ' pil ned h*oe*jre
    Text = Replace(Text, VBA.ChrW(8601), "\swarrow ")    ' pil ned venstre
    Text = Replace(Text, VBA.ChrW(8728), "\circ ")    ' grad-tegn
    Text = Replace(Text, VBA.ChrW(8451), "^\circ\text{C} ")    ' specielt oC tegn
    Text = Replace(Text, VBA.ChrW(8457), "^\circ\text{F} ")    ' specielt oF tegn
    Text = Replace(Text, VBA.ChrW(176), "^\circ ")    '
    Text = Replace(Text, VBA.ChrW(8736), "\angle ")    '
    Text = Replace(Text, VBA.ChrW(160), "")    ' foran accent f.eks. vector pil 160 8407
    Text = Replace(Text, VBA.ChrW(8743), "\wedge ")    ' og
    Text = Replace(Text, VBA.ChrW(8744), "\vee ")    ' eller

'    text = Replace(text, VBA.ChrW(8721), "\sum ")    ' sumtegn. Skal m*aa*ske nok laves mere avanceret
'    text = Replace(text, VBA.ChrW(8719), "\prod ")    ' produkttegn. Skal m*aa*ske nok laves mere avanceret
    Text = Replace(Text, VBA.ChrW(8748), "\iint ")    ' dobbeltintegraltegn. Skal m*aa*ske nok laves mere avanceret
    Text = Replace(Text, VBA.ChrW(8749), "\iiint ")    ' tripelintegraltegn. Skal m*aa*ske nok laves mere avanceret
    Text = Replace(Text, VBA.ChrW(8750), "\oint ")    ' kurveintegraltegn. Skal m*aa*ske nok laves mere avanceret


    '   Fjern problematiske usynlige paranteser efter og foran / m*aa*ske ikke n*oe*dvendig for latex? ihvertfald problematisk for x/10^2 -> x/(10)^2
    '    ea2.StartBracket = VBA.ChrW(12310)
    '    ea2.EndBracket = VBA.ChrW(12311)
    '    pos = 1
    '    Do
    '    pos = InStr(pos, text, VBA.ChrW(12310))
    '    If pos > 0 Then
    '        ib = False
    '        ea2.text = text
    '        ea2.pos = pos
    '        ea2.FindPrevBreak
    '        pos2 = ea2.pos
    '        ea.text = ea2.GetNextBracketContent(pos)
    '        If ea.ValidateBrackets And InStr(ea.text, "=") < 1 Then ' fors*oe*ger at l*oe*se problem hvor skjulte og alm paranteser er blandet som f.eks.[( ])
    '            If ea2.ChrByIndex(ea2.pos) = "/" Or ea2.ChrByIndex(pos2 - 1) = "/" Then
    '            ea2.InsertTextAt "(", pos2
    '            ea2.FindNextBreak
    '            ea2.RemoveChar
    '            ea2.InsertTextAt ")", ea2.pos
    '            ea2.RemoveChar pos + 1
    '            text = ea2.text
    '            End If
    '        End If
    'skipfb:
    '        pos = pos + 1
    '    End If
    '    Loop While pos > 0

    ea2.SetTuborgBrackets
    ea.Text = Text

    ea.Text = Replace(ea.Text, VBA.ChrW(12310), "")    ' specielle usynlige paranteser fjernes   flyttet fra f*oe*r trig med store bogstaver for at f*aa* insertfunctionbrackets til at virke
    ea.Text = Replace(ea.Text, VBA.ChrW(12311), "")    ' specielle usynlige paranteser fjernes

'lim
    Pos = -1
    Do
        Pos = InStr(Pos + 2, ea.Text, "lim")
        If Pos > 0 And ea.ChrByIndex(Pos - 1) <> "\" Then
            ea.Pos = Pos
            If ea.ChrByIndex(Pos + 3) = VBA.ChrW(9516) Then
                ea.ReplaceStringAt "_", Pos + 3
            End If
            subs = ea.FindChr("_")
            If ea.ChrByIndex(subs + 1) = "(" Then
                ea.SetNormalBrackets
                s = ea.GetNextBracketContent(subs)
                ea.ReplaceStringAt "{", subs + 1
                ea.ReplaceStringAt "}", subs + Len(s) + 2
            End If
            ea.SetNormalBrackets
            sups = ea.FindChr("^")
            If ea.ChrByIndex(sups + 1) = "(" Then
                ea.SetTuborgBrackets
                s = ea.GetNextBracketContent(sups)
                ea.ReplaceStringAt "{", sups + 1
                ea.ReplaceStringAt "}", sups + Len(s) + 2
            End If
            ends = ea.FindChr(VBA.ChrW(8289))
            If ends > 0 Then ea.ReplaceStringAt " ", ends
            ea.Text = Left(ea.Text, Pos - 1) & "\lim " & right(ea.Text, Len(ea.Text) - Pos - 2)
        End If
    Loop While Pos > 0

'    ea.text = Replace(ea.text, "cosh^(-1)", "cosh^{-1}")
'    ea.text = Replace(ea.text, "sinh^(-1)", "sinh^{-1}")
'    ea.text = Replace(ea.text, "tanh^(-1)", "tanh^{-1}")
'    ea.text = Replace(ea.text, "cos^(-1) ", "cos^{-1}")    ' under sj*ae*ldne omst*ae*ndigheder kan der v*ae*re et mellemrum istedet for et 8289 tegn efter
'    ea.text = Replace(ea.text, "sin^(-1) ", "sin^{-1}")    ' det blive forst*aa*et som et gangetegn hvis ikke fjernes
'    ea.text = Replace(ea.text, "tan^(-1) ", "tan^{-1}")
'    ea.text = Replace(ea.text, "cos^(-1)", "cos^{-1}")
'    ea.text = Replace(ea.text, "sin^(-1)", "sin^{-1}")
'    ea.text = Replace(ea.text, "tan^(-1)", "tan^{-1}")

    ea.Text = ConvertTrigPowLatex("sin", ea.Text) ' skal v*ae*re f*oe*r funktionstegn fjernes
    ea.Text = ConvertTrigPowLatex("cos", ea.Text) ' skal v*ae*re f*oe*r funktionstegn fjernes
    ea.Text = ConvertTrigPowLatex("tan", ea.Text) ' skal v*ae*re f*oe*r funktionstegn fjernes

    ea.Text = Replace(ea.Text, VBA.ChrW(8289), "") ' funktionstegn

    'fbox
    ea.Text = ConvertPreAccent(ea.Text, VBA.ChrW(9645), "\fbox{", "}")

    'overbrace
    ea.Text = ConvertPreAccent(ea.Text, VBA.ChrW(9182), "\overbrace{", "}")

    'underbrace
    ea.Text = ConvertPreAccent(ea.Text, VBA.ChrW(9183), "\underbrace{", "}")


    'vektor-notation med pil over bogstav
    ea.Text = ConvertAccent(ea.Text, VBA.ChrW(8407), "\overrightarrow{", "}")

    'tv*ae*rvektor \hat
    ea.Text = ConvertAccent(ea.Text, VBA.ChrW(770), "\widehat{", "}")
    
    '\overline
    ea.Text = ConvertAccent(ea.Text, VBA.ChrW(773), "\overline{", "}")
    'dobbelt overline
    ea.Text = ConvertAccent(ea.Text, VBA.ChrW(831), "\overline{\overline{", "}}")
    
    'halvpil
    ea.Text = ConvertAccent(ea.Text, VBA.ChrW(8401), "\overset{\rightharpoonup}{", "}")

    'dobbeltpil over
    ea.Text = ConvertAccent(ea.Text, VBA.ChrW(8417), "\overleftrightarrow{", "}")
    
    'venstrepil over
    ea.Text = ConvertAccent(ea.Text, VBA.ChrW(8406), "\overleftarrow{", "}")

    '\widetilde
    ea.Text = ConvertAccent(ea.Text, VBA.ChrW(771), "\widetilde{", "}")

    '\dot
    ea.Text = ConvertAccent(ea.Text, VBA.ChrW(775), "\dot{", "}")
    '\ddot
    ea.Text = ConvertAccent(ea.Text, VBA.ChrW(776), "\ddot{", "}")
    '\ddot
    ea.Text = ConvertAccent(ea.Text, VBA.ChrW(8411), "\dddot{", "}")

    '\underset
    Do
        Pos = InStr(ea.Text, VBA.ChrW(9516))
        If Pos > 0 Then
            If ea.ChrByIndex(Pos + 1) = "(" Then
                ea.SetNormalBrackets
                s = ea.GetNextBracketContent(Pos + 1)
                pos2 = ea.Pos - 1
            Else
                s = ea.GetNextVar(Pos)
                pos2 = ea.Pos
            End If
            If ea.ChrByIndex(Pos - 1) = ")" Then
                ea.SetNormalBrackets
                t = ea.GetPrevBracketContent(Pos)
                pos1 = ea.Pos - 2
            ElseIf ea.ChrByIndex(Pos - 1) = "}" Then
                ea.SetTuborgBrackets
                t = ea.GetPrevBracketContent(Pos)
                pos1 = ea.Pos - 2
                Do While ea.IsLatexLetterPos(pos1)
                    pos1 = pos1 - 1
                Loop
                If ea.ChrByIndex(pos1) = "\" Then
                    t = Mid(ea.Text, pos1, Pos - pos1)
                    pos1 = pos1 - 1
                End If
            Else
                t = ea.GetPrevVar(Pos)
                pos1 = ea.Pos - Len(t)
            End If
            ea.Text = Left(ea.Text, pos1) & "\underset{" & t & "}{" & s & "}" & right(ea.Text, Len(ea.Text) - pos2)
        End If
    Loop While Pos > 0
    
    

    ' sumtegn
    Do
        Pos = InStr(ea.Text, VBA.ChrW(8721))
        If Pos > 0 Then
            ea.Pos = Pos
            subs = ea.FindChr("_")
            If ea.ChrByIndex(subs + 1) = "(" Then
                ea.SetNormalBrackets
                s = ea.GetNextBracketContent(subs)
                ea.ReplaceStringAt "{", subs + 1
                ea.ReplaceStringAt "}", subs + Len(s) + 2
            End If
            ea.SetNormalBrackets
            sups = ea.FindChr("^")
            If ea.ChrByIndex(sups + 1) = "(" Then
                ea.SetTuborgBrackets
                s = ea.GetNextBracketContent(sups)
                ea.ReplaceStringAt "{", sups + 1
                ea.ReplaceStringAt "}", sups + Len(s) + 2
            End If
            ends = ea.FindChr(VBA.ChrW(9618))
            If ends > 0 Then ea.ReplaceStringAt " ", ends
            ea.Text = Left(ea.Text, Pos - 1) & "\sum " & right(ea.Text, Len(ea.Text) - Pos)
        End If
    Loop While Pos > 0

    ' produkttegn
    Do
        Pos = InStr(ea.Text, VBA.ChrW(8719))
        If Pos > 0 Then
            ea.Pos = Pos
            subs = ea.FindChr("_")
            If ea.ChrByIndex(subs + 1) = "(" Then
                ea.SetNormalBrackets
                s = ea.GetNextBracketContent(subs)
                ea.ReplaceStringAt "{", subs + 1
                ea.ReplaceStringAt "}", subs + Len(s) + 2
            End If
            ea.SetNormalBrackets
            sups = ea.FindChr("^")
            If ea.ChrByIndex(sups + 1) = "(" Then
                ea.SetTuborgBrackets
                s = ea.GetNextBracketContent(sups)
                ea.ReplaceStringAt "{", sups + 1
                ea.ReplaceStringAt "}", sups + Len(s) + 2
            End If
            ends = ea.FindChr(VBA.ChrW(9618))
            If ends > 0 Then ea.ReplaceStringAt " ", ends
            ea.Text = Left(ea.Text, Pos - 1) & "\prod " & right(ea.Text, Len(ea.Text) - Pos)
        End If
    Loop While Pos > 0

    ' integrale
    Pos = InStr(ea.Text, VBA.ChrW(8747))
    Do While Pos > 0
        pos1 = InStr(Pos, ea.Text, "_")
        pos2 = InStr(Pos, ea.Text, "^")
        pos3 = InStr(Pos, ea.Text, VBA.ChrW(9618))
        If pos3 = 0 Then pos3 = Pos
        If pos1 > 0 Then
            Dim ul As String, ll As String
            ll = Mid(ea.Text, pos1 + 1, pos2 - pos1 - 1)
            ul = Mid(ea.Text, pos2 + 1, pos3 - pos2 - 1)
            If Left(ll, 1) = "(" And right(ll, 1) = ")" Then ll = Mid(ll, 2, Len(ll) - 2) ' Word alm inds*ae*tter parenteser som ikke skal ses. Hvis der skal v*ae*re parenteser s*aa* er der dobbelt
            If Left(ul, 1) = "(" And right(ul, 1) = ")" Then ul = Mid(ul, 2, Len(ul) - 2)
            ea.Text = Left(ea.Text, Pos - 1) & "\int_{" & ll & "}^{" & ul & "} " & right(ea.Text, Len(ea.Text) - pos3)
        Else    ' ubestemt
            ea.Text = Left(ea.Text, Pos - 1) & "\int " & right(ea.Text, Len(ea.Text) - pos3)
        End If
        Pos = InStr(ea.Text, VBA.ChrW(8747))
    Loop

    ' ^
    Pos = InStr(ea.Text, "^")
    Do While Pos > 0
        If ea.ChrByIndex(Pos + 1) = "(" Then
            s = Trim(ea.GetNextBracketContent(Pos))
            If Not (Left(s, 1) = "{" And right(s, 1) = "}") Then s = "{" & s & "}"
            ea.Text = Left(ea.Text, Pos) & s & right(ea.Text, Len(ea.Text) - ea.Pos + 1)
        ElseIf ea.ChrByIndex(Pos + 1) = "{" Then ' hvis allerede er konverteret
        Else
            ea.Pos = Pos + 1
            ea.FindHierakiBreak 4, False
            s = ea.SubString(Pos + 1, ea.Pos)
            If Not (Left(s, 1) = "{" And right(s, 1) = "}") Then s = "{" & s & "}"
            ea.Text = Left(ea.Text, Pos) & s & right(ea.Text, NotZero(ea.Length - ea.Pos))
        End If
        Pos = InStr(Pos + 1, ea.Text, "^")
    Loop



    '_
    Pos = InStr(ea.Text, "_")
    Do While Pos > 0
        If ea.ChrByIndex(Pos + 1) = "(" Then
            s = ea.GetNextBracketContent
            If Not (Left(s, 1) = "{" And right(s, 1) = "}") Then s = "{" & s & "}"
            ea.Text = Left(ea.Text, Pos) & s & right(ea.Text, Len(ea.Text) - ea.Pos + 1)
        Else
            ea.Pos = Pos + 1
            ea.FindHierakiBreak 4, False
            s = ea.SubString(Pos + 1, ea.Pos)
            If Not (Left(s, 1) = "{" And right(s, 1) = "}") Then s = "{" & s & "}"
            ea.Text = Left(ea.Text, Pos) & s & right(ea.Text, NotZero(ea.Length - ea.Pos))
        End If
        Pos = InStr(Pos + 1, ea.Text, "_")
    Loop

    ' /
    Pos = InStr(ea.Text, "/")
    Dim n As String
    Do While Pos > 0
        If ea.ChrByIndex(Pos - 1) = ")" Then
            ea.SetNormalBrackets
            t = ea.GetPrevBracketContent(Pos)
            If ea.IsLetterPos(ea.Pos - 2) Or ea.IsNumber(ea.Pos - 2) Or ea.IsSquareRoot(ea.Pos - 2) Or ea.ChrByIndex(ea.Pos - 2) = "^" Then
                ea.Pos = ea.Pos - 2
                '                pos1 = ea.FindPrevBreak
                pos1 = ea.FindPrevHierakiBreak(4, False)
                t = Mid(ea.Text, pos1, Pos - pos1)
            Else
                pos1 = ea.Pos - 1
            End If
        ElseIf ea.ChrByIndex(Pos - 1) = "}" Then
            ea.SetTuborgBrackets
            t = ea.GetPrevBracketContent(Pos)
            If ea.IsLetterPos(ea.Pos - 2) Or ea.IsNumber(ea.Pos - 2) Or ea.IsSquareRoot(ea.Pos - 2) Or ea.ChrByIndex(ea.Pos - 2) = "^" Then
                ea.Pos = ea.Pos - 2
                '                pos1 = ea.FindPrevBreak
                pos1 = ea.FindPrevHierakiBreak(3, False)
                t = Mid(ea.Text, pos1, Pos - pos1)
            Else
                pos1 = ea.Pos - 1
            End If
        Else
            ea.Pos = Pos - 1
            '            pos1 = ea.FindPrevBreak
            pos1 = ea.FindPrevHierakiBreak(4, False)
            t = Mid(ea.Text, pos1, Pos - pos1)
        End If

        '        If ea.ChrByIndex(pos + 1) = "(" Then
        '            n = ea.GetNextBracketContent(pos)
        '            pos2 = ea.pos
        '        Else
        ea.Pos = Pos + 1
        ''            pos2 = ea.FindNextBreak
        pos2 = ea.FindHierakiBreak(3, False)
        
        n = Mid(ea.Text, Pos + 1, pos2 - Pos)
        If Left(n, 1) = "(" And right(n, 1) = ")" Then
            n = Mid(n, 2, Len(n) - 2)
        End If
        pos2 = pos2 + 1
        '        End If
        ea.Text = Left(ea.Text, pos1 - 1) & "\frac{" & t & "}{" & n & "}" & right(ea.Text, NotZero(Len(ea.Text) - pos2 + 1))
        Pos = InStr(ea.Text, "/")
    Loop


    ' kubikrod
    '    ea.text = Replace(ea.text, VBA.ChrW(8731) & "(", VBA.ChrW(8730) & "(3&")
    '    ea.text = Replace(ea.text, VBA.ChrW(8731), VBA.ChrW(8730) & "(3&")

    ' 3 og 4 rod
    For rod = 3 To 4
        Do
            Pos = InStr(ea.Text, VBA.ChrW(8728 + rod))
            If Pos > 0 Then
                ea.Pos = Pos + 1
                If Mid(ea.Text, Pos + 1, 1) <> "(" Then
                    ea.SetTuborgBrackets
                    ea.InsertUnderstoodBracketPair
                    ea.Text = Replace(ea.Text, VBA.ChrW(8728 + rod), "\sqrt[" & rod & "]", 1, 1)
                Else
                    ea.SetNormalBrackets
                    s = ea.GetNextBracketContent()
                    ea.Text = Left(ea.Text, Pos - 1) & "\sqrt[" & rod & "]{" & s & "}" & right(ea.Text, NotZero(ea.Length - ea.Pos + 1))
                End If
                ea.Pos = Pos
            End If
        Loop While Pos > 0
    Next

    'kvadratrod
    Do
        Pos = InStr(ea.Text, VBA.ChrW(8730))
        If Pos > 0 Then
            If Mid(ea.Text, Pos + 1, 1) <> "(" Then
                ea.Text = Replace(ea.Text, VBA.ChrW(8730), "\sqrt", 1, 1)
                Pos = Pos + 5
                ea.Pos = Pos
                ea.SetTuborgBrackets
                ea.InsertUnderstoodBracketPair
            Else
                ea.Pos = Pos
                Arr = Split(ea.GetNextBracketContent, "&")
                pos2 = ea.Pos
                If UBound(Arr) = 0 Then
'                    ea.text = Replace(ea.text, VBA.ChrW(8730), "\sqrt ", 1, 1)
                    ea.Text = Left(ea.Text, Pos - 1) & "\sqrt{" & Arr(0) & "}" & right(ea.Text, NotZero(ea.Length - ea.Pos + 1))
                ElseIf UBound(Arr) = 1 Then
                    ea.Text = Left(ea.Text, Pos - 1) & "\sqrt[" & Arr(0) & "]{" & Arr(1) & "}" & right(ea.Text, NotZero(ea.Length - ea.Pos + 1))
                End If
            End If
        End If
    Loop While Pos > 0

    ' matricer
    ea.Text = Replace(ea.Text, VBA.ChrW(9608), VBA.ChrW(9632))
    Pos = InStr(ea.Text, VBA.ChrW(9632))
    Do While Pos > 0
        ea.SetNormalBrackets
        s = ea.GetNextBracketContent(Pos)
        ea.Text = Left(ea.Text, Pos - 1) & "\begin{matrix}" & Replace(s, "@", "\\") & "\end{matrix}" & right(ea.Text, NotZero(ea.Length - Pos - Len(s) - 2))
        Pos = InStr(ea.Text, VBA.ChrW(9632))
    Loop

    ' funktioner

    ea.ReplaceVar "sin", "\sin "    '
    ea.ReplaceVar "cos", "\cos "
    ea.ReplaceVar "tan", "\tan "
    ea.ReplaceVar "log", "\log "
    ea.ReplaceVar "ln", "\ln "
    '    ea.text = Replace(ea.text, "sin" & VBA.ChrW(8289), "\sin ") '
    '    ea.text = Replace(ea.text, "cos" & VBA.ChrW(8289), "\cos ") '
    '    ea.text = Replace(ea.text, "tan" & VBA.ChrW(8289), "\tan ") '
    '    ea.text = Replace(ea.text, "log" & VBA.ChrW(8289), "\log ") '
    '    ea.text = Replace(ea.text, "ln" & VBA.ChrW(8289), "\ln ") '
    '    ea.text = Replace(ea.text, "sin^{-1}", "\sin^{-1} ") '

    ' enheder ikke i kursiv \,\text{cm}
    If LatexUnits Then
    ea.Pos = 1
    s = ea.GetNextVar()
    Do
        ea.Pos = ea.Pos + 1
        If ea.IsUnit(s) Then
            If ea.IsNumber(ea.Pos - Len(s) - 1) Or ea.ChrByIndex(ea.Pos - Len(s) - 1) = " " Then
                t = "\,\mathrm{" & s & "}"
            Else
                t = "\mathrm{" & s & "}"
            End If
            ea.ReplaceVar s, t, ea.Pos - Len(s), 1
            ea.Pos = ea.Pos + Len(t) - Len(s)
        End If
        ea.Text = Replace(ea.Text, VBA.ChrW(8289), "")    ' funktionstegn
        s = ea.GetNextVar()
    Loop While s <> ""
    End If

    'gr*ae*ske bogstaver
    ea.Text = Replace(ea.Text, VBA.ChrW(916), "\Delta ")
    ea.Text = Replace(ea.Text, VBA.ChrW(948), "\delta ")
    ea.Text = Replace(ea.Text, VBA.ChrW(945), "\alpha ")
    ea.Text = Replace(ea.Text, VBA.ChrW(946), "\beta ")
    ea.Text = Replace(ea.Text, VBA.ChrW(947), "\gammaLB ")
    ea.Text = Replace(ea.Text, VBA.ChrW(952), "\theta ")
    ea.Text = Replace(ea.Text, VBA.ChrW(920), "\Theta ")
    ea.Text = Replace(ea.Text, VBA.ChrW(955), "\lambda ")
    ea.Text = Replace(ea.Text, VBA.ChrW(923), "\Lambda ")
    ea.Text = Replace(ea.Text, VBA.ChrW(956), "\mu ")
    ea.Text = Replace(ea.Text, VBA.ChrW(961), "\rho ")
    ea.Text = Replace(ea.Text, VBA.ChrW(963), "\sigma ")
    ea.Text = Replace(ea.Text, VBA.ChrW(931), "\Sigma ")
    ea.Text = Replace(ea.Text, VBA.ChrW(981), "\phi ")
    ea.Text = Replace(ea.Text, VBA.ChrW(934), "\Phi ")
    ea.Text = Replace(ea.Text, VBA.ChrW(949), "\varepsilon ")
    ea.Text = Replace(ea.Text, VBA.ChrW(1013), "\epsilon ")
    ea.Text = Replace(ea.Text, VBA.ChrW(968), "\psi ")
    ea.Text = Replace(ea.Text, VBA.ChrW(936), "\Psi ")
    ea.Text = Replace(ea.Text, VBA.ChrW(926), "\Xi ")
    ea.Text = Replace(ea.Text, VBA.ChrW(958), "\xi ")
    ea.Text = Replace(ea.Text, VBA.ChrW(935), "\Chi ")
    ea.Text = Replace(ea.Text, VBA.ChrW(967), "\chi ")
    ea.Text = Replace(ea.Text, VBA.ChrW(928), "\Pi ")
    ea.Text = Replace(ea.Text, VBA.ChrW(964), "\tau ")
    ea.Text = Replace(ea.Text, VBA.ChrW(957), "\greek-nu ")
    ea.Text = Replace(ea.Text, VBA.ChrW(954), "\kappa ")
    ea.Text = Replace(ea.Text, VBA.ChrW(951), "\eta ")
    ea.Text = Replace(ea.Text, VBA.ChrW(950), "\zeta ")
    ea.Text = Replace(ea.Text, VBA.ChrW(960), "\pi ")

    ea.Text = Replace(ea.Text, VBA.ChrW(8592), "\leftarrow ")    '
    ea.Text = Replace(ea.Text, VBA.ChrW(8594), "\rightarrow ")    '
    ea.Text = Replace(ea.Text, VBA.ChrW(8652), "\rightleftharpoons ")    ' egentlig \leftrightharpoons   but not supported by default Latex
    ea.Text = Replace(ea.Text, VBA.ChrW(8651), "\rightleftharpoons ")    '
    ea.Text = Replace(ea.Text, VBA.ChrW(8592), "\leftarrow ")    '
    ea.Text = Replace(ea.Text, VBA.ChrW(10236), "\longmapsto ")    '
    ea.Text = Replace(ea.Text, VBA.ChrW(8640), "\rightharpoonup ")    '
    ea.Text = Replace(ea.Text, VBA.ChrW(8641), "\rightharpoondown ")    '
    ea.Text = Replace(ea.Text, VBA.ChrW(183), "\cdot ")    ' prik erstattes med gange
    ea.Text = Replace(ea.Text, VBA.ChrW(215), "\times ")    ' x gangetegn
    ea.Text = Replace(ea.Text, " \", "\") ' i Moodle bliver mellemrum lavet. Derfor skal de helst fjernes. mellemrum f*oe*r \ er ihvertfald ikke n*oe*dvendigt

    ' parenteser
    ea.Text = Replace(ea.Text, "(", "\left(")    '
    ea.Text = Replace(ea.Text, ")", "\right)")    '
    
    'kantede parenteser ' ikke s*aa* simpelt fordi de indg*aa*r i latex syntaks
    ea.SetSquareBrackets
    pos1 = 1
    Do
        Pos = InStr(pos1, ea.Text, "[")
        If Pos > 0 And ea.ChrByIndex(Pos - 1) <> "}" Then
            pos1 = Pos + 6
            pos2 = Pos
            ea.GetNextBracketContent Pos - 1
            pos3 = ea.Pos
            If Not ea.ChrByIndex(pos3) = "{" Then
                ea.Text = Left(ea.Text, pos2 - 1) & "\left" & right(ea.Text, Len(ea.Text) - pos2 + 1)
                ea.Text = Left(ea.Text, pos3 + 3) & "\right" & right(ea.Text, Len(ea.Text) - pos3 - 3)
            End If
        Else
            pos1 = Pos + 1
        End If
    Loop While Pos > 0

    ConvertToLatex = ea.Text
End Function
Function ConvertTrigPowLatex(Trig As String, Text As String) As String
' omdanner f.eks. sin^2(x) -> sin^2(x)
Dim ea As ExpressionAnalyser
Dim pow As String
Dim pos2 As Integer, Pos As Integer
Pos = 1
Do
    Pos = InStr(Pos, Text, Trig & "^")
    If Pos > 0 Then
        Pos = Pos + Len(Trig) + 1
        pos2 = InStr(Pos, Text, VBA.ChrW(8289))
        If pos2 = 0 Then
            Set ea = New ExpressionAnalyser
            ea.SetNormalBrackets
            ea.Text = Text
            ea.Pos = Pos
            pos2 = ea.FindEndOfBracket
        End If
        pow = Mid(Text, Pos, pos2 - Pos)
        If pow = "(-1)" Then pow = "-1"
        Text = Left(Text, Pos - 1) & "{" & pow & "}" & right(Text, Len(Text) - pos2)
'        InsertTextAt "^" & pow, FindEndOfBracket
        
    End If
Loop While Pos > 0
ConvertTrigPowLatex = Text
End Function
Function ConvertAccent(Text As String, Accent As String, aString, bString) As String
' ConvertAccent(ea.text,VBA.ChrW(8407), "\overrightarrow{","}")
Dim ea As New ExpressionAnalyser, Pos As Integer, pos1 As Integer, s As String
    ea.Text = Text
    Pos = InStr(ea.Text, Accent)
    Do While Pos > 0
        If ea.ChrByIndex(Pos + 1) = " " Then ea.RemoveChar (Pos + 1) ' der kan v*ae*re mellemrum efter pilen som kan forstyrre
        If Mid(ea.Text, Pos - 1, 1) = ")" Then
            ea.SetNormalBrackets
            s = ea.GetPrevBracketContent(Pos)
            ea.Text = Left(ea.Text, Pos - Len(s) - 3) & aString & s & bString & right(ea.Text, Len(ea.Text) - Pos)
        ElseIf Mid(ea.Text, Pos - 1, 1) = "}" Then
            ea.SetTuborgBrackets
            s = ea.GetPrevBracketContent(Pos)
            pos1 = ea.Pos - 2
            Do While ea.IsLetterPos(pos1, False)
                pos1 = pos1 - 1
            Loop
            If ea.ChrByIndex(pos1) = "\" Then
                ea.Text = Left(ea.Text, pos1 - 1) & aString & Mid(ea.Text, pos1, Pos - pos1) & bString & right(ea.Text, Len(ea.Text) - Pos - Len(Accent) + 1)
            End If
        Else
            ea.Text = Left(ea.Text, Pos - 2) & aString & Mid(ea.Text, Pos - 1, 1) & bString & right(ea.Text, Len(ea.Text) - Pos - Len(Accent) + 1)
        End If
        Pos = InStr(Pos + 1, ea.Text, Accent)
    Loop
    ConvertAccent = ea.Text
End Function
Function ConvertPreAccent(Text As String, Accent As String, aString, bString) As String
    Dim ea As New ExpressionAnalyser
    Dim Pos As Integer, s As String
    ea.Text = Text
    Do
        Pos = InStr(ea.Text, Accent)
        If Pos > 0 Then
            If ea.ChrByIndex(Pos + 1) = "(" Then
                ea.SetNormalBrackets
                s = ea.GetNextBracketContent(Pos)
                ea.Text = Left(ea.Text, Pos - 1) & aString & s & bString & right(ea.Text, Len(ea.Text) - ea.Pos + 1)
            Else
                s = ea.GetNextVar(Pos)
                ea.Text = Left(ea.Text, Pos - 1) & aString & s & bString & right(ea.Text, Len(ea.Text) - ea.Pos)
            End If
        End If
    Loop While Pos > 0
    ConvertPreAccent = ea.Text
End Function
Function ConvertToLatexMaxima(Optional Text As String) As String
' converts to Latex using Maxima
    If Text <> "" Then Kommando = Text
    Dim pkommando As String
    
    pkommando = Kommando
    Kommando = omax.CodeForMaxima(Kommando)
    Kommando = Replace(Kommando, "'integrate", "integrate")
    Kommando = Replace(Kommando, "NIntegrate", "integrate")
    Kommando = Replace(Kommando, "Integrate", "integrate")
    Kommando = Replace(Kommando, "integrate", "'integrate")
    Kommando = Replace(Kommando, "'diff", "diff")
    Kommando = Replace(Kommando, "diff", "'diff")
    
    Kommando = "tex1('(" & Kommando & "))"
    
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"

    MaximaInputStrengSec = ""

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec

    Call RunMaxima(0, True, False)
    MaximaOutput = TrimB(MaximaOutput, """")
    MaximaOutput = Replace(MaximaOutput, "\\", "\")
    If StopNow Then GoTo Slut

    '    MaximaOutput = Replace(MaximaOutput, "%i", "i")
    ConvertToLatexMaxima = MaximaOutput
Slut:


End Function
Function ConvertLatexToWord(Text As String) As String
    Dim ea As New ExpressionAnalyser
    Dim p As Integer
    Dim t As String, n As String

    Text = Replace(Text, VBA.ChrW(11), "")    ' shift-enter
    Text = Replace(Text, vbCrLf, "")    ' enter
    Text = Replace(Text, VBA.ChrW(10), "")    ' shift-enter
    Text = Replace(Text, VBA.ChrW(13), "")    ' shift-enter
    Text = Replace(Text, "\infty", VBA.ChrW(8734))    '
    Text = Replace(Text, "\cdot", VBA.ChrW(183))    'prik gange

    'gr*ae*ske bogstaver
    Text = Replace(Text, "\Delta", VBA.ChrW(916))
    Text = Replace(Text, "\delta", VBA.ChrW(948))
    Text = Replace(Text, "\alpha", VBA.ChrW(945))
    Text = Replace(Text, "\beta", VBA.ChrW(946))
    Text = Replace(Text, "\gamma", VBA.ChrW(915))
    Text = Replace(Text, "\theta", VBA.ChrW(952))
    Text = Replace(Text, "\Theta", VBA.ChrW(920))
    Text = Replace(Text, "\lambda", VBA.ChrW(955))
    Text = Replace(Text, "\Lambda", VBA.ChrW(923))
    Text = Replace(Text, "\rho", VBA.ChrW(961))
    Text = Replace(Text, "\phi", VBA.ChrW(981))
    Text = Replace(Text, "\Phi", VBA.ChrW(934))
    Text = Replace(Text, "\varepsilon", VBA.ChrW(949))
    Text = Replace(Text, "\epsilon", VBA.ChrW(1013))
    Text = Replace(Text, "\psi", VBA.ChrW(968))
    Text = Replace(Text, "\Psi", VBA.ChrW(936))
    Text = Replace(Text, "\sigma", VBA.ChrW(963))
    Text = Replace(Text, "\Sigma", VBA.ChrW(931))
    Text = Replace(Text, "\mu", VBA.ChrW(956))
    Text = Replace(Text, "\Ohm", VBA.ChrW(937))
    Text = Replace(Text, "\Omega", VBA.ChrW(937))
    Text = Replace(Text, "omega", VBA.ChrW(969))
    Text = Replace(Text, "\Xi", VBA.ChrW(926))
    Text = Replace(Text, "\xi", VBA.ChrW(958))
    Text = Replace(Text, "\Chi", VBA.ChrW(935))
    Text = Replace(Text, "\chi", VBA.ChrW(967))
    Text = Replace(Text, "\tau", VBA.ChrW(964))
    Text = Replace(Text, "\Pi", VBA.ChrW(928))
    '    text = Replace(text, "\greek-nu", VBA.ChrW(957))
    Text = Replace(Text, "\kappa", VBA.ChrW(954))
    Text = Replace(Text, "\zeta", VBA.ChrW(950))
    Text = Replace(Text, "\eta", VBA.ChrW(951))    ' skal v*ae*re sidst da eta indg*aa*r i andre
    Text = Replace(Text, "\inc", VBA.ChrW(8710))  ' specielt delta increment
    '    text = Replace(text, "degC", VBA.ChrW(8451))   ' specielt oC tegn
    '    text = Replace(text, "Symangle", VBA.ChrW(8736))  '

    ea.SetTuborgBrackets
    ea.Text = Text
    p = InStr(ea.Text, "\frac")
    Do While p > 0
        t = ea.GetNextBracketContent
        n = ea.GetNextBracketContent
        ea.Text = Left(ea.Text, p - 1) & "(" & t & ")/(" & n & ")" & right(ea.Text, ea.Length - ea.Pos + 1)
        p = InStr(ea.Text, "\frac")
    Loop

    ConvertLatexToWord = ea.Text

End Function
Sub ReadSelection()
    ' inds*ae*tter det markerede som kommando
    Dim mtext As String, copytext As String
    Dim antal As Integer
    Dim sstart As Long
    Dim sslut As Long
    Dim j As Integer, antalgenstart As Integer
    On Error GoTo Slut

    ' hvis bold giver det buildup problemer ' i v. 1.24.1 forsøges det håndteret uden at ændre i originalen
'        For j = 1 To Selection.OMaths.Count
'            Selection.OMaths(j).Range.Font.Bold = False
'        Next
    sstart = Selection.start
    sslut = Selection.End
    antal = Selection.OMaths.Count
    Dim r As Range, sr As Range
    '    With tempDoc
    '        .Content.Font.Name = "Tahoma"
    '        .SaveAs FileName:="Sample.doc"
    '    End With

    If antal = 0 Then
        '    kommando = Selection.text
        Selection.Collapse (wdCollapseEnd)
    ElseIf antal = 1 Then
        '    If Len(Selection.text) < 3 Then
        '        Selection.OMaths(1).Range.Select
        '    End If
        '    If Selection.OMaths(1).Range.End < Selection.Range.End Then
        '        Selection.OMaths(1).Range.Select
        '    End If
        Selection.OMaths(1).Range.Select    ' ovenst*aa*ende var ikke nok, men hvorfor var det der?

        antalgenstart = 0
genstart:
        Kommando = ReadEquationCopy()
        If Kommando = "" And antalgenstart = 0 Then    ' fors*oe*g p*aa* at rette sj*ae*ldent problem hvor alt l*ae*ses som "" m*aa*ske skyldes det tempdoc
            '            tempDoc.Close (False)
            '            Set tempDoc = Nothing
            LukTempDoc
            OpretTempdoc
            antalgenstart = antalgenstart + 1
            GoTo genstart
        ElseIf Kommando = "" And antalgenstart = 1 Then
            RestartWordMat
            antalgenstart = antalgenstart + 1
            GoTo genstart
        ElseIf Kommando = "" Then
            MsgBox "WordMat cant read the command. This can be caused by error which can be corrected by restarting Word (Close all Word documents).", vbOKOnly, Sprog.Error
        End If

        If InStr(Kommando, VBA.ChrW(8743)) Then
            Kommandoer = Split(Kommando, VBA.ChrW(8743))
            '        For j = UBound(Kommandoer) To 0 Step -1
            '            Kommandoer(j + 1) = Kommandoer(j)
            '        Next
        End If

        Selection.Collapse (wdCollapseEnd)
    ElseIf antal > 1 Then
        ReDim Kommandoer(0 To 100)
        '    For j = 1 To antal
        '        Selection.OMaths(j).Range.Select ' ovenst*aa*ende var ikke nok, men hvorfor var det der?
        '        Kommandoer(j - 1) = Replace(ReadEquationCopy(), Chr(11), "") ' fjern shift-enter
        '    Next

        '    GoTo hop ' fors*oe*g

        Set sr = Selection.Range
        Selection.Copy
        '    Set at = NormalTemplate.AutoTextEntries.Add(Name:="WordMat", Range:=Selection.Range)
        '    Set at = ActiveDocument.AttachedTemplate.AutoTextEntries.Add(name:="WordMat", Range:=Selection.Range)


        '    Set tempDoc = Documents.Add(, , , False)
        Set r = tempDoc.Range
        '    at.Insert r, True
        '    r.text = copytext
        '    r.OMaths.Add Range:=r
        r.Paste
        With r
            .OMaths.BuildUp
            .OMaths.Linearize
        End With

        Dim i As Integer
        Dim mathcol As OMaths
        Set mathcol = r.OMaths
        For j = 0 To antal - 1
            mathcol(j + 1).ConvertToNormalText
            Kommandoer(j) = Replace(mathcol(j + 1).Range.Text, ChrW(11), "")    ' fjern shift-enter
        Next
        

hop:
        ReDim Preserve Kommandoer(0 To antal - 1)
        sr.Select
    End If

Slut:
    Selection.start = sstart
    Selection.End = sslut

End Sub
Function ReadEquation(Optional ir As Range) As String
    ' Overs*ae*tter selection der er omath til streng
    ' kan g*oe*res hurtigere hvis der kopieres indenfor samme document, men virker ikke 100%
    ' autotextentries kan kopiere istedet for clipboard, og er m*aa*ske hurtigere, men det er meget tvivlsomt.
    ' har dog oplevet at det giver problemer. Ligningssystem tan(30)=h/x og tan(25)=h/(x+34) p*aa* l*oe*sningen beregnes i enden. Det g*aa*r helt galt, men virker n*aa*r autotextentries ikke anvendes

    Dim sr As Range, r As Range
    If ir Is Nothing Then
        Set ir = Selection.Range
    End If

    Set sr = Selection.Range
    Set at = ActiveDocument.AttachedTemplate.AutoTextEntries.Add(Name:="WordMat", Range:=ir)

    Set r = tempDoc.Range
    at.Insert r, True

    With r
        .Font.Bold = False
        .OMaths.BuildUp
        .OMaths.Linearize
        .OMaths(1).ConvertToNormalText
    End With

    ReadEquation = r.Text
    r.Delete    ' dette afhj*ae*lper muligvis at word bliver langsom

    sr.Select

End Function
Function ReadEquationCopy(Optional ir As Range) As String
' bruger selection.copy som nok er langsommere end autotextentries og fylder clipboard
' Overs*ae*tter selection der er omath til streng
    Dim sr As Range, r As Range
    '    Dim t As Variant
    '    t = Timer
    If ir Is Nothing Then

    End If

    '    copytext = Selection.Range.text ' fors*oe*g med ikke at klippe og kopiere problemet er at man er n*oe*dt til at konvertere til text for at l*ae*se ordentlig
    Set sr = Selection.Range

    Selection.Copy

    Set r = tempDoc.Range
    
    With r
        .Paste
        .Font.Bold = False
        .OMaths.BuildUp
        .OMaths.Linearize
        .OMaths(1).ConvertToNormalText
    End With

    ReadEquationCopy = r.Text
    sr.Select

    r.Delete
    '    ClearClipboard

    '    tempDoc.Range.Delete ' dette afhj*ae*lper muligvis at word bliver langsom
    '    MsgBox Timer - t
End Function
Function ReadEquation2(ir As Range) As String
    Dim sr As Range
    Dim r As Range
    Set sr = Selection.Range
    ir.Copy
    Set r = tempDoc.Range
'    at.Insert r, True   ' fjernet 19/2 den er vel erstattet af paste under with nedenfor?

    With r
        .Paste
        .Font.Bold = False
        .OMaths.BuildUp
        .OMaths.Linearize
        .OMaths(1).ConvertToNormalText
    End With

    ReadEquation2 = r.Text
    sr.Select

    tempDoc.Range.Delete    ' dette afhj*ae*lper muligvis at word bliver langsom

End Function
Function ReadFromWord(Expr As String) As String
' tager udtryk som streng. Indsætter i word, bygger op og returnerer

    Dim r As Range, oldr As Range

    Set oldr = Selection.Range

    Set r = tempDoc.Range
    r.Delete
    Set r = tempDoc.Range
    With r
        .OMaths.Add r
        .OMaths(1).Range.Text = Expr
        .OMaths(1).Range.Font.Bold = False
        .OMaths.BuildUp
        .OMaths.Linearize
        .OMaths(1).ConvertToNormalText
    End With

    ReadFromWord = r.Text

    r.Delete
    
    oldr.Select
End Function
Sub ReadSelectionOld()
' inds*ae*tter det markerede som kommando
    Dim mtext As String
    Dim antal As Integer
    Dim sstart As Integer
    Dim sslut As Integer
    Dim j As Integer

    ' hvis bold giver det buildup problemer
    For j = 1 To Selection.OMaths.Count
        Selection.OMaths(j).Range.Font.Bold = False
    Next

    sstart = Selection.start
    sslut = Selection.End
    antal = Selection.OMaths.Count

    If antal = 0 Then
        If Len(Selection.Text) < 2 Then
            Kommando = ""
            MsgBox "Marker det der skal evalueres af Maxima.", vbOKOnly, "V*ae*lg"
            Exit Sub
        End If
        Kommando = Selection.Text
        Selection.Collapse (wdCollapseEnd)
    ElseIf antal = 1 Then
        If Len(Selection.Text) < 3 Then
            Selection.OMaths(1).Range.Select
        End If
        If Selection.OMaths(1).Range.End < Selection.Range.End Then
            Selection.OMaths(1).Range.Select
        End If

        Selection.OMaths.BuildUp
        Selection.OMaths.Linearize
        Selection.OMaths(1).ConvertToNormalText
        Kommando = Selection.Text
        Selection.Text = Replace(Selection.Text, VBA.ChrW(8727), MaximaGangeTegn)    'fjern *
        Selection.Text = Replace(Selection.Text, "*", MaximaGangeTegn)    'fjern *
        If DecSeparator = "," Then
            Selection.Text = Replace(Selection.Text, ".", ",")    'N*aa*r komma er decseparator kan . *ae*ndres til komma
        End If
        Selection.OMaths(1).ConvertToMathText
        Selection.OMaths(1).Range.Select
        Selection.OMaths.BuildUp
        Selection.Collapse (wdCollapseEnd)
    ElseIf antal > 1 Then
        ReDim Kommandoer(1 To 100)
        Selection.OMaths.Linearize
        Dim i As Integer
        Dim mathcol As OMaths
        Set mathcol = Selection.OMaths
        For i = 1 To antal
            mathcol(i).ConvertToNormalText
            Kommandoer(i) = Replace(mathcol(i).Range.Text, ChrW(11), "")
            mathcol(i).ConvertToMathText
            mathcol(i).BuildUp
        Next
        ReDim Preserve Kommandoer(1 To antal)
        '    mathcol(mathcol.Count).Range.Select
        '    Selection.OMaths(1).Range.Select
        '    Selection.Collapse (wdCollapseEnd)
        '    Selection.OMaths(1).Range.Select
        '    Selection.Collapse (wdCollapseEnd)
    End If
    Selection.start = sstart
    Selection.End = sslut

    'tempDoc.Close

    'Selection.Collapse (wdCollapseEnd)

End Sub
Sub GoToEndOfSelectedMaths2()
    Dim mc As OMaths
    Set mc = Selection.OMaths
    On Error Resume Next
    mc(mc.Count).ParentOMath.Range.Select    ' parentomath n*oe*dvendig hvis cursor st*aa*r i lysebl*aa* firkant
    On Error GoTo Slut
    '        mc(mc.Count).Range.Select  ' virker med word 2010, parentomath giver tilgeng*ae*ld problemer. Hmm problem med valgt del af udtryk og reducer
Slut:
    Selection.Collapse wdCollapseEnd
    Dim r As Range
    Set r = Selection.Range
    r.MoveStart wdCharacter, -1
    If r.Text = VBA.ChrW(11) Then    ' hvis der er shift-enter i slutningen erstattes med alm. retur
        r.Text = VBA.ChrW(13)
    End If
End Sub
Sub GoToEndOfSelectedMaths()
    Dim mc As OMaths
    Dim i As Integer
    
    If Selection.Tables.Count > 0 Then 'nummereret ligning
        If Selection.Tables(1).Columns.Count = 3 Then
            If Selection.Tables(1).Cell(1, 3).Range.Fields.Count > 0 Then
                Selection.Tables(1).Select
                Selection.Collapse wdCollapseEnd
'                Selection.MoveRight wdCharacter, 1
                Exit Sub
            End If
        End If
    End If
    
    Selection.Collapse wdCollapseEnd
    Set mc = Selection.OMaths
    If mc.Count > 0 Then
        On Error Resume Next
        mc(mc.Count).ParentOMath.Range.Select
        On Error GoTo Slut
        mc(mc.Count).Range.Select  ' virker med word 2010, parentomath giver tilgeng*ae*ld problemer. Hmm problem med valgt del af udtryk og reducer
    Else
        i = 0
        Do While Selection.OMaths.Count = 0 And i < 100
            Selection.MoveLeft wdCharacter, 1
            i = i + 1
        Loop
    End If
Slut:
    On Error Resume Next
    Selection.Collapse wdCollapseEnd
    Dim r As Range
    Set r = Selection.Range
    r.MoveStart wdCharacter, -1
    If r.Text = VBA.ChrW(11) Then    ' hvis der er shift-enter i slutningen erstattes med alm. retur
        r.Text = VBA.ChrW(13)
    End If
End Sub
Sub InsertMaximaOutput()
' inds*ae*tter maxima output i word document
    Dim mo As Range, IsBold As Boolean, s As String

    Selection.Range.Font.Bold = False
    If MaximaKommando Then
        Set mo = Selection.OMaths.Add(Selection.Range)
        Selection.TypeText Text:=Kommando
        mo.OMaths.BuildUp
        '        GoToEndOfSelectedMaths
        Selection.TypeParagraph
        '        Selection.TypeText Kommando
    End If
    '    Set mo = Selection.Range
    '    mo.text = MaximaOutput
    '    Set mo = Selection.OMaths.Add(mo)
    '    mo.OMaths(1).BuildUp

    If Len(MaximaOutput) > 500 Then
        MaximaOutput = Left(MaximaOutput, 500) + Sprog.A(119)
    End If
    Set mo = Selection.OMaths.Add(Selection.Range)
    If mo.OMaths(1).Range.Bold Then ' Der er en fejl i Word build virker ikke korrekt med fed skrift. Fx løsningen af 2000=3000*0,85?x
        mo.OMaths(1).Range.Bold = False
        IsBold = True
    End If
    Selection.TypeText MaximaOutput
    mo.OMaths.BuildUp
    If IsBold Then
        mo.OMaths(1).Range.Bold = True
    End If

    InsertAssumptions

    If Len(KommentarOutput) > 3 Then
        s = TranslateReplaceComment(KommentarOutput)
        If Len(s) > 3 Then
            Selection.TypeParagraph
            Selection.InsertAfter s
            Selection.Collapse (wdCollapseEnd)
        End If
    End If

End Sub
Function TranslateReplaceComment(comm As String) As String

    Dim outtext As String

    ' trim linefeeds in front
    'Do While Left(comm, 1) = vbLf
    '    comm = Right(comm, Len(comm) - 1)
    'Loop
    ' trim linefeeds behind
    'Do While Right(comm, 1) = vbLf
    '    comm = Left(comm, Len(comm) - 1)
    'Loop

    If InStr(comm, "Solution unsure") > 0 Then
        outtext = outtext & Sprog.A(101)
    End If

    If InStr(comm, "solve: using arc-trig functions to get a solution.") > 0 Then
        outtext = outtext & Sprog.A(102)
    End If

    If InStr(comm, "Some solutions will be lost.") > 0 Then
        outtext = outtext & Sprog.A(103)
    End If

    If InStr(comm, "A number was found where a variable was expected -`solve'") > 0 Then
        outtext = outtext & ""
    End If

    If InStr(comm, "A number was found where a variable was expected -`solve'") > 0 Then
        outtext = outtext & Sprog.ErrorGeneral
    End If

    If InStr(comm, "Division by 0") > 0 Then
        outtext = outtext & Sprog.A(104)
    End If

    If InStr(comm, "syntax error") > 0 Then
        outtext = outtext & Sprog.A(105)
    End If

    If InStr(comm, "infix operator") > 0 Then
        outtext = outtext & Sprog.A(106)
    End If

    If InStr(comm, "solving system of equations") > 0 Then
        outtext = outtext & ""
    End If

    If InStr(comm, "Warning too few variables, only") > 0 Then
        outtext = outtext & Sprog.A(107)
    End If



    'comm = Replace(comm, "solve: using arc-trig functions to get a solution.", "Da det er en trigonometrisk ligning er der uendeligt mange l*oe*sninger. Kun en af disse l*oe*sninger er angivet.")
    'comm = Replace(comm, "Some solutions will be lost.", " Alle l*oe*sninger er ikke fundet.")
    'comm = Replace(comm, "A number was found where a variable was expected -`solve'", "Du har bedt om at l*oe*se en ligning for en variabel der er defineret til en v*ae*rdi. Lav en 'slet def:'  kommando f*oe*r du l*oe*ser for denne variabel.")
    'comm = Replace(comm, " -- an error. To debug this try: debugmode(true);", "Der skete en fejl.")
    'comm = Replace(comm, "Division by 0", "Fejl: Division med nul.")

    TranslateReplaceComment = outtext
End Function
Sub InsertAssumptions()
    Dim gemfontsize As Integer
    Dim gemitalic As Boolean
    Dim gemfontcolor As Integer
    Dim gemsb As Integer
    Dim gemsa As Integer
    Dim mo As Range
    Dim assum As String

    If Not (ShowAssum) Then GoTo Slut

    If InStr(KommentarOutput, "Assumptions:") > 0 Then
        Dim Pos As Integer
        Dim ea As New ExpressionAnalyser
        Pos = InStr(KommentarOutput, "Assumptions:")
        ea.Text = KommentarOutput
        ea.Pos = Pos
        assum = Trim(omax.ConvertToWordSymbols(Replace(ea.GetNextBracketContent, ",", " , ")))
    End If
    If Len(assum) < 2 Then GoTo Slut

    gemfontsize = Selection.Font.Size
    gemitalic = Selection.Font.Italic
    gemfontcolor = Selection.Font.ColorIndex
    gemsb = Selection.ParagraphFormat.SpaceBefore
    gemsa = Selection.ParagraphFormat.SpaceAfter

    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 2
        .SpaceAfterAuto = False
        '        .LineUnitBefore = 0
        '        .LineUnitAfter = 0
        .Alignment = wdAlignParagraphCenter
    End With
    Selection.Font.Size = 8
    Selection.Font.ColorIndex = wdGray50
    Selection.Font.Italic = True

    Selection.TypeParagraph
    Selection.TypeText "L*oe*sningsbetingelser:  "
    '            Selection.TypeText text:=vbTab
    Set mo = Selection.OMaths.Add(Selection.Range)
    Selection.TypeText assum
    mo.OMaths.BuildUp

    Selection.TypeParagraph
    Selection.Font.Size = gemfontsize
    Selection.Font.Italic = gemitalic
    Selection.Font.ColorIndex = gemfontcolor
    With Selection.ParagraphFormat
        .SpaceBefore = gemsb
        '        .SpaceBeforeAuto = False
        .SpaceAfter = gemsa
        '        .SpaceAfterAuto = False
    End With

    GoTo Slut
Fejl:
Slut:

End Sub

Function FindMaximaOutput(Text As String) As String
' plukker output2 fra hele maxima output
    Dim opos, ipos As Integer
    Dim text2 As String
    On Error GoTo Fejl

    opos = InStr(Text, "(%o2)")
    ipos = InStr(Text, "(%i3)")
    'MsgBox text
    If opos < 1 Then GoTo Fejl

    text2 = Trim(Mid(Text, opos + 5, ipos - opos - 5))
    text2 = Replace(text2, "[", "")
    text2 = Replace(text2, "]", "")
    'text2 = Mid(text2, 2, Len(text2) - 3) ' fjern kantede paranteser

    FindMaximaOutput = text2

    GoTo Slut
Fejl:
    opos = InStr(Text, "(%i2)")
    text2 = right(Text, Len(Text) - opos - 4)
    text2 = Replace(text2, "(%i2)", "")
    text2 = Sprog.A(108) & vbCrLf & vbCrLf & text2
    MsgBox text2, vbOKOnly, Sprog.A(109)
    FindMaximaOutput = ""
Slut:

End Function
Sub FindLastMaximaOutput(Optional ConvertToWord As Boolean = True, Optional Mcommand As String = "")
' plukker sidste output fra hele maxima output
Dim same As Boolean
    GenerateOutputArray (Mcommand)
    Dim nr As Integer
    nr = UBound(Minput)
    If MaximaUnits Then
        If nr > 1 Then
            If moutput(nr) = moutput(nr - 1) Then
                same = True
            End If
        End If
        If moutput(nr) <> "true" And moutput(nr) <> "false" And moutput(nr) <> "unknown" Then
            moutput(nr) = ConvertMaximaUnits(moutput(nr), False)
        End If
        If same Then ' speeds up for solve since only last output used
                moutput(nr - 1) = moutput(nr)
        ElseIf nr > 1 Then
                moutput(nr - 1) = ConvertMaximaUnits(moutput(nr - 1), False)
        End If
    End If
    If ConvertToWord Then
        MaximaOutput = ConvertToWordSymbols(moutput(nr))
    Else
        MaximaOutput = moutput(nr)
    End If
    KommentarOutput = Minput(nr)
    If InStr(KommentarOutput, "rat: replaced") Then
        KommentarOutput = ""
    End If

    'If MaximaOutput = "" Then
    '    KommentarOutput = "Der kom f*oe*lgende fejlmelding fra Maxima:" & vbCrLf & vbCrLf & KommentarOutput
    'End If

    GoTo Slut
Fejl:
Slut:
End Sub
Sub GenerateOutputArray(Optional Text As String)
'    Dim text As String
    Dim Output As String
    Dim nr As Integer
    Dim AntalCifre As Integer
    Dim Pos As Long
    Dim pos2 As Long
    Dim posb As Long
    Dim ArrIndex As Integer 'mac
    ArrIndex = 1 'mac

    On Error GoTo Fejl
    If Text = "" Then Text = MaxProc.LastMaximaOutput
    'text = Right(text, Len(text) - 270)
    ReDim Minput(1 To 100) As String
    ReDim moutput(1 To 100) As String
    Dim slutnu As Boolean
    'MsgBox text
    
'    pos = InStr(1, text, vbLf & "(%")
'    pos2 = InStr(pos + 2, text, "(%") - 1
   Pos = FindNextio(1, Text) - 1
   pos2 = FindNextio(Pos + 2, Text) - 1

    'pos2 = InStr(pos + 2, text, vbLf & "(%")
    Do
        If Pos < 1 Then
            GoTo Slut
        End If
        If pos2 < 1 Then
            pos2 = Len(Text)
            slutnu = True
        End If
        posb = InStr(Pos, Text, ")")

        nr = CInt(Mid(Text, Pos + 4, posb - Pos - 4))
        If nr < 10 Then
            AntalCifre = 1
        ElseIf nr < 100 Then
            AntalCifre = 2
        ElseIf nr < 1000 Then
            AntalCifre = 3
        Else
            AntalCifre = 4
        End If

        Output = Mid(Text, posb + 1, pos2 - Pos - 4 - AntalCifre)
        If Mid(Text, Pos + 3, 1) = "o" Then Output = RemovePrintetBool(Output)
        Output = Replace(Trim(Output), VbCrLfMac, "")
'#If Mac Then ' SBCL
        If Mid(Text, Pos + 3, 1) = "i" Then
            ' n*ae*ste er input
            Minput(ArrIndex) = Minput(ArrIndex) & Output
        Else
            ' n*ae*ste er output
            moutput(ArrIndex) = moutput(ArrIndex) & Output
            ArrIndex = ArrIndex + 1 'mac
        End If
'#Else ' GCL
'        If Mid(text, pos + 3, 1) = "i" Then
'            ' n*ae*ste er input
'            Minput(nr) = Minput(nr) & Output
'        Else
'            ' n*ae*ste er output
'            moutput(nr) = moutput(nr) & Output
'        End If
'#End If
        Pos = pos2
'        pos2 = InStr(pos + 3, text, vbLf & "(%")
      pos2 = FindNextio(Pos + 3, Text) - 1
    Loop Until slutnu
'#If Mac Then
    nr = ArrIndex 'mac
'#End If
    If moutput(nr) = "" And Len(Minput(nr)) < 6 Then
        nr = nr - 1
    End If


    'If nr > 1 Then
    'If Definitions And Len(Minput(2)) > 1 Then
    '    DefFejl = True
    'End If
    'Else
    '
    'End If

    '    ShowOutputArray ' til test

    GoTo Slut
Fejl:
    '    MsgBox text
Slut:
   If nr < 1 Then nr = 1
    ReDim Preserve Minput(1 To nr) As String
    ReDim Preserve moutput(1 To nr) As String
End Sub
Function FindNextio(StartPos As Long, Text As String) As Long

On Error GoTo Fejl
Dim Pos As Long, pos2 As Long, io As String
Dim ea As New ExpressionAnalyser

FindNextio = 0
ea.Text = Text
Do
   Pos = InStr(StartPos, Text, "(%")
   If Pos <= 0 Then Exit Function
   io = ea.ChrByIndex(Pos + 2)
   If io = "o" Or io = "i" Then
      If ea.IsNumber(Pos + 3) Then
         FindNextio = Pos
         Exit Function
      End If
   End If
   StartPos = Pos + 2
Loop While StartPos < Len(Text)
   
GoTo Slut
Fejl:
   FindNextio = 0
'    MsgBox "Fejl " & Err.Number & " (" & Err.Description & ") i procedure FindNextio, linje " & Erl & ".", vbOKOnly Or vbCritical Or vbSystemModal, "Fejl"
Slut:
End Function
'#If Mac Then
Function RemovePrintetBool(s As String) As String
' removes the true/false which can be printed in output on Mac
' mac auto. Der kan blive printet et true eller false i slutningen
    Dim Arr As Variant
    Dim t As String, ant As Integer, i As Integer
    On Error GoTo Fejl
    s = TrimR(Trim(s), vbLf)
    s = TrimR(Trim(s), vbCr)
'#If Mac Then ' vbcrlfmac can ikke bruges her, da det er lf p*aa* mac. Ikke efter 1.23
'    arr = Split(s, vbCr)
'#Else
'    arr = Split(s, vbLf)
'#End If
Arr = Split(s, VbCrLfMac) ' *ae*ndret v. 1.23

    ant = UBound(Arr)
    If ant > 0 Then
        t = Trim(Arr(ant))
        If t = "true" Or t = "false" Or t = "unknown" Or t = "expt: undefined: 0 to a negative exponent." Then
            s = vbNullString
            For i = 0 To ant - 1
                s = s & Arr(i)
            Next
        End If
    End If
GoTo Slut
Fejl:
    MsgBox "removeprintetbool error", vbOKOnly, Sprog.Error
Slut:

    RemovePrintetBool = s
End Function
'#End If
Function GetMaximaUnitOutput() As String
    Dim Text As String
    Dim Output As String
    Dim nr As Integer
    Dim AntalCifre As Integer
    Dim Pos As Integer
    Dim pos2 As Integer
    Dim posb As Integer
    
    Dim MinputU() As String
    Dim moutputU() As String
    

    On Error GoTo Fejl
#If Mac Then
    Text = MaxProc.LastMaximaOutput
#Else
    Text = MaxProcUnit.LastMaximaOutput
#End If
    ReDim MinputU(1 To 100) As String
    ReDim moutputU(1 To 100) As String
    Dim slutnu As Boolean
'    MsgBox text

    Pos = InStr(1, Text, vbLf & "(%")
    pos2 = InStr(Pos + 2, Text, "(%") - 1
    Do
        If Pos < 1 Then
            Exit Function
        End If
        If pos2 < 1 Then
            pos2 = Len(Text)
            slutnu = True
        End If
        posb = InStr(Pos, Text, ")")

        nr = CInt(Mid(Text, Pos + 4, posb - Pos - 4))
        If nr < 10 Then
            AntalCifre = 1
        ElseIf nr < 100 Then
            AntalCifre = 2
        Else
            AntalCifre = 3
        End If
        Output = Replace(Trim(Mid(Text, posb + 1, pos2 - Pos - 4 - AntalCifre)), vbCrLf, "")

        If Mid(Text, Pos + 3, 1) = "i" Then
            ' n*ae*ste er input
            MinputU(nr) = MinputU(nr) & Output
        Else
            ' n*ae*ste er output
            moutputU(nr) = moutputU(nr) & Output
        End If
        Pos = pos2
        pos2 = InStr(Pos + 3, Text, vbLf & "(%")
    Loop Until slutnu

    If moutputU(nr) = "" And Len(MinputU(nr)) < 6 Then
        nr = nr - 1
    End If

'    ReDim Preserve MinputU(1 To nr) As String
'    ReDim Preserve moutputU(1 To nr) As String

    '    ShowOutputArray ' til test
    GetMaximaUnitOutput = moutputU(nr)

    GoTo Slut
Fejl:
    '    MsgBox text
Slut:
End Function
Sub ShowOutputArray()
' til test
    Dim i As Integer
    For i = 1 To UBound(Minput)
        MsgBox "Minput:" & i & vbCrLf & (Minput(i)) & vbCrLf & vbCrLf & "Moutput:" & vbCrLf & moutput(i)
    Next

End Sub
Function RemoveBrackets(t As String) As String
' fjerner alle [] fra strengen
    t = Replace(t, "[", "")
    t = Replace(t, "]", "")
    t = Replace(t, "{", "")
    t = Replace(t, "}", "")
    RemoveBrackets = t
End Function
Public Property Get FilOutput() As Variant
    Dim outputfil As String
    Dim filnr As Integer
    Dim mtext As String
    filnr = FreeFile
    Open outputfil For Input As filnr
    mtext = Input$(LOF(1), 1)
    Close filnr
    FilOutput = mtext
End Property

Sub ActivateTask(navn As String)
    Dim task1 As Task
    Dim tasksave As Task
    For Each task1 In Tasks
        If InStr(task1.Name, navn) > 0 Then
            Set tasksave = task1
            Exit For
        End If
    Next

    Dim i As Integer
    On Error GoTo start
start:
    i = i + 1
    Wait (0.1)
    If i > 1 Then GoTo Slut
    tasksave.Activate

Slut:
End Sub
Sub FindVariable(Optional Text As String, Optional MaximaCode As Boolean = True)
' finder variable i texten og putter dem i Vars-string
    Dim var As String
    Dim Pos As Integer
    Dim i As Integer
    Dim j As Integer
    Dim arrint As Variant
    AntalVars = 0
    vars = ""
    IntegrationVars = ""
    If Text = "" Then
        Text = Kommando
        If Text = "" Then
            On Error GoTo Fejl
            For j = 0 To UBound(Kommandoer)
                Text = Text & "+" & Kommandoer(j)    ' der skal bare v*ae*re tegn mellem som ikke kan indg*aa* i variabel
            Next
        End If
    End If
    Dim ea As New ExpressionAnalyser
    If MaximaCode Then
        ea.Text = CodeForMaxima(Text)
    Else
        ea.Text = Text
    End If
    ea.Pos = 1
    Dim vararr(30) As String    ' gemmer variablene kun til sammenligning
    Dim varind As Integer
    Dim varexists As Boolean

    ' f*oe*rst tils*ae*t integrationsvariable
    varind = 1
    If Len(IntegrationVars) > 0 Then
        arrint = Split(IntegrationVars, ";")
        For i = 0 To UBound(arrint)
            If Len(arrint(i)) > 0 Then
                vararr(varind) = arrint(i)
                varind = varind + 1
            End If
        Next
    End If
    var = ea.GetNextVar()
    var = Replace(var, vbCrLf, "")
    Do While var <> ""
        If ea.ChrByIndex(ea.Pos + 1) <> "(" And ea.ChrByIndex(ea.Pos + 1) <> "'" And ea.ChrByIndex(ea.Pos + 1) <> VBA.ChrW(8289) And Not (InStr(DefString, var & ":") = 1 Or InStr(DefString, "$" & var & ":") > 0) And InStr(var, "qkqg") < 1 Then
            varexists = False
            For i = 1 To varind - 1
                If vararr(i) = var Then varexists = True
            Next
            If Not (varexists) And var <> "e" And var <> "i" And var <> "pi" Then
                If vars <> "" Then
                    vars = vars & ";" & var
                Else
                    vars = var
                End If
                vararr(varind) = var
                varind = varind + 1
                AntalVars = AntalVars + 1
            End If
        End If
        ea.Pos = ea.Pos + 1
        var = ea.GetNextVar()
    Loop


    '    If Vars <> "" Then
    '        Vars = Mid(Vars, 2, Len(Vars) - 2)
    '    End If

    '    ReDim Preserve Vars(1 To i - 1)
Fejl:
End Sub
Function FindDefinitions() As String
    Dim s As String
    Dim t As String
    Dim i As Integer
    Dim start As Long, sslut As Long
    Dim ra As Range

    On Error GoTo Fejl

    Application.ScreenUpdating = False
    Definitions = False
    defindex = 0
    defstringtext = ""
    ReDim ldefname(100)
    ReDim ldefvalue(100)
    start = Selection.Range.start
    sslut = Selection.Range.End
    Set ra = ActiveDocument.Range
    ra.End = start
    For i = ra.OMaths.Count To 1 Step -1
        '    MsgBox ActiveDocument.Range.OMaths(i).Range.Text
        s = ActiveDocument.Range.OMaths(i).Range.Text
        '        If ActiveDocument.Range.OMaths(i).Range.start > start Then GoTo slut

        If InStr(s, VBA.ChrW(8788)) > 0 Or InStr(s, VBA.ChrW(8797)) > 0 Or InStr(s, ":") > 0 Or InStr(s, VBA.ChrW(8801)) > 0 Then    ' kun se p*aa* felter med :  tripelligmed def= og specielt := tegn
#If Mac Then
            s = ReadEquation2(ActiveDocument.Range.OMaths(i).Range) ' kan ikke huske hvad forskel p*aa* readeq1 og 2 er. Ihvert et tilf*ae*lde p*aa* mac numerisk grafisk ligningsl*oe*sning med foreg*aa*ende definition fejler. s*aa* derfor alm. readeq p*aa* mac. april2022 skiftet tilbage til readeq2, da der opstod problemer med definitioner anden gang de blev l_st. De kom ind som tekst med en masse "" og mellemrum
#Else
            s = ReadEquation2(ActiveDocument.Range.OMaths(i).Range)
#End If

            t = AddDefinition(s)
            If t = "STOP" Then Exit For
            FindDefinitions = t & FindDefinitions
        End If
    Next
    If defindex > 0 Then
        ReDim Preserve ldefname(defindex - 1)
        ReDim Preserve ldefvalue(defindex - 1)
    End If

    GoTo Slut
Fejl:
    deffejl = True
    '    DefString = ""
    FindDefinitions = ""
Slut:
    If FindDefinitions <> "" Then
        '            FindDefinitions = "[" & Left(FindDefinitions, Len(FindDefinitions) - 1) & "]$"
        Definitions = True
    End If
    DefString = FindDefinitions
    Selection.start = start
    Selection.End = sslut
    '    Selection.Collapse (wdCollapseStart)

End Function
Function AddDefinition(s As String) As String
   Dim t As String, i As Integer, Expr As String, c As String, j As Integer, Pos As Integer, pos2 As Integer, k As Integer
   Dim ea As New ExpressionAnalyser
   Dim ea2 As New ExpressionAnalyser
   Dim defenk As String, ass As String
   Dim Arr As Variant, arr2() As String, arr3() As String
   Dim deflign As Boolean

   ea.SetNormalBrackets
   s = Replace(s, VBA.ChrW(12310), "")    ' specielle usynlige paranteser fjernes
   s = Replace(s, VBA.ChrW(12311), "")    ' specielle usynlige paranteser fjernes
   s = Replace(s, vbCrLf, "")    ' enter kan forekomme
   s = Replace(s, vbCr, "")    '

   If InStr(s, VBA.ChrW(8788)) > 0 Or InStr(s, VBA.ChrW(8797)) > 0 Or InStr(s, ":=") > 0 Or InStr(VBA.LCase(s), "definer:") > 0 Or InStr(VBA.LCase(s), "define:") > 0 Or InStr(VBA.LCase(s), "ligning:") > 0 Or InStr(VBA.LCase(s), "equation:") > 0 Or InStr(VBA.LCase(s), "slet def") > 0 Or InStr(VBA.LCase(s), "delete def") > 0 Or InStr(s, VBA.ChrW(8801)) > 0 Then

      If VBA.LCase(Left(s, 8)) = "definer:" Then
         s = right(s, Len(s) - 8)
      ElseIf VBA.LCase(Left(s, 7)) = "define:" Then
         s = right(s, Len(s) - 7)
      ElseIf VBA.LCase(Left(s, 16)) = "definer ligning:" Then
         s = right(s, Len(s) - 16)
         deflign = True
      ElseIf VBA.LCase(Left(s, 16)) = "define equation:" Then
         s = right(s, Len(s) - 16)
         deflign = True
      End If
      s = Replace(s, VBA.ChrW(8788), ":=")    ' \equiv og
      s = Replace(s, VBA.ChrW(8797), ":=")    ' def=
      s = Replace(s, VBA.ChrW(8801), ":=")    ' specielt := tegn

      s = Replace(s, VBA.ChrW(11), "")

      '        If InStr(s, "block") < 1 Then
      '            s = Replace(s, VBA.ChrW(34), "") ' apostrof
      '            s = Replace(s, "=", ":")
      '            s = Replace(s, "::", ":")
      '            s = Replace(s, "):", "):=")
      '        End If

      '        s = Replace(s, "):=:", "):=")

      If InStr(VBA.LCase(s), "slet def") > 0 Or InStr(VBA.LCase(s), "delete def") > 0 Then
         '            FindDefinitions = ""
         Dim sletvar As String
         s = Replace(Trim(Split(s, ":")(1)), vbCr, "")
         If s = "" Then
            AddDefinition = "STOP"
            Exit Function
         Else
            '                AddDefinition = "kill(" & Replace(Replace(s, ";", ","), vbCr, "") & ")$" & AddDefinition
            AddDefinition = "kill(" & CodeForMaxima(Replace(s, ",", ";"), CASengine) & ")$" & AddDefinition
         End If
      Else    ' gem definitioner
         '            FindDefinitions = CodeForMaxima(s) & "," & FindDefinitions
         defstringtext = defstringtext & s & "$"
         ea.Text = Replace(s, ";", " , ")    ' fors*oe*g med at fjerne semikolon
         ea.ConvertDecSeparator
         ea.Pos = 1
         Do
            defenk = Trim(ea.GetNextListItem(ea.Pos, ","))
            If Len(defenk) = 0 Then GoTo hop
            If InStr(defenk, "block") > 0 Then
               defenk = Replace(defenk, ";", ",")
               defenk = Replace(defenk, VBA.ChrW(9633), "")    ' tegn foran paranteser der g*oe*r dem usynlige. Muligvis problematisk da paranteserne ikke fjernes. g*oe*r ogs*aa* br*oe*ker sm*aa*
               defenk = Replace(defenk, VBA.ChrW(9508), "")    ' tegn for tom plads ved ihvertfald [ ]
               defenk = Replace(defenk, VBA.ChrW(11), "")    ' shift-enter
               defenk = Replace(defenk, vbLf, "")    ' shift-enter og enter
               defenk = Replace(defenk, vbCrLf, "")
               defenk = Replace(defenk, vbCr, "")
               defenk = Replace(defenk, VBA.ChrW(183), "*")    ' prik erstattes med gange
               defenk = Replace(defenk, VBA.ChrW(8901), "*")    ' \cdot , ved ikke hvorfor der tilsyneladende er to
               defenk = Replace(defenk, VBA.ChrW(8729), "*")    ' prik \cdot erstattes med gange
               defenk = Replace(defenk, VBA.ChrW(8226), "*")    ' tyk prik erstattes med gange
               defenk = Replace(defenk, VBA.ChrW(247), "/")    '
               defenk = Replace(defenk, VBA.ChrW(8800), "#")    ' ikkeligmed
               defenk = Replace(defenk, VBA.ChrW(8804), "<=")    '
               defenk = Replace(defenk, VBA.ChrW(8805), ">=")    '
               defenk = Replace(defenk, VBA.ChrW(160), "")    ' foran accent f.eks. vector pil 160 8407
               defenk = Replace(defenk, VBA.ChrW(8289), "") ' funktionstegn
               defenk = Replace(defenk, VBA.ChrW(9618), "") ' funktionstegn
               defenk = Replace(defenk, VBA.ChrW(12310), "(") ' skjulte parenteser
               defenk = Replace(defenk, VBA.ChrW(12311), ")") ' skjulte parenteser
            ElseIf InStr(defenk, VBA.ChrW(9608)) > 0 And InStr(defenk, VBA.ChrW(9508)) > 0 Then ' tuborg parentes
            
               ea2.Text = defenk
               ea2.SetNormalBrackets
               i = 0
               Do
                  Pos = InStr(ea2.Text, "{" & VBA.ChrW(9608))
                  If Pos > 0 Then
                     ea2.Text = Replace(ea2.Text, ")=", "):=", 1, Pos)
                     Pos = InStr(ea2.Text, "{" & VBA.ChrW(9608))
                     pos2 = InStr(Pos, ea2.Text, VBA.ChrW(9508))
                     
                     Expr = Mid(ea2.Text, Pos + 2, pos2 - Pos - 1)
                     Expr = ea2.GetNextBracketContent(Pos)
                     arr2 = Split(Expr, "@")
                     t = "(if "
                     For j = 0 To UBound(arr2)
                        If InStr(arr2(j), ";") > 0 Then
                           c = ";"
                        ElseIf InStr(arr2(j), ", ") > 0 Then
                           c = ", "
                        ElseIf InStr(arr2(j), " ,") > 0 Then
                           c = " ,"
                        ElseIf InStr(arr2(j), ",") > 0 Then
                           c = ","
                        ElseIf InStr(arr2(j), "|") > 0 Then
                           c = "|"
                        Else
                           c = ""
                        End If
                        arr3 = Split(arr2(j), c)
                        If UBound(arr3) > 1 Then
                           k = 1
                           ass = ""
                           Do While k <= UBound(arr3) ' hvis delt op forkert s*aa* repareres
                              If Len(ass) > 0 Or InStr(arr3(k), "<") > 0 Or InStr(arr3(k), ">") > 0 Or InStr(arr3(k), VBA.ChrW(8712)) > 0 Then
                                 If ass = "" Then
                                    ass = arr3(k)
                                 Else
                                    ass = ass & c & arr3(k)
                                 End If
                              Else
                                 arr3(0) = arr3(0) & c & arr3(k)
                              End If
                              k = k + 1
                           Loop
                        ElseIf UBound(arr3) = 1 Then
                           ass = arr3(1)
                        End If
                        If UBound(arr3) > 0 Then
                           '                           t = t & sqbrAssump(SplitAssump(CodeForMaxima(Trim(arr3(1))))) & " then " & CodeForMaxima(arr3(0))
                           t = t & sqbrAssump(SplitAssump(Trim(ass))) & " then " & CodeForMaxima(arr3(0), CASengine)
                        Else
                           t = t & " 0=0 then " & CodeForMaxima(arr3(0), CASengine)
                        End If
                        If j < UBound(arr2) Then t = t & " elseif "
                     Next
                     ea2.Text = Left(ea2.Text, Pos - 1) & t & ")"  ' & Right(ea2.Text, Len(ea2.Text) - pos - Len(expr) - 6)
                     i = i + 1
                  End If
               Loop While Pos > 0 And i < 1000
               defenk = ea2.Text
            
            
            Else
               defenk = Replace(defenk, VBA.ChrW(34), "")    ' apostrof
               defenk = CodeForMaxima(defenk, CASengine)
'#If Mac Then ' p*aa* mac bruges ## som multiplikator ved vektorer, s*aa* ikke n*oe*dvendigvis 'ikke ligmed'  nu ogs*aa* p*aa* windows 1.18
               If InStr(defenk, "<") > 0 Or InStr(defenk, ">") > 0 Or (InStr(defenk, "#") > 0 And InStr(defenk, "##") <= 0) Then
'#Else
'               If InStr(defenk, "<") > 0 Or InStr(defenk, ">") > 0 Or InStr(defenk, "#") > 0 Then
'#End If
                  AddDefinition = AddDefinition & GetAssumeText(defenk)
                  GoTo hop
               End If
               If Not deflign Then
                  defenk = Replace(defenk, "=", ":")
                  defenk = Replace(defenk, "::", ":")
                  defenk = Replace(defenk, "):", "):=")
               End If
            End If
            Arr = Split(defenk, ":")
            If UBound(Arr) = 0 Then
               '                    If InStr(defenk, "setunits") > 0 Or InStr(defenk, "assume") > 0 Then
               GoTo hop2
               '                    End If
            End If
            If Not ValidateDef(Arr(0), Arr(1)) Then GoTo hop

            If Left(Arr(1), 1) = "=" Then Arr(1) = right(Arr(1), Len(Arr(1)) - 1)
            If Arr(1) = "" Then GoTo hop
            If UBound(ldefname) < defindex Then
               ReDim Preserve ldefname(defindex)
               ReDim Preserve ldefvalue(defindex)
            End If
            ldefname(defindex) = Arr(0)
            ldefvalue(defindex) = Arr(1)
            defindex = defindex + 1
hop2:                    ' hvis setunit eller assume eller declare
            ea2.Text = defenk
            ea2.Pos = 1
            t = ea2.GetNextVar

            If InStr(Split(defenk, ":")(0), "diff") > 0 Then
               MsgBox Sprog.A(116) & s
            ElseIf (t = "e" And InStr(defenk, "[") <= 0) Or t = "integrate" Or t = "sin" Or t = "cos" Or t = "tan" Or t = "diff" Or t = "solve" Or t = "at" Or t = "for" Or t = "in" Or t = "do" Or t = "if" Or t = "and" Or t = "or" Or t = "product" Or t = "step" Or t = "next" Or t = "from" Or t = "limit" Or t = "sum" Or t = "then" Or t = "and" Or t = "or" Or t = "else" Or t = "unless" Or t = "while" Then
               MsgBox Sprog.Variable & " " & t & " " & Sprog.A(434), vbOKOnly, Sprog.Error
            ElseIf t = "setunits" Or t = "assume" Or t = "declare" Then
               '                    AddDefinition =  AddDefinition & CodeForMaxima(defenk) & "$"
               AddDefinition = AddDefinition & defenk & "$"
            ElseIf t <> "" Then
               KillDef = KillDef & t & ","
               AddDefinition = AddDefinition & defenk & "$"
            End If
hop:
         Loop Until ea.Pos >= Len(ea.Text)
      End If
   End If

End Function
Function SplitAssump(s As String) As String
   Dim ea As New ExpressionAnalyser
   Dim v As String, p As Integer, p2 As Integer
   Dim pre As String, pst As String
   
   ea.SetNormalBrackets
   
    s = Replace(s, VBA.ChrW(8804), "<=")    '
    s = Replace(s, VBA.ChrW(8805), ">=")    '

   ea.Text = s
   
   v = ea.GetNextVar()
   If v = "" Then
      SplitAssump = s
      Exit Function
   End If
   
   p = InStr(s, "<=" & v & "<=")
   If p > 0 Then
      pre = Left(s, p - 1)
      pst = right(s, Len(s) - p - 4)
      SplitAssump = v & ">=" & pre & " and " & v & "<=" & pst
'      SplitAssump = pre & "<=" & v & " and " & v & "<=" & pst
      Exit Function
   End If
   p = InStr(s, "<=" & v & "<")
   If p > 0 Then
      pre = Left(s, p - 1)
      pst = right(s, Len(s) - p - 3)
      SplitAssump = pre & "<=" & v & " and " & v & "<" & pst
      Exit Function
   End If
   p = InStr(s, "<" & v & "<=")
   If p > 0 Then
      pre = Left(s, p - 1)
      pst = right(s, Len(s) - p - 3)
      SplitAssump = pre & "<" & v & " and " & v & "<=" & pst
      Exit Function
   End If
   p = InStr(s, "<" & v & "<")
   If p > 0 Then
      pre = Left(s, p - 1)
      pst = right(s, Len(s) - p - 2)
      SplitAssump = pre & "<" & v & " and " & v & "<" & pst
      Exit Function
   End If
   SplitAssump = s
End Function
Function sqbrAssump(s As String) As String
   Dim ea As ExpressionAnalyser, t As String
   Dim v As String, p As Integer, p2 As Integer, c As String
   Dim pre As String, pst As String, Arr() As String, l1 As String, l2 As String
   t = s
   p = InStr(s, VBA.ChrW(8712))
   If p <= 0 Then
      sqbrAssump = s
      Exit Function
   End If
   Set ea = New ExpressionAnalyser
   ea.Text = s
   v = ea.GetNextVar()
   If v = "" Then
      sqbrAssump = s
      Exit Function
   End If
   ea.Pos = p
   s = right(s, Len(s) - p)
   ea.Text = s
   ea.StartBracket = ea.ChrByIndex(1)
   ea.EndBracket = ea.ChrByIndex(ea.Length)
   
   If ea.StartBracket = "[" Then
      l1 = ">="
   ElseIf ea.StartBracket = "]" Then
      l1 = ">"
   Else
      sqbrAssump = t
      Exit Function
   End If
   If ea.EndBracket = "[" Then
      l2 = ">"
   ElseIf ea.EndBracket = "]" Then
      l2 = ">="
   Else
      sqbrAssump = t
      Exit Function
   End If
   ea.Pos = 1
   s = ea.GetNextBracketContent
   If InStr(s, ";") > 0 Then
      c = ";"
   ElseIf InStr(s, ", ") > 0 Then
      c = ", "
   ElseIf InStr(s, " ,") > 0 Then
      c = " ,"
   Else
      c = ","
   End If
   Arr = Split(s, c)
   If UBound(Arr) = 1 Then
      sqbrAssump = v & l1 & Arr(0) & " and " & Arr(1) & l2 & v
   Else
      sqbrAssump = t
   End If

End Function
Sub ResetDefinitions()
    ReDim ldefname(100)
    ReDim ldefvalue(100)
    defindex = 0

End Sub
Function ValidateDef(ByVal lhs As String, ByVal rhs As String) As Boolean
' returnerer false hvis der er fejl
    Dim ea As New ExpressionAnalyser
    Dim Pos As Integer
    On Error GoTo Fejl
    ValidateDef = False

    Pos = InStr(lhs, "[")    ' Man kan skrive alt muligt i subscript/index
    If Pos > 0 Then
        lhs = Left(lhs, Pos - 1)
    End If

    ' If Not ea.IsAllText Then ikke god nok da , og ; ikke bliver taget h*oe*jde for
    If InStr(lhs, "+") > 0 Or InStr(lhs, "-") > 0 Or InStr(lhs, "*") > 0 Or InStr(lhs, "/") > 0 Then
        MsgBox Sprog.A(117) & vbCrLf & vbTab & ConvertToAscii(lhs & "=" & rhs) & vbCrLf & Sprog.A(118), vbOKOnly, Sprog.Error
        Exit Function
    End If

    If right(rhs, 1) = "+" > 0 Then
        MsgBox Sprog.A(117) & vbCrLf & vbTab & ConvertToAscii(lhs & "=" & rhs) & vbCrLf & Sprog.A(118), vbOKOnly, Sprog.Error
        Exit Function
    End If


    ValidateDef = True
Fejl:
End Function
Function GetAssumeText(defenk) As String
' tager text fra word og omdanner til assume samt inds*ae*tter til forget
    Dim ea As ExpressionAnalyser
    Dim ea2 As ExpressionAnalyser
    Dim s As String
    Dim Arr As Variant
    Dim p1 As Integer, p2 As Integer
    If Len(defenk) > 0 Then
        If Left(defenk, 6) = "assume" Then
            GetAssumeText = defenk & "$"
            Set ea = New ExpressionAnalyser
            ea.Text = defenk
            ea.SetNormalBrackets
            defenk = ea.GetNextBracketContent(6)
            ForgetList = ForgetList & defenk & ","
        Else
            Set ea = New ExpressionAnalyser
            Set ea2 = New ExpressionAnalyser
            ea.Text = defenk
            defenk = ""
            ea.SetNormalBrackets
            Do
                s = ea.GetNextListItem(ea.Pos)
                If InStr(s, "#") > 0 Then
                    Arr = Split(s, "#")
                    s = "notequal(" & Arr(0) & "," & Arr(1) & ")"
                ElseIf Len(s) > 0 Then
                    ea2.Text = s
                    p1 = ea2.FindNextGreaterOrLower(1)
                    p2 = ea2.FindNextGreaterOrLower(p1 + 1)
                    If p2 > 0 Then    ' 2<=x<=3
                        If CASengine = 0 Then
                            If ea.ChrByIndex(p1 + 1) = "=" Then
                                s = Left(s, p2 - 1) & "," & right(s, Len(s) - p1 - 1)
                            Else
                                s = Left(s, p2 - 1) & "," & right(s, Len(s) - p1)
                            End If
                        End If
                    End If
                End If
                If Len(s) > 0 Then defenk = defenk & s & ","
            Loop While s <> ""
            defenk = Left(defenk, Len(defenk) - 1)
            ForgetList = ForgetList & defenk & ","
            GetAssumeText = "assume(" & defenk & ")$"
        End If
    End If

End Function


Sub AddToKillDefs(DefS As String)
    Dim Arr As Variant
    Dim j As Integer
    Dim s As String
    If DefS <> "" Then
        Arr = Split(DefS, ListSeparator)
        For j = 0 To UBound(Arr)
            s = Arr(j)
            s = Split(s, "=")(0)
            s = Split(s, "(")(0)
            KillDef = KillDef & s & ","
        Next
    End If

End Sub

Function IsSolved(ByVal Ligning As String, variabel As String) As Boolean
' checker om ligningen er helt l*oe*st for variabelen
' starter den med variabel= ?
' indg*aa*r variabel p*aa* h*oe*jresiden af =tegn
    On Error GoTo Slut
    Dim NotSolved As Boolean
    Dim ea As New ExpressionAnalyser
    Dim var As String
    Dim Arr As Variant
    If CASengine > 0 Then
        ea.SetTuborgBrackets
    End If
    Ligning = RemoveBrackets(Ligning)
    Arr = Split(Ligning, "=")

    If UBound(Arr) = 1 Then    ' der er pr*ae*cis et =tegn
        ea.Text = Arr(1)    'h*oe*jresiden
        IsSolved = True
        Do
            var = ea.GetNextVar
            ea.Pos = ea.Pos + 1
            If var = variabel Then
                IsSolved = False    'variabel fundet p*aa* h*oe*jresiden s*aa* er det ihvertfald ikke l*oe*st
                GoTo Slut
            End If
        Loop Until var = ""

        If Left(Ligning, 1 + Len(variabel)) = variabel & "=" Then  '  ikke nok alene
            IsSolved = True
        ElseIf Len(variabel) > 1 And Mid(Ligning, Len(variabel) + 1, 1) = "=" Then    ' hvis l*ae*ngde af variabel samme som venstresiden. N*oe*dvendig for variable som f.eks. 1+r da det kan *ae*ndres til r+1. Men nok ikke 100% sikker
            IsSolved = True
        Else    ' det m*aa* antages at den ikke er l*oe*st hvis den ikke passer med noget ovenfor
            IsSolved = False
        End If
    Else
        IsSolved = False
    End If

    GoTo Slut
Fejl:
    IsSolved = False
Slut:
End Function

Function IsAllSolved(ByVal ligninger As String, ByVal variabel As String, Optional ByVal Sep As String = ",") As String
' checker om en liste af ligninger er l*oe*st helt
' true hvis en er l*oe*st, false hvis ingen. All hvis alle.
    Dim oneSolved As Boolean
    Dim allsolved As Boolean
    Dim Arr As Variant
    Dim i As Integer
    oneSolved = False
    allsolved = True
    If ligninger = "" Then
        IsAllSolved = "false"
        Exit Function
    End If

    If Not LmSet Then
        If InStr(ligninger, "%if(") > 0 Then GoTo Slut

        ligninger = RemoveBrackets(ligninger)
        Arr = Split(ligninger, Sep)

        For i = 0 To UBound(Arr)
            If IsSolved(Arr(i), variabel) Then
                oneSolved = True
            Else
                allsolved = False
            End If
        Next
    End If
Slut:
    If allsolved Then
        IsAllSolved = "all"
    ElseIf oneSolved Then
        IsAllSolved = "true"
    Else
        IsAllSolved = "false"
    End If


End Function

Public Property Get KommandoerStreng() As Variant
    Dim i As Integer
    On Error GoTo Fejl

    If Kommando <> "" Then
        KommandoerStreng = Replace(Kommando, vbCr, "")
    Else
        For i = 0 To UBound(Kommandoer)
            KommandoerStreng = KommandoerStreng & Kommandoer(i) & ListSeparator
        Next
        KommandoerStreng = Left(KommandoerStreng, Len(KommandoerStreng) - 1)
    End If

    GoTo Slut
Fejl:
    KommandoerStreng = ""
Slut:
End Property
Private Function TrigFunction(Expr As String) As Boolean
    TrigFunction = False
    If InStr(Expr, "sin") > 0 Then
        TrigFunction = True
    ElseIf InStr(Expr, "cos") > 0 Then
        TrigFunction = True
    ElseIf InStr(Expr, "tan") > 0 Then
        TrigFunction = True
    End If
End Function
Public Property Get AntalKom() As Integer
' antal ligninger/udtryk
    On Error GoTo Fejl
    AntalKom = UBound(Kommandoer) + 1
    GoTo Slut
Fejl:
    If Kommando <> "" Then
        AntalKom = 1
    Else
        AntalKom = 0
    End If
Slut:
End Property

Function CheckForError() As Boolean
    Dim UFerror As UserFormError
    Dim fejltekst As String
    Dim Pos As Integer, pos2 As Integer, pos3 As Integer, pos4 As Integer
    Dim t As String
    Dim deffejl As Boolean
    Dim ndeferror As Integer

    If DefString <> "" Then
        ndeferror = 5
    Else
        ndeferror = 3
    End If
    On Error Resume Next
    CheckForError = False
    If InStr(MaximaOutput, "syntax error") > 0 Then
        fejltekst = Sprog.SyntaxErrorLong
        CheckForError = True
    ElseIf InStr(MaximaOutput, "incorrectsyntax:Missing") > 0 Then
        Pos = InStr(MaximaOutput, "incorrectsyntax:Missing")
        fejltekst = Sprog.SyntaxError & ". " & vbCrLf & Sprog.Missing & " " & Mid(MaximaOutput, Pos + 22, 1)
        deffejl = True
        CheckForError = True
    ElseIf InStr(MaximaOutput, "incorrectsyntax:Toomany") > 0 Then
        Pos = InStr(KommentarOutput, "incorrectsyntax:Toomany")
        fejltekst = Sprog.SyntaxError & ". " & vbCrLf & Sprog.TooMany & " " & Mid(MaximaOutput, Pos + 29, 1)
        deffejl = True
        CheckForError = True
    ElseIf InStr(MaximaOutput, "isnotaprefixoperator") > 0 Then
        fejltekst = GetErrorText("isnotaprefixoperator")
        deffejl = True
        CheckForError = True
    ElseIf InStr(MaximaOutput, "isnotaninfixoperator") > 0 Then
        fejltekst = GetErrorText("isnotaninfixoperator")
        deffejl = True
        CheckForError = True
    ElseIf InStr(MaximaOutput, "Prematureterminationofinputat") > 0 Then
        fejltekst = GetErrorText("Prematureterminationofinputat")
        deffejl = True
        CheckForError = True
    ElseIf InStr(MaximaOutput, "incorrectsyntax:") > 0 Then
        fejltekst = Sprog.SyntaxError & "."
        deffejl = True
        CheckForError = True
    ElseIf InStr(MaximaOutput, "lisp error") And InStr(omax.MaximaOutput, "[") > 0 Then
        fejltekst = Sprog.LispError
        CheckForError = True
    ElseIf InStr(MaximaOutput, "encounteredaLisperror") > 0 Then
        fejltekst = Sprog.LispError
        CheckForError = True
'    ElseIf InStr(KommentarOutput, "Division by 0") > 0 Then ' maybe not relevant in SBCL
'        fejltekst = Sprog.DivisionByZero
'        CheckForError = True
    ElseIf InStr(MaximaOutput, "expt:undefined:0toanegativeexponent") > 0 Then
        fejltekst = Sprog.DivisionByZero
        CheckForError = True
    ElseIf InStr(MaximaOutput, "anerrorTodebugthistry:debugmode(true)") > 0 Then
        fejltekst = Sprog.LispError
        CheckForError = True
    ElseIf MaximaOutput = "?merror(""Anumberwasfoundwhereavariablewasexpected-`solve'"")" Then
        fejltekst = Sprog.A(133) '"Du har bedt om at l*oe*se ligningen for en variabel der allerede er defineret. Inds*ae*t en 'slet def:' kommando f*oe*r ligningen"
        CheckForError = True
    ElseIf (omax.deffejl = True) Then
        fejltekst = Sprog.DefError & vbCrLf & VisDef
        CheckForError = True
    End If

    If CheckForError Then
        Set UFerror = New UserFormError
        UFerror.Label_maximaoutput.Caption = KommentarOutput & vbCrLf & vbCrLf & MaximaOutput    ' eller maxproc.lastmaximaoutput
        If deffejl Then
            If UBound(Minput) < ndeferror Then
                fejltekst = fejltekst & vbCrLf & vbCrLf & Sprog.DefError2
            Else
                fejltekst = fejltekst & vbCrLf & vbCrLf & Sprog.DefError3
            End If
            UFerror.TextBox_definitioner.Text = DefinitionsNice
        Else
            UFerror.TextBox_definitioner.visible = False
            UFerror.Label_definitioner.visible = False
        End If
        UFerror.Label_fejltekst.Caption = fejltekst
        UFerror.Show
    End If

End Function

Function GetErrorText(Text As String) As String
    Dim Pos As Integer, pos2 As Integer, pos3 As Integer, pos4 As Integer
    Dim t As String
    Dim l As Integer
    On Error Resume Next
    l = Len(Text)
    Pos = InStr(MaximaOutput, "incorrectsyntax")
    pos2 = InStr(Pos, MaximaOutput, Text)
    pos4 = InStr(pos2 + l, MaximaOutput, "^")
    If pos4 < 1 Then
        pos4 = Len(MaximaOutput)
    End If
    t = Mid(MaximaOutput, pos2 + l, pos4 - pos2 - l + 1)
    t = Replace(t, "^", vbCrLf & "    ^", 1, 1)
    GetErrorText = Sprog.SyntaxError & vbCrLf & Sprog.IllegalSymbol & ":" & vbCrLf & t

End Function
Function DefinitionsNice() As String
    Dim DefS As String
    DefS = DefString
    If Len(DefS) > 3 Then
        '    defs = Mid(defs, 2, Len(defs) - 3)
        DefS = ConvertToAscii(DefS)
        DefS = Replace(DefS, "$", vbCrLf)
        DefS = Replace(DefS, ":=", vbTab & "= ")
        DefS = Replace(DefS, ":", vbTab & "= ")
        If DecSeparator = "," Then
            DefS = Replace(DefS, ",", ";")
            DefS = Replace(DefS, ".", ",")
        End If
    End If

    DefinitionsNice = DefS
End Function

Public Property Get DefName(Index As Integer) As String
    On Error Resume Next
    DefName = ldefname(Index)
End Property

Public Property Get DefValue(Index As Integer) As String
    On Error Resume Next
    DefValue = ldefvalue(Index)
End Property

Public Property Get KommandoArray(Index As Integer) As String
    On Error Resume Next
    If UBound(Kommandoer) = 0 And Index = 0 Then
        KommandoArray = Kommando
    ElseIf Index <= UBound(Kommandoer) Then
        KommandoArray = Kommandoer(Index)
    Else
        KommandoArray = ""
    End If
End Property

Public Property Get KommandoArrayLength() As Integer
    On Error Resume Next
    KommandoArrayLength = UBound(Kommandoer)
End Property

'Public Property Get Is2DVector() As Boolean
'    Is2DVector = (Rows = 2 And Columns = 1)
'End Property


