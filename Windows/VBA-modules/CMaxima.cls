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
Public Kommando As String    ' kommando som maxima skal udføre
Private Kommandoer() As String    ' kommandoer som maxima skal udføre. KommandoerStreng må bruges public
Public MaximaOutput As String    ' resultat fra Maxima
Public KommentarOutput As String    ' evt. kommentar fra Maxima
Private AntalKommandoer As Integer    ' bruges af runmaxima til at se hvornår færdig
Public vars As String    ' variable i udtryk adskilt ad ;
Public AntalVars As Integer    ' antal variable fundet af findvariable
Private IntegrationVars As String    ' variable der bruges til integration skal ikke med i varlisten
Public DiffEqVar As String    ' variabel til differentialligninger. sikrer at subst ikke bruges i diff
Public MaximaInstalled As Boolean
Public CmdTask As Task
Public StopNow As Boolean
Public tempDefs As String    ' midlertidige definitioner kun for denne kommando
Public deffejl As Boolean    ' sættes hvis der er fejl i definitioner
Public Definitions As Boolean    ' sættes hvis der er definitioner i dokumentet
Public defstring As String    ' contains the output from finddefinition
Public defstringtext As String    ' contains definitions from finddefinitions before ConvertToMaxima
Private ldefname() As String    ' array med alle definitioner
Private ldefvalue() As String    ' array med alle definitioner
Public defindex As Integer
Public KillDef As String    'indeholder listseparatoret liste over sidste definitioner
Public ForgetList As String    ' indeholder liste over assumes der skal slettes
Private Minput() As String
Private moutput() As String
Public ConvertErrorText As String    ' hvis der opstår fejl i forsøget på at oversætte syntaks indeholder denne fejlteksten
Public MaximaInputStreng As String
Public MaximaInputStrengSec As String    ' secondary inputstring to be run after question
Private matvekt As Boolean    ' sættes hvis der registreres matricer. Så aktiveres prik og krydsprodukt.
Private matrixstartbracket As String
Private matrixendbracket As String
Private at As AutoTextEntry
Public ConvertLnLog As Boolean    ' whethter converttomaxima converts ln(x) -> log(x) and log(x)-> log(x)/log(10)
Public prevspr As String    ' bruges til når Maxima spr. gentagne gange om det samme
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
    'If at Is Nothing Then ' giver fejl hvis køres i autoexec
    '    Set at = NormalTemplate.AutoTextEntries.Add(name:="WordMat", Range:=Selection.Range)
    'End If
End Sub
Public Sub PrepareNewCommand(Optional finddef As Boolean = True)
    Dim defs As String
    On Error Resume Next    ' der kommer fejl ved autoexec
'    If at Is Nothing Then    ' giver fejl hvis køres i autoexec så skal også prøves her . 19/2  bruges vel ikke mere
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
    tempDefs = ""
    ReDim Kommandoer(0 To 0)
    ReDim Minput(1 To 100) As String
    ReDim moutput(1 To 100) As String
    MaximaInputStreng = ""
    MaximaInputStrengSec = ""
    ConvertErrorText = ""

    If finddef Then
        Definitions = False
        InsertKillDef
        defs = FindDefinitions
        If Len(defs) > 0 Then
            MaximaInputStreng = MaximaInputStreng & "[" & Left(Replace(defs, "$", ","), Len(defs) - 1) & "]$"
        End If
        MaximaInputStreng = MaximaInputStreng & "fejl;"
    End If

    Exit Sub

    ' skal nok bruges til mac på et tidspunkt:
    Dim fkommando As String
    MaximaInputStreng = "["
    'MaximaInputStreng = MaximaInputStreng & "display2d:false,"
    'MaximaInputStreng = MaximaInputStreng & "breakup:false$"  ' ikke dele udtryk op
    'MaximaInputStreng = MaximaInputStreng & "fpprec:5$" 'antal cifre der skal bruges i beregninger, default 16 mindre gør beregninger upræcise
    MaximaInputStreng = MaximaInputStreng & "fpprintprec:" & MaximaCifre & ","    ' antal cifre der skal vises fpprec styrer hvor mange der skal bruges i beregninger
    MaximaInputStreng = MaximaInputStreng & "ratprint:false,"    'ingen advarsel om konvertering fra decimal til brøk
    MaximaInputStreng = MaximaInputStreng & "ratsimpexpons:true,"  ' simplificer også eksponenter,
    MaximaInputStreng = MaximaInputStreng & "algebraic:false,"  ' false er default value, simplifikation af algebraiske heltal bl.a. noget med %i. Udtryk der er relle bliver ikke reduceret så %i fjernes helt hvis denne er false
    MaximaInputStreng = MaximaInputStreng & "logexpand:super,"  ' alle log regneregler
    'MaximaInputStreng = MaximaInputStreng & "realonly:" & VBA.LCase(MaximaComplex) & "," 'reelle eller komplekse løsninger, men kun til algsys
    If MaximaComplex Then
        MaximaInputStreng = MaximaInputStreng & "domain:complex,"  '
    Else
        MaximaInputStreng = MaximaInputStreng & "realonly:true,"    'reelle eller komplekse løsninger, men kun til algsys
    End If
    If MaximaExact = 2 Then
        MaximaInputStreng = MaximaInputStreng & "numer:true,"  ' numeriske løsninger
    Else
        MaximaInputStreng = MaximaInputStreng & "numer:false,"  ' exact
    End If
    If MaximaUnits Then
        MaximaInputStreng = MaximaInputStreng & "load(unit),setunits([N,J,W]),"    'enheder har nok bugs
    Else
        '    MaximaInputStreng = MaximaInputStreng & "load(noninteractive)," 'nogle gange spørges om fortegn på variable. giver output differentieret. er vidst ikke stabil
    End If
    If Not Radians Then
        '    MaximaInputStreng = MaximaInputStreng & "HasInvTrig(udt):=block([i,del],partswitch:true,i:0,start,del:inpart(udt,i),if del=end then return (false),if del=asin then return (true),if del=acos then return (true),if del=atan then return (true), i:i+1, go (start))," ' nødvendigt at checke for inverse trig da det tager ca. 0,5s ekstra at implicere inverse trig funktioner uden grund
        '    MaximaInputStreng = MaximaInputStreng & "ConvertToDegr(udt):=block([aliz],remfunction(sing),matchdeclare(aliz,all),tellsimp(sin(aliz),sing(180/%pi*aliz)),tellsimp(cos(aliz),cosg(180/%pi*aliz)),tellsimp(tan(aliz),tang(180/%pi*aliz)),if not(HasInvTrig(udt)) then go(slut),tellsimp(asin(aliz),%pi/180*asing(aliz)),tellsimp(acos(aliz),%pi/180*acosg(aliz)),tellsimp(atan(aliz),%pi/180*atang(aliz)),slut,udt:ev(udt,simp),clear_rules(),udt),"
    End If

    'MaximaInputStreng = MaximaInputStreng & "assume_pos_pred:true," '
    'MaximaInputStreng = MaximaInputStreng & "assume_pos:true," 'nogle gange spørges om fortegn på variable. Nu antages pos. Men så kan den ikke løse y2/y1=(x2/x1)^a for a
    MaximaInputStreng = MaximaInputStreng & "radexpand:true,"  ' default true, styrer reduktion af rødder. hvis all er sqrt(x^2)=x hvis true sqrt(x^2)=abs(x) ' all giver problemer ved reduktion af imaginære udtryk til reelle
    MaximaInputStreng = MaximaInputStreng & "listconstvars:true,"  ' for at %i også kommer med blandt variable så kompleks kan findes
    MaximaInputStreng = MaximaInputStreng & "solveradcan:false"    'default false, kan løse flere ligninger men langsommere. Bl.a. bruges log nu som løsning. Men kan ikke løse f.eks. 20=3*x^2.9 så
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
    
    If StopTime = 0 Then StopTime = 10000
    AntalB = AntalB + 1

    '    MaximaInputStreng = "mu:2;sigma:1;Integrate(1/(sqrt(2*%pi)*sigma)*%e^(-1/2*((y-mu)/sigma)^2),y,minf,1000);"
    '     MsgBox (MaximaInputStreng)
    '    Selection.InsertAfter MaximaInputStreng

    If DebugWM Then
        '        UserFormDebug.Label_time.Caption = UserFormDebug.Label_time.Caption & "vba tid: " & Timer - tid & vbCrLf
        '        ShowDebug
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
    Dim ufwait2 As UserFormWaitForMaxima
    Set ufwait2 = New UserFormWaitForMaxima

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
            AnswerQuestion StopTime, False ' muligvis problematisk efter ændret til at form kan vises med answerquestion
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
        ufwait2.Show vbModeless
        DoEvents
        MaxProc.ExecuteMaximaCommand MaximaInputStreng, 1
        FindLastMaximaOutput
        ResetMaxima
        Unload ufwait2
        '        MaxProc.CloseProcess ' gør at den hænger lidt, måske nødvendig
        ShowDebug
        Exit Sub
        GoTo Slut
    Else

        If MaxProc.Finished = 0 Then
            ufwait2.Show vbModeless
            DoEvents
            Do While MaxProc.Finished = 0 And i < 50
                If Round(i / 8) = i / 8 Then
                    ufwait2.Label_progress.Caption = ufwait2.Label_progress.Caption & "*"
                End If
                Wait (0.1)
                i = i + 1
            Loop
            ufwait2.Label_tip.Font.Size = 10
            ufwait2.Label_tip.Font.Italic = False
            ufwait2.Label_tip.Caption = GetRandomTip
        End If
        If StopNow Then
            Unload ufwait2
            MaxProc.CloseProcess
            MaxProc.StartMaximaProcess
            Exit Sub
        End If

        MaxProc.ExecuteMaximaCommand MaximaInputStreng, 0
        ufwait2.Label_progress.Caption = ufwait2.Label_progress.Caption & "*"
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
    Unload ufwait2
    GoTo Slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
Slut:
End Sub
Function ConvertMaximaUnits(Expr As String, Optional basedim As Boolean = True) As String
' Converts a maxima expression, using the maxima-image with units loaded
#If Mac Then ' noget af nedenstående må kunne bruges, men først prøves dette
    ConvertMaximaUnits = Expr
#Else
    Dim i As Integer, k As String, arr As Variant
    
    Expr = Replace(Expr, """", "") ' svaret vil typisk være kørt igennem num-funktion der tilføjer "" om tal
    
    If basedim Then ' all units are reverted to base SI-units
        k = Expr & ",numer;applyb1(%,unitrule),numer;dectalallNum(%),numer;"
    Else
        k = "numer:true$uforget(append(globalbaseunitlisting,globalderivedunitlisting))$"
        k = k & "setunits(usersetunits);"
        If Not OutUnits = vbNullString Then
            arr = Split(OutUnits, ",")
            For i = 0 To UBound(arr) ' Det skal gøres på denne måde fordi man ikke kan køre setunits på en liste hvor der optræder den samme type enhed to gange. fx eV og J
                k = k & "setunits(" & omax.ConvertUnits(arr(i)) & ")$"
            Next
        End If
'        k = k & expr & ",numer;applyb1(%,unitrule),numer;scanmap(processunits,%);dectalallNum(%),numer;"
        k = k & Expr & ",numer;ConvertUnit(%);"
        If MaximaVidNotation Then
            k = k & "scinotall(%);"
        Else
'            k = k & "dectalallNum(%),numer;" ' giver problemer for enheder med % prefix da enhederne så vises først.
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
    Dim ufwait2 As New UserFormWaitForMaxima
    
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
        ufwait2.Show vbModeless
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
                ufwait2.Label_progress.Caption = ufwait2.Label_progress.Caption & "*"
            ElseIf Round(i / 60) = i / 60 Then
                ufwait2.Label_tip.Font.Size = 10
                ufwait2.Label_tip.Font.Italic = False
                ufwait2.Label_tip.Caption = GetRandomTip
            ElseIf Len(ufwait2.Label_progress.Caption) > 33 Then
                ufwait2.Label_progress.Caption = "*"
            End If
            '        If i = 50 Then
            '            ufwait2.Label_tip.Caption = "         Regner stadig..."
            '        ElseIf i = 150 Then
            '            ufwait2.Label_tip.Caption = "       Det tager sin tid..."
            '        ElseIf i = 250 Then
            '            ufwait2.Label_progress.Caption = "**"
            '            ufwait2.Label_tip.Caption = " Jeg arbejder stadig på det."
            '        ElseIf i = 350 Then
            '            ufwait2.Label_tip.Caption = "Tryk evt. stop nu."
            '        ElseIf stoptime > 450 Then
            '            DoEvents
            If i = StopTime - 200 Then
                ufwait2.Label_tip.Caption = "om 20s stoppes automatisk."
            ElseIf i = StopTime - 150 Then
                ufwait2.Label_tip.Caption = "om 15s stoppes automatisk."
            ElseIf i = StopTime - 100 Then
                ufwait2.Label_tip.Caption = "om 10s stoppes automatisk."
            ElseIf i = StopTime - 50 Then
                ufwait2.Label_tip.Caption = "om 5s stoppes automatisk."
            ElseIf i = StopTime Then
                ufwait2.Label_tip.Caption = "Det lykkedes ikke. Stopper."
            End If
            If StopNow Then
                '            MaxProc.ConsoleInterrupt
                Unload ufwait2
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
            Unload ufwait2
            MaxProc.CloseProcess
            Exit Sub
        Else
            Unload ufwait2
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
        Dim pos As Integer
        Dim pos2 As Integer
        pos = InStr(MaxProc.LastMaximaOutput, "positive, negative or zero?")
        pos2 = InStr(MaxProc.LastMaximaOutput, "positive or negative?")
        If pos > 0 Or pos2 > 0 Then
            Dim udtryk As String
            Dim UF As New UserFormAskSign
            If pos <= 0 And pos2 > 0 Then pos = pos2
            pos2 = InStr(MaxProc.LastMaximaOutput, "Is ")
            UF.Label_udtryk.Caption = Mid(MaxProc.LastMaximaOutput, pos2 + 3, pos - pos2 - 4)
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
            svar = InputBox(MaxProc.LastMaximaOutput, "spr", "")    ' MaxProc.Question skal først til passes
        ElseIf InStr(MaxProc.LastMaximaOutput, "variables to solve for?") Then
            Dim antalvar As Integer, var As String, ea As New ExpressionAnalyser, sprtext As String
            pos = InStr(MaxProc.LastMaximaOutput, " variables to solve for?")
            antalvar = val(Mid(MaxProc.LastMaximaOutput, pos - 2, 2))
            sprtext = Sprog.A(416) & " " & antalvar & " " & Sprog.A(417)
            ea.text = right(MaxProc.LastMaximaOutput, Len(MaxProc.LastMaximaOutput) - pos - 24)
            ea.SetSquareBrackets
            ea.pos = 0
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
        If right(MaximaInputStrengSec, 1) = ";" Then    ' ved answerquestion sættes automatisk ;
            MaximaInputStrengSec = Left(MaximaInputStrengSec, Len(MaximaInputStrengSec) - 1)
        End If
        MaxProc.AnswerQuestion (MaximaInputStrengSec)
        WaitForMaximaForm StopTime
    End If

End Sub
Public Sub RunMaximaFile()
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
    shellcmd = "cmd /C " & maximasti & " < " & inputfil & " > " & outputfil   '/K holder cmd åben
    'shellcmd = "cmd /C " & maximasti & " maxima.bat < " & InputFil & " > " & OutputFil   '/K holder cmd åben
    'shellcmd = "cmd /C " & maximasti & " start /high ""maxima.bat < " & InputFil & " > " & OutputFil & """"   '/K holder cmd åben
    appnr = Shell(shellcmd, vbHide)    'vbNormalFocus vbMinimizedNoFocus vbhide

    ' vent på maxima har genereret outputfil
    i = 0
    Do
        i = i + 1
        Wait (0.1)
    Loop Until Dir(outputfil) > "" Or i > 80

    FindCmdTask    ' find kommandoprompten mens der ventes alligevel

    ' vent på outputfilen er skrevet færdig
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

        '    MsgBox "Maxima tog meget lang tid om beregningen. Så den blev afsluttet inden den var færdig.", vbOKOnly, "Fejl"
    End If


    ' vent på maximakommandoprompt er lukket
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
    Dim Result As VbMsgBoxResult
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

    Result = MsgBox("Det er ikke muligt at løse problemet uden at stille nogle spørgsmål." & vbCrLf & "Du bliver nu sendt til wxMaxima hvor du skal svare på et eller flere spørgsmål. Angiv svar som y for yes eller n for no etc. tryk shift-enter efter svarene." & vbCrLf & vbCrLf & "Det resulterende udtryk kan så kopieres til Word igen.", vbOKCancel, "wxMaxima")

    If Result = vbOK Then
        appnr = Shell(shellcmd, vbNormalFocus)    'vbNormalFocus vbMinimizedNoFocus vbhide
    End If

End Sub

Public Function WaitForMaxima(t As Integer) As Boolean
' vent på outputfil er genereret færdig eller indtil der er gået t sekunder
' returnerer true hvis den er færdig
    Dim fillen As Integer
    Dim fillen2 As Integer
    Dim i As Integer
    Dim filtext As String
    Dim NoRuns As Integer
    Dim arr As Variant
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
    Loop Until (fillen >= 297 And InStr(280, filtext, "(%i" & AntalKommandoer & ")") > 0) Or i > t * 10 Or fillen > 4000 Or InStr(279, filtext, "Incorrect syntax")  ' 301 er korteste mulige outputfil der er færdig
    'ekstra sikkerhed fjernet: fillen = fillen2 And
    'MsgBox Len(FilOutput)

    If fillen > 4000 And NoRuns = 0 Then
        CloseCmd
        arr = Split(vars, ";")
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
' typisk skal denne køres efter maximastregen er sat op til at køre kommandoen
' denne sørger så for at output kommer som anvist i indstillinger

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
' udfører den kommando som er i kommando-strengen eller angives som parameter
    Dim text As String
    If MaxCmd <> "" Then
        Kommando = MaxCmd    ' VBA.LCase måske
    End If


    Application.ScreenUpdating = False
    If TranslateToMaxima Then
        MaximaInputStreng = MaximaInputStreng & CodeForMaxima(Kommando) & ";"
    Else
        text = Kommando
        text = Replace(text, VBA.ChrW(9633), "")    ' tegn foran paranteser der gør dem usynlige. Muligvis problematisk da paranteserne ikke fjernes. gør også brøker små
        text = Replace(text, VBA.ChrW(9508), "")    ' tegn for tom plads ved ihvertfald [ ]
        text = Replace(text, VBA.ChrW(11), "")    ' shift-enter
        text = Replace(text, vbLf, "")    ' shift-enter og enter
        text = Replace(text, vbCrLf, "")
        text = Replace(text, vbCr, "")
        text = Replace(text, VBA.ChrW(183), "*")    ' prik erstattes med gange
        text = Replace(text, VBA.ChrW(8901), "*")    ' \cdot , ved ikke hvorfor der tilsyneladende er to
        text = Replace(text, VBA.ChrW(8729), "*")    ' prik \cdot erstattes med gange
        text = Replace(text, VBA.ChrW(8226), "*")    ' tyk prik erstattes med gange
        text = Replace(text, VBA.ChrW(247), "/")    '
        text = Replace(text, VBA.ChrW(8800), "#")    ' ikkeligmed
        text = Replace(text, VBA.ChrW(8804), "<=")    '
        text = Replace(text, VBA.ChrW(8805), ">=")    '
        text = Replace(text, VBA.ChrW(160), "")    ' foran accent f.eks. vector pil 160 8407
        text = Replace(text, VBA.ChrW(8289), "") ' funktionstegn
        text = Replace(text, VBA.ChrW(9618), "") ' funktionstegn
        text = Replace(text, VBA.ChrW(12310), "(") ' skjulte parenteser
        text = Replace(text, VBA.ChrW(12311), ")") ' skjulte parenteser
        
        Kommando = text
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
    Kommando = "Solve(" & Kommando & "," & variabel & ")"    ' kun til output af kommando forklaring

    ckommando = CodeForMaxima(ikommando)    ' coded kommando

    '    kommando = "CSolve(" & ckommando & "," & variabel & ")"
    If tempDefs <> "" Then
        '        MaximaInputStreng = MaximaInputStreng & "[" & CodeForMaxima(Replace(TempDefs, "=", ":")) & "]$"
        '        AddToKillDefs (TempDefs)

        MaximaInputStreng = MaximaInputStreng & AddDefinition("definer:" & tempDefs)
    End If
    If AllTrig Then
        MaximaInputStreng = MaximaInputStreng & "AllTrig:true$"
    Else
        MaximaInputStreng = MaximaInputStreng & "AllTrig:false$"
    End If

    If MaximaExact = 2 Then
        rkommando = "ev(%,numer)"    ' forsøg på at løse : 6^2=5^2+3^2-2*5*3*cos(C) løs for C med numer, giver delvist eksakt
    Else
        rkommando = "ev(%)"    ' kommando hvor log og sin konverteres
    End If
    '    rkommando = "ev(%th(2))" ' kommando hvor log og sin konverteres  th2 pga slutprut når der bruges answerquestion


    '    If Not Radians And Not MaximaExact = 2 Then ' kan ikke huske hvorfor  not num, men det er et problem når der kommer udtryk med variabel
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

    ' først en simpel reducering af ligningen. nok ikke nødvendig. Giver problemer med spørgsmål da kommando med spr. er sidste der udføres
    '    If MaximaExact = 2 Then
    '        MaximaInputStreng = MaximaInputStreng & "fullratsimp(" & ckommando & "),numer;"
    '    Else
    '        MaximaInputStreng = MaximaInputStreng & ckommando & ";"
    '    End If
    
    If MaximaUnits Then 'reducer først enheder på udtrykket inden kommandoen. Kan nogen gange give et mere kompliceret udtryk.
        If Not (InStr(ckommando, "diff") > 0 And InStr(ckommando, "subst") > 0) Then ' hvis der indgår noget f' må det ikke reduceres inden
            ckommando = Replace(ConvertMaximaUnits(ckommando), """", "")
        End If
    End If
    
'    If MaximaUnits Then ' fjernet pga maximaunitimage
'        MaximaInputStreng = MaximaInputStreng & "rhs(" & ckommando & ");"    ' kryptisk, men ellers kan komme lisp fejl f.eks. (-t+1)*m/s=0
'        MaximaInputStreng = MaximaInputStreng & "lhs(" & ckommando & ");"    ' kryptisk, men ellers kan komme lisp fejl f.eks. -(252178095*m*x-206157266*m)/(342482374*s)=0
'        MaximaInputStreng = MaximaInputStreng & "%=%th(2);"    ' kryptisk, men ellers kan komme lisp fejl f.eks. (-t+1)*m/s=0
'        MaximaInputStreng = MaximaInputStreng & "fullratsimp(%),numer;"    ' hvis lisp fejl på denne bliver % bare lig forrige
'        MaximaInputStreng = MaximaInputStreng & "applyunitrule(lhs(%)-rhs(%))=0;"
'        MaximaInputStreng = MaximaInputStreng & "CSolve(%," & variabel & ");"
    If MaximaExact = 2 Then    ' nødvendigt for at trig ligninger med sing ikke løses eksakt
        MaximaInputStreng = MaximaInputStreng & "Solve(ev(" & ckommando & ",numer)," & variabel & ");"
    Else
        MaximaInputStreng = MaximaInputStreng & "Solve(" & ckommando & "," & variabel & ");"
    End If


    '    RunMaxima reset:=False
    '    MaximaInputStreng = ""

    MaximaInputStrengSec = MaximaInputStrengSec & "fullratsimp(" & rkommando & ");"    'så log og trig konverteringer


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
    End If    ' slut på L=ø
    ResetMaxima

    If gemMaxExct >= 0 Then MaximaExact = gemMaxExct
Slut:
End Sub
Sub MaximaSolveInequality(variabel As String)
    Dim nr As Integer
    Dim arr As Variant
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
        If InStr(moutput(nr - 1), variabel & " = ") > 0 Then    ' hvis løst analytisk skal analytisk løsning bruges ellers numerisk
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
                arr = Split(MaximaOutput, ListSeparator)
                MaximaOutput = ""
                For i = 0 To UBound(arr)
                    If InStr(arr(i), "%i") < 1 Then     'kun reelle løsninger
                        If MaximaOutput <> "" Then MaximaOutput = MaximaOutput & "    " & VBA.ChrW(8743) & "    "
                        MaximaOutput = MaximaOutput & arr(i)
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
        Kommentar = "Der skal være et ligmed tegn i en ligning."
        MaximaOutput = ""
        Exit Sub
    End If
    Kommando = Replace(Kommando, "=", "-(") & ")"    ' skal være udtryk ikke ligning. Dette sættes lig nul
    Kommando = "newton(" & CodeForMaxima(Kommando) & "," & variabel & "," & guess & "," & eps & ")"
    MaximaInputStreng = MaximaInputStreng & "[ratsimpexpons:false,load(""newton1"")]$" & Kommando & ";"
    ' af uransagelige årsager kan der komme eksakte løsninger ud

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
    Kommando = Replace(Kommando, "=", "-(") & ")"    ' skal være udtryk ikke ligning. Dette sættes lig nul
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
    If tempDefs <> "" Then
        MaximaInputStreng = MaximaInputStreng & AddDefinition("definer:" & tempDefs)
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
        Kommentar = Sprog.A(139) & "Der skal være et ligmed tegn i en ligning."
        MaximaOutput = ""
        Exit Sub
    End If
    If tempDefs <> "" Then
        MaximaInputStreng = MaximaInputStreng & AddDefinition("definer:" & tempDefs)
        '        MaximaInputStreng = MaximaInputStreng & "[" & CodeForMaxima(Replace(TempDefs, "=", ":")) & "]$"
        '        AddToKillDefs (TempDefs)
    End If

    rkommando = "ev(%)"    ' kommando hvor log og sin konverteres
    '    If Not Radians And Not MaximaExact = 2 Then ' kan ikke huske hvorfor  not num, men det er et problem når der kommer udtryk med variabel
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

    ' ikke sikker på at dette er godt ved ligningssystemer
'    If MaximaUnits Then 'reducer først enheder på udtrykket inden kommandoen. Kan nogen gange give et mere kompliceret udtryk.
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

    If MaximaExact = 2 Then    ' nødvendigt for at trig ligninger med sing ikke løses eksakt
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

    MaximaInputStrengSec = MaximaInputStrengSec & "fullratsimp(" & rkommando & ");"    'så log og trig konverteringer

    AddStandardSecondary

    Call RunMaxima
    If StopNow Then GoTo Slut

    ConvertOutputToSolved2 (Variable)

Slut:
End Sub
Sub Eliminate(Variable As String)
    Dim Kommentar As String
    Dim i As Integer
    Dim arr As Variant
    If Variable = "" Then Exit Sub
    If InStr(Kommandoer(0), "=") < 1 Then
        Kommentar = Sprog.A(139) '"Der skal være et ligmed tegn i en ligning."
        MaximaOutput = ""
        Exit Sub
    End If

    If tempDefs <> "" Then
        MaximaInputStreng = MaximaInputStreng & "[" & CodeForMaxima(Replace(tempDefs, "=", ":")) & "]$"
        AddToKillDefs (tempDefs)
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
    arr = Split(MaximaOutput, ListSeparator)
    If Not (InStr(arr(0), "=") > 0) Then
        MaximaOutput = arr(0) & "=0"
    Else
        MaximaOutput = arr(0)
    End If
    For i = 1 To UBound(arr)
        If Not (InStr(arr(i), "=") > 0) Then
            MaximaOutput = MaximaOutput & "    " & VBA.ChrW(8743) & "    " & arr(i) & "=0"
        Else
            MaximaOutput = MaximaOutput & "    " & VBA.ChrW(8743) & "    " & arr(i)
        End If
    Next

Slut:
End Sub
Sub SolveSystemNumeric(ByVal Variable As String, ByVal guess As String, Optional ByVal eps As String)
    Dim Kommentar As String
    Dim i As Integer
    If Variable = "" Then Exit Sub
    If InStr(Kommandoer(0), "=") < 1 Then
        Kommentar = Sprog.A(139) '"Der skal være et ligmed tegn i en ligning."
        MaximaOutput = ""
        Exit Sub
    End If
    Dim sep As String
    sep = ListSeparator

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
    Dim inkom As String, pkommando As String, StartKommando As String, StartKommando2 As String, StartKommando3 As String
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

    If right(Kommando, 1) = "=" Then    ' fjern ligmed hvis det er yderst til højre
        Kommando = Left(Kommando, Len(Kommando) - 1)
    End If

    Do    ' gå tilbage til nærmeste ligmed
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
    StartKommando = Kommando    ' kommando før der bliver pillet ved den
    StartKommando3 = Kommando
    
    pkommando = Kommando    ' kommando før der bliver pillet ved den
    pkommando = Replace(pkommando, VBA.ChrW(12310), "")    ' specielle usynlige paranteser fjernes, giver problemer at fjerne her
    pkommando = Replace(pkommando, VBA.ChrW(12311), "")    ' specielle usynlige paranteser fjernes
    '    pkommando = Replace(pkommando, " ", "") ' ikke altid optimalt men nok det bedste. nej ikke efter underforståede gangetegn

'    StartKommando2 = pkommando ' til at sammenligne om numerisk resultat er lig input
'    StartKommando2 = Replace(StartKommando2, "sin", "sing")
'    StartKommando2 = Replace(StartKommando2, VBA.ChrW(8289), "")
'    StartKommando2 = CodeForMaxima(StartKommando2)
'    StartKommando2 = Replace(StartKommando2, "log10", "lgog")


' ændret pga maximaimage
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
    '    If Not MaximaExact = 2 Then ' kan give mærkelig resultat #1,INF med ,numer   72*e^((0.619/0.22*(e^22-e^0.22*t )) )
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
        '        MaximaInputStreng = MaximaInputStreng & kommando & ",numer;" & inkom & ";%,numer;" ' måske bruge dectalall for at få mere præcis antal bet cifre
        If MaximaBigFloat Then
            MaximaInputStreng = MaximaInputStreng & "fpprec:50$convertbfloat(" & Kommando & ");"
        Else
            MaximaInputStreng = MaximaInputStreng & Kommando & ",numer;"
        End If
'        MaximaInputStrengSec = "simplifynum(" & inkom & ");dectalallNum(%),numer;ev(is(equal(ev(scanmap(strtonum,%th(1)),nouns)," & StartKommando2 & ")),numer);"    ' dectal er nødvendig for enheder og korrekt antal bet cif. . forsøg med ,numer
'        MaximaInputStrengSec = "simplifynum(" & inkom & ");dectalallNum(%),numer;ev(is(equal(ev(scanmap(strtonum,%th(1)),nouns),%th(2))),numer);"    ' dectal er nødvendig for enheder og korrekt antal bet cif. . forsøg med ,numer
        MaximaInputStrengSec = "simplifynum(" & inkom & ");"    ' dectal er nødvendig for enheder og korrekt antal bet cif. . forsøg med ,numer
#If Mac Then
        If MaximaUnits Then
'            MaximaInputStrengSec = "applyb1(%,unitrule),numer;" 'gør det ikke bedre
            If InStr(ConvertUnits(OutUnits), "%") > 0 Then ' hvis % i outunits giver det muligvis enhed der placeres før tal
                MaximaInputStrengSec = MaximaInputStrengSec & "%,numer;" ' dectalall placerer enheder med %m prefix foran tal, så hellere forkert antal betcif?
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
        MaximaInputStrengSec = MaximaInputStrengSec & "ev(is(equal(ev(scanmap(strtonum,%th(1)),nouns),%th(2))),numer);"     ' dectal er nødvendig for enheder og korrekt antal bet cif. . forsøg med ,numer
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
        If tempDefs <> "" Then
            inkom = inkom & "," & tempDefs
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

    If MaximaExact = 0 And Not MaximaVidNotation And Not MaximaUnits Then
        Dim exactres As String
        Dim numres As String
        moutput(nr - 1) = Replace(moutput(nr - 1), vbLf, "")
        moutput(nr - 2) = Replace(moutput(nr - 2), vbLf, "")
'        numres = Replace(moutput(nr - 1), """", "") ' for upræcist.
'        exactres = Replace(moutput(nr - 3), """", "")
        numres = ConvertToWordSymbols(moutput(nr - 1))
        If UBound(moutput) > 3 Then exactres = ConvertToWordSymbols(moutput(nr - 3))
        If exactres = "" And numres = "" Then exactres = MaximaOutput    ' hvis der har været spørgsmål
        eql = MaximaOutput
        If eql = "true" Or StartKommando3 = numres Then
             eql = "="
        Else
             eql = VBA.ChrW(&H2248)
        End If
        MaximaOutput = ""
        pkommando = ConvertToWordSymbols(Replace(pkommando, DecSeparator, ".")) ' pkommando kan ikke sammenlignes 100% med output før den også er kørt igennem
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
    '    KommentarOutput = "Der kom følgende fejlmelding fra Maxima:" & vbCrLf & vbCrLf & KommentarOutput
    'End If
    
    If gemMaxExct >= 0 Then MaximaExact = gemMaxExct
Slut:
End Sub
Sub PolySolve(ByVal variabel As String)
' løser polynomium numerisk vha. speciel metode der giver alle løsninger
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
' kun beregnet til numerisk løsning
    Dim Variable As String
    Dim arr As Variant
    Dim i As Integer
    Dim sep As String
    If LmSet Then
        ' notation med L={}
        MaximaOutput = Trim(MaximaOutput)
        If MaximaOutput <> "" Then
            MaximaOutput = Mid(MaximaOutput, 2, Len(MaximaOutput) - 2)
            MaximaOutput = Replace(MaximaOutput, "[", "(")
            MaximaOutput = Replace(MaximaOutput, "]", ")")
            MaximaOutput = Replace(MaximaOutput, "=", "")
            arr = Split(Mid(Variable, 2, Len(Variable) - 2), ",")
            For i = 0 To UBound(arr)
                MaximaOutput = Replace(MaximaOutput, arr(i), "")
            Next
            MaximaOutput = "L={" & MaximaOutput & "}"
        End If
    Else    ' notation med x=...

        Dim ea As New ExpressionAnalyser
        Dim o As String
        ea.text = MaximaOutput
        MaximaOutput = ""
        ea.text = ea.GetNextBracketContent
        ea.pos = 1
        o = ea.GetNextBracketContent
        If o = "" Then
            MaximaOutput = ""
            Exit Sub
        End If
        MaximaOutput = o
        MaximaOutput = Replace(MaximaOutput, sep, "    " & VBA.ChrW(8743) & "    ")
        o = ea.GetNextBracketContent
        If o = "" Then Exit Sub    ' kun en løsning

        MaximaOutput = "(" & MaximaOutput & ")"
        o = "(" & Replace(o, sep, "    " & VBA.ChrW(8743) & "    ") & ")"
        MaximaOutput = MaximaOutput & "   " & VBA.ChrW(8744) & "   " & o

        Do
            o = ea.GetNextBracketContent
            If o = "" Then Exit Do
            o = "(" & Replace(o, sep, "    " & VBA.ChrW(8743) & "    ") & ")"
            MaximaOutput = MaximaOutput & "   " & VBA.ChrW(8744) & "   " & o
        Loop
    End If

End Sub
Sub ConvertOutputToSolved2(Optional Variable As String)
' konverterer output fra [[x=1,y=3],[x=3,y=5]] til (x=1 og x=3) v ... eller L={(1,3),(3,5)}
' kun beregnet til symbolsk løsning og mnewton og SolveIneq
    Dim arr As Variant
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
                arr = Split(Mid(Variable, 2, Len(Variable) - 2), ",")
                For i = 0 To UBound(arr)
                    MaximaOutput = Replace(MaximaOutput, arr(i), "")
                Next
                MaximaOutput = "L={" & MaximaOutput & "}"
            End If

        Else    ' notation med x=...

            Dim ea As New ExpressionAnalyser
            Dim o As String

            ea.text = MaximaOutput
            MaximaOutput = ""
            ea.text = ea.GetNextBracketContent
            ea.pos = 1
            o = ea.GetNextBracketContent
            If o = "" Then
                MaximaOutput = ea.text
                Exit Sub
            End If
            MaximaOutput = o
            MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8743) & "    ")

            o = ea.GetNextBracketContent
            If o = "" Then Exit Sub    ' kun en løsning

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

    If tempDefs <> "" Then
        MaximaInputStreng = MaximaInputStreng & AddDefinition("definer:" & tempDefs)
    End If

    If MaximaUnits Then 'reducer først enheder på udtrykket inden kommandoen. Kan nogen gange give et mere kompliceret udtryk.
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
    If MaximaUnits Then 'reducer først enheder på udtrykket inden kommandoen. Kan nogen gange give et mere kompliceret udtryk.
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
        Kommando = "ConvertToDegr(" & Kommando & ")"    ' Det er måske problematisk at ratsimp ikke kan kaldes da der expandes igen. converttodegree så virker ikke?
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
        Kommando = "ConvertToDegr(" & Kommando & ")"    ' Det er måske problematisk at ratsimp ikke kan kaldes da der expandes igen. converttodegree så virker ikke?
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
    
    If tempDefs <> "" Then
        MaximaInputStreng = MaximaInputStreng & AddDefinition("definer:" & tempDefs)
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

    If tempDefs <> "" Then
        MaximaInputStreng = MaximaInputStreng & AddDefinition("definer:" & tempDefs)
    End If

'    If MaximaUnits Then
'        Kommando = "applyunitrule(" & Kommando & ")"
'    End If

    '    kommando = "ev(ode2(" & Replace(kommando, "diff(", "'diff(") & "," & funktion & "," & variabel & "),logabs=false)" ' logabs=false fjernet da logistisk ikke løses korrekt så
    If InStr(funktion, "(") > 0 Then    ' hvis funktion
        Kommando = Replace(Kommando, funktion, "qvar")
        ea.text = Kommando
        ea.ReplaceVar Split(funktion, "(")(0), "qvar"
        Kommando = ea.text
        '        Kommando = Replace(Kommando, Split(funktion, "(")(0), "qvar") ' hvis både N(t) og N indgår
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


    Call RunMaxima
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
    Dim text As String
    Dim arr As Variant
    Dim i As Integer
    Dim fkt As String
    colindex = 0
    text = "draw2d("
    text = text & "font=""Arial"",font_size=8,"
    If SaveFile Then
        text = text & "terminal='gif,file_name=""" & Replace(Environ("TEMP"), "\", "/") & "/WordMatGraf"",pic_height=460,pic_width=660,"
    End If

    text = text & "nticks=100,axis_bottom = false,axis_top = false,axis_left = false,axis_right = false,grid=true,xtics_axis = true,ytics_axis = true"

    text = text & ",xaxis_width = 2,xaxis_color = black,xaxis_type  = solid,xaxis=true,yaxis_width = 2,yaxis_color = black,yaxis_type=solid,yaxis=true"
    '    text = text & "explicit(" & forskrifter & ",x," & xmin & "," & xmax & ",y," & ymin & "," & ymax & ")"
    If Not MaximaComplex Then
'            fkt = "'RealOnly(" & fkt & ")"
        text = text & ",draw_realpart = false"
    End If
    arr = Split(forskrifter, ";")
    For i = 0 To UBound(arr)
        fkt = omax.CodeForMaxima(arr(i))
        text = text & ",key=""" & ConvertToAscii(arr(i)) & """,color=" & GetNextColor() & ",explicit(" & fkt & "," & var & "," & xmin & "," & xmax & ")"
    Next
    If punkter <> "" Then
        text = text & ",key=""punkter"",point_type=filled_circle,point_size=1,points(" & punkter & ")"
    End If

    '    If titel <> "" Then
    '        text = text & ",title=" & titel
    '    End If
    If xaksetitel <> "" Then
        text = text & ",xlabel=" & xaksetitel
    End If
    If yaksetitel <> "" Then
        text = text & ",ylabel=" & yaksetitel
    End If

    text = text & ");"
    MaximaInputStreng = MaximaInputStreng & text
    Call RunMaxima(-30)

End Sub
Sub Draw2D(graphobj As String, df As String, xaksetitel As String, yaksetitel As String, grid As Boolean, SaveFile As Boolean, ByVal resolution As Integer)
    Dim text As String
    Dim waittime As Integer
    If df = "" Then
#If Mac Then
'        text = "gnuplot_command:""/Applications/WordMat/gnuplot/gnuplot""$set_plot_option([gnuplot_term, aqua])$draw2d("
        text = "gnuplot_command:""/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/gnuplot/gnuplot""$set_plot_option([gnuplot_term, aqua])$draw2d("
'        waittime = -20
#Else
        text = "draw2d("
#End If
        waittime = -400
    Else
#If Mac Then
        text = "gnuplot_command:""/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/gnuplot/gnuplot""$set_plot_option([gnuplot_term, aqua])$load(drawdf)$drawdf(" & df & ","
'        waittime = -25
#Else
        text = "load(drawdf)$drawdf(" & df & ","
#End If
        waittime = -600
    End If


    If SaveFile Then
#If Mac Then
        text = text & "terminal=pdf"
        text = text & ",file_name=""" & Replace(GetTempDir(), "\", "/") & "WordMatGraf""," & MacDrawDims(500, 350) & ","
#Else
        text = text & "terminal='gif"
        text = text & ",file_name=""" & Replace(GetTempDir(), "\", "/") & "WordMatGraf"",dimensions=[" & resolution * 660 & "," & resolution * 460 & "],"
#End If
'        text = text & "terminal='gif,file_name=""" & Replace(Environ("TEMP"), "\", "/") & "/WordMatGraf"",dimensions=[" & resolution * 660 & "," & resolution * 460 & "],"
        '        text = text & "terminal='gif,file_name=""" & Replace(Environ("TEMP"), "\", "/") & "/WordMatGraf"",dimensions=[1980,1380],font_size=20,"
        '        text = text & "terminal='gif,file_name=""" & Replace(Environ("TEMP"), "\", "/") & "/WordMatGraf"",dimensions=[660,460],font_size=20,"
    End If
    text = text & "axis_bottom = false,axis_top = false,axis_left = false,axis_right = false"
    text = text & ",xaxis_width = 2,xaxis_color = black,xaxis_type = solid,xaxis=true,yaxis_width = 2,yaxis_color = black,yaxis_type=solid,yaxis=true,"

    text = text & graphobj

    text = text & ",font_size=" & resolution * 12
    '    text = text & ",font_size=60"

    If grid Then
        If resolution = 1 Then
            text = text & ",grid=true"
        Else
            '    text = text & ",user_preamble = ""set arrow from graph 0.95, first 0 to graph 1, first 0""" ' pil på x-akse ikke pæn
            '            text = text & ",user_preamble = ""set grid lt -1 """ ' sorte linjer solid
            text = text & ",user_preamble = ""set grid lt 0 lc -1 lw 1"""    ' sorte dottede linjer
            '           text = text & ",user_preamble = ""set grid lt 0 lw " & resolution & """" ' linewidth>1 vises ikke korrekt på vandrette linjer
            ' lw 2 er linewidth 2, lt er linetype, lc er linecolor
        End If
    End If
    If xaksetitel <> "" Then
        text = text & ",xlabel=""" & xaksetitel & """"
    End If
    If yaksetitel <> "" Then
        text = text & ",ylabel=""" & yaksetitel & """"
    End If

    text = text & "),numer;" ' numer indført fordi der var et udtryk med numerisk integrale, hvor den bare ventede uden. Der er ingen grund til eksakt udtryk når der skal plottes.
    MaximaInputStreng = MaximaInputStreng & text
    
    Call RunMaxima(waittime)

End Sub
Sub PlotDF(lign As String, IndepVar As String, DepVar As String, xmin As String, xmax As String, ymin As String, ymax As String, xtrajec As String, ytrajec As String, parametre As String)
'plot af kurveintegraler

    Dim text As String

    text = "plotdf("

    text = text & lign

    If xmin <> "" And xmax <> "" Then
        text = text & ",[" & IndepVar & "," & DepVar & "],[" & IndepVar & "," & xmin & "," & xmax & "]"
    End If
    If ymin <> "" And ymax <> "" Then
        text = text & ",[" & DepVar & "," & ymin & "," & ymax & "]"
    End If
    If Len(xtrajec) > 0 And Len(ytrajec) > 0 Then
        text = text & ",[trajectory_at," & xtrajec & "," & ytrajec & "]"
    End If
    If Len(parametre) > 0 Then
        text = text & "," & parametre
    End If

    text = text & ");"
    MaximaInputStreng = MaximaInputStreng & text
    Call RunMaxima(-500)

End Sub
Sub Plot2Dold(forskrifter As String, xmin As String, xmax As String, ymin As String, ymax As String, xaksetitel As String, yaksetitel As String)
' forskrifter adskilt ad komma
    Dim text As String
    text = "plot2d(["
    text = text & forskrifter
    text = text & "]"

    text = text & ",[x," & xmin & "," & xmax & "]"
    If xaksetitel <> "" Then
        text = text & ",[xlabel," & xaksetitel & "]"
    End If
    If yaksetitel <> "" Then
        text = text & ",[ylabel," & yaksetitel & "]"
    End If
    '    text = text & ",[box,false]"
    If InStr(forskrifter, "discrete") > 0 Then
        text = text & ",[style,points,lines]"
    Else
        text = text & ",[style,lines]"
    End If

    text = text & ",[point_type,diamond],[axes,true],[gnuplot_out_file, false],[plot_format,xmaxima]"
    text = text & ");"
    MaximaInputStreng = text
    Call RunMaxima

End Sub
Sub Plot3D(forskrifter As String, xvar As String, xmin As String, xmax As String, yvar As String, ymin As String, ymax As String, xaksetitel As String, yaksetitel As String, kunlinjer As Boolean, gridlines As Integer)
' forskrifter adskilt ad komma
    Dim text As String
    Dim tid As Integer
    tid = 20
    tid = tid + Len(forskrifter) / 2
    text = "plot3d(["
    text = text & forskrifter

    text = text & ",[" & xvar & "," & xmin & "," & xmax & "]"
    text = text & ",[" & yvar & "," & ymin & "," & ymax & "]"

    text = text & "]"

    If xaksetitel <> "" Then
        text = text & ",[xlabel," & xaksetitel & "]"
    End If
    If yaksetitel <> "" Then
        text = text & ",[ylabel," & yaksetitel & "]"
    End If
    '    text = text & ",[box,false]"
    '    If InStr(forskrifter, "discrete") > 0 Then
    '        text = text & ",[style,points,lines]"
    '    Else
    '        text = text & ",[style,lines]"
    '    End If
    If kunlinjer = True Then
        text = text & ",[palette,false]"
    End If
    If gridlines <> 30 Then
        text = text & ",[grid," & gridlines & "," & gridlines & "]"
    End If

    '    text = text & ",[point_type,diamond],[axes,true],[gnuplot_out_file, false],[plot_format,xmaxima]"
    text = text & ");"
    MaximaInputStreng = MaximaInputStreng & text
    Call RunMaxima(-1 * tid)

End Sub
Sub Draw3D(graphobj As String, Optional antalobj As Integer = 1)
'    text = "load(draw)$draw3d("
    Dim text As String
    Dim tid As Integer

    tid = 20 + antalobj * 10

#If Mac Then
        text = "gnuplot_command:""/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/gnuplot/gnuplot""$set_plot_option([gnuplot_term, aqua])$draw3d("
        text = text & "terminal=pdf"
        text = text & ",file_name=""" & Replace(GetTempDir(), "\", "/") & "WordMatGraf""," & MacDrawDims(500, 350) & ","
'        waittime = -20
#Else
    text = "draw3d("
#End If
    text = text & graphobj

    text = text & ");"
    MaximaInputStreng = MaximaInputStreng & text

    Call RunMaxima(-1 * tid)

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

Function CodeForMaxima(ByVal text As String) As String
    Dim ea As New ExpressionAnalyser
    Dim ea2 As New ExpressionAnalyser
    Dim pos As Integer
    Dim posb As Integer
    Dim rod As Integer
    Dim pos2 As Integer
    Dim pos3 As Integer
    Dim pos4 As Integer
    Dim pos5 As Integer
    Dim posog As Integer
    Dim arr As Variant, arr2() As String, arr3() As String
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
    Dim antalm As Integer, varl As Integer, di As Integer

    On Error GoTo Fejl
    ea.SetNormalBrackets
    ea2.StartBracket = VBA.ChrW(12310)
    ea2.EndBracket = VBA.ChrW(12311)

    text = Trim(text)
    '    For j = 1 To Len(text)
    '    i = AscW(Mid(text, j, 1))
    '    s = s & Mid(text, j, 1) & " - " & i & vbCrLf
    '    Next
    '    MsgBox s


    text = Replace(text, VBA.ChrW(9633), "")    ' tegn foran paranteser der gør dem usynlige. Muligvis problematisk da paranteserne ikke fjernes. gør også brøker små
    text = Replace(text, VBA.ChrW(9508), "")    ' tegn for tom plads ved ihvertfald [ ]
    text = Replace(text, VBA.ChrW(11), "")    ' shift-enter
    text = Replace(text, vbLf, "")    ' shift-enter og enter
    text = Replace(text, vbCrLf, "")
    text = Replace(text, vbCr, "")
    text = Replace(text, """", "")    ' apostrof fjernes
    text = Replace(text, "+-", "-")    ' plus-minus giver minus
    text = Replace(text, VBA.ChrW(8212), "+")    'dobbelt minustegn giver plus
    text = Replace(text, ChrW(8711) & VBA.ChrW(183), "divq ")    ' nabla gange
    text = Replace(text, ChrW(8711) & VBA.ChrW(215), "curlq ")    ' nabla kryds - curl
    text = Replace(text, ChrW(8711) & "^2", "laplaceq ")    ' nabla
    text = Replace(text, ChrW(8711), "gradq ")    ' nabla
    text = Replace(text, VBA.ChrW(183), "*")    ' prik erstattes med gange
    text = Replace(text, VBA.ChrW(8901), "*")    ' \cdot , ved ikke hvorfor der tilsyneladende er to
    text = Replace(text, VBA.ChrW(8729), "*")    ' prik \cdot erstattes med gange
    text = Replace(text, VBA.ChrW(8226), "*")    ' tyk prik erstattes med gange
    text = Replace(text, VBA.ChrW(247), "/")    '
    text = Replace(text, VBA.ChrW(180) & VBA.ChrW(180), "^''")    ' der er tre forskellige '´` apostrof lignende de forståes alle som diff
    text = Replace(text, VBA.ChrW(96) & VBA.ChrW(96), "^''")    ' der er tre forskellige '´`
    text = Replace(text, VBA.ChrW(180), "^'")    ' der er tre forskellige '´` apostrof lignende de forståes alle som diff
    text = Replace(text, VBA.ChrW(96), "^'")    ' der er tre forskellige '´`
    text = Replace(text, VBA.ChrW(&H393) & "_incomplete ", "gamma_incomplete")
    text = Replace(text, VBA.ChrW(915), "gamma")    ' stort gammategn
    text = Replace(text, VBA.ChrW(8800), "#")    ' ikkeligmed
    text = Replace(text, VBA.ChrW(8804), "<=")    '
    text = Replace(text, VBA.ChrW(8805), ">=")    '
    text = Replace(text, VBA.ChrW(8710), "increment")    ' specielt delta increment
    text = Replace(text, "^ " & VBA.ChrW(176) & "C", "degC")    ' specielt gradtegn(\circ) og C
    text = Replace(text, "^ " & VBA.ChrW(8728) & "C", "degC")    ' specielt gradtegn(\circ) og C
    text = Replace(text, VBA.ChrW(8451), "degC")    ' specielt oC tegn
    text = Replace(text, VBA.ChrW(176), "tgrader")    ' gradtegn
    text = Replace(text, VBA.ChrW(8728), "tgrader")    ' gradtegn
    text = Replace(text, "%", "/100")    ' procenttegn
    text = Replace(text, VBA.ChrW(160), "")    ' foran accent f.eks. vector pil 160 8407
    If Not MaximaComplex Then
        text = Replace(text, VBA.ChrW(8736), "Symangle")    '
    End If

    ' som test flyttes disse fra slutningen til starten
    ' øæå ÆØÅ ' ændres ikke tilbage
    text = Replace(text, "æ", "ae")
    text = Replace(text, "ø", "oe")
    text = Replace(text, "å", "aa")
    text = Replace(text, "Æ", "AE")
    text = Replace(text, "Ø", "OE")
    text = Replace(text, "Å", "AA")

    'græske bogstaver
    text = Replace(text, VBA.ChrW(916), "Delta")
    text = Replace(text, VBA.ChrW(948), "delta")
    text = Replace(text, VBA.ChrW(945), "alpha")
    text = Replace(text, VBA.ChrW(946), "beta")
    text = Replace(text, VBA.ChrW(947), "gammaLB")
    text = Replace(text, VBA.ChrW(952), "theta")
    text = Replace(text, VBA.ChrW(920), "Theta")
    text = Replace(text, VBA.ChrW(955), "lambda")
    text = Replace(text, VBA.ChrW(923), "Lambda")
    text = Replace(text, VBA.ChrW(956), "mu")
    text = Replace(text, VBA.ChrW(961), "rho")
    text = Replace(text, VBA.ChrW(963), "sigma")
    text = Replace(text, VBA.ChrW(931), "Sigma")
    text = Replace(text, VBA.ChrW(966), "varphi")
    text = Replace(text, VBA.ChrW(981), "phi")
    text = Replace(text, VBA.ChrW(934), "Phi")
    text = Replace(text, VBA.ChrW(949), "varepsilon")
    text = Replace(text, VBA.ChrW(1013), "epsilon")
    text = Replace(text, VBA.ChrW(968), "psi")
    text = Replace(text, VBA.ChrW(936), "Psi")
    text = Replace(text, VBA.ChrW(926), "Xi")
    text = Replace(text, VBA.ChrW(958), "xi")
    text = Replace(text, VBA.ChrW(935), "Chi")
    text = Replace(text, VBA.ChrW(967), "chi")
    text = Replace(text, VBA.ChrW(928), "Pi")
    text = Replace(text, VBA.ChrW(964), "tau")
    text = Replace(text, VBA.ChrW(957), "greek-nu")
    text = Replace(text, VBA.ChrW(954), "kappa")
    text = Replace(text, VBA.ChrW(951), "eta")
    text = Replace(text, VBA.ChrW(950), "zeta")
    text = Replace(text, VBA.ChrW(969), "omega")    ' lille omega


    ' ½ symbol bliver til 1/2 eller tekst Symhalf afh. af kontekst. Symhalf hvis bogstavforan
    '    text = Replace(text, VBA.ChrW(189), "(1/2)") '
    '    text = Replace(text, VBA.ChrW(189), "Symhalf") ' 189 er symbolet for en halv
    
    ea.text = text
    pos = 0
    Do
        pos = InStr(pos + 1, ea.text, VBA.ChrW(189))
        If pos > 0 Then
            If ea.IsLetterPos(pos - 1, False) Then
                ea.RemoveChar pos
                ea.InsertBeforePos "Symhalf", pos
            Else
                ea.RemoveChar pos
                ea.InsertBeforePos "(1/2)", pos
            End If
        End If
    Loop While pos > 0
    text = ea.text

    'lambertW0
    Do
        pos = InStr(text, "W_0")
        If pos > 0 Then
            ea.text = text
            text = Left(text, pos - 1) & "generalized_lambert_w(0," & right(text, Len(text) - pos - 4)
        End If
    Loop While pos > 0
    'lambertW-1
    Do
        pos = InStr(text, "W_(-1)")
        If pos > 0 Then
            ea.text = text
            text = Left(text, pos - 1) & "generalized_lambert_w(-1," & right(text, Len(text) - pos - 7)
        End If
    Loop While pos > 0

    ' autoregistrering af matrixparenteser
    If InStr(text, "[") > 0 Then
        matrixstartbracket = "["
        matrixendbracket = "]"
    End If

    ' nu overflødig: virker ikke med (ab)\vec
    '    text = Replace(text, VBA.ChrW(8407), "SymVecta") ' vektorpil
    '    text = Replace(text, VBA.ChrW(8401), "SymVectb") ' vektorpil men ikke hel pil
    '    text = Replace(text, VBA.ChrW(773), "SymVectc") ' matrixtegn
    '    text = Replace(text, VBA.ChrW(175), "SymVectc") ' matrixtegn \overbar

    ' vektortegn til SymVecta
    ea.text = text
    pos = 1
    Do
        pos = InStr(pos, ea.text, VBA.ChrW(8407))
        If pos > 0 Then
            matvekt = True
            If ea.ChrByIndex(pos - 1) = ")" Then
                s = ea.GetPrevBracketContent(pos)
                s = Replace(s, " ", "SPC")
                MaximaInputStreng = MaximaInputStreng & "declare(" & s & "SymVecta,nonscalar)$"
                KillDef = KillDef & s & "SymVecta,"
                'AddDefinition ("definer: declare(" & s & "SymVecta,nonscalar)")
                ea.text = Left(ea.text, ea.pos - 2) & s & "SymVecta" & right(ea.text, Len(ea.text) - pos)
            Else
                '            AddDefinition ("definer: declare(" & ea.ChrByIndex(pos - 1) & "SymVecta,nonscalar)")
                s = ea.ChrByIndex(pos - 1)
                MaximaInputStreng = MaximaInputStreng & "declare(" & s & "SymVecta,nonscalar)$"
                KillDef = KillDef & s & "SymVecta,"
                ea.text = Left(ea.text, pos - 1) & "SymVecta" & right(ea.text, Len(ea.text) - pos)
            End If
            pos = pos + 7
        End If
    Loop While pos > 0

    ' vektortegn 2 ikke hel pil til SymVectb
    pos = 1
    Do
        pos = InStr(pos, ea.text, VBA.ChrW(8401))
        If pos > 0 Then
            matvekt = True
            If ea.ChrByIndex(pos - 1) = ")" Then
                s = ea.GetPrevBracketContent(pos)
                ea.text = Left(ea.text, ea.pos - 2) & s & "SymVectb" & right(ea.text, Len(ea.text) - pos)
            Else
                ea.text = Left(ea.text, pos - 1) & "SymVectb" & right(ea.text, Len(ea.text) - pos)
            End If
            pos = pos + 7
        End If
    Loop While pos > 0

    ' matrixtegn til SymVectc
    pos = 1
    Do
        pos = InStr(pos, ea.text, VBA.ChrW(773))
        If pos > 0 Then
            matvekt = True
            If ea.ChrByIndex(pos - 1) = ")" Then
                s = ea.GetPrevBracketContent(pos)
                ea.text = Left(ea.text, ea.pos - 2) & s & "SymVectc" & right(ea.text, Len(ea.text) - pos)
            Else
                ea.text = Left(ea.text, pos - 1) & "SymVectc" & right(ea.text, Len(ea.text) - pos)
            End If
            pos = pos + 7
        End If
    Loop While pos > 0


    ' matrixtegn \overbar til SymVectc
    
    pos = 1
    Do
        pos = InStr(pos, ea.text, VBA.ChrW(175))
        If pos > 0 Then
            matvekt = True
            s = ea.GetNextVar(pos)
            If Len(ea.text) = ea.pos Then
                ea.text = Left(ea.text, pos - 1) & " " & s & "SymVectc"
            Else
                ea.text = Left(ea.text, pos - 1) & " " & s & "SymVectc" & right(ea.text, Len(ea.text) - ea.pos)
            End If
            pos = pos + 7
        End If
    Loop While pos > 0

    text = ea.text

    If InStr(text, VBA.ChrW(9608)) > 0 Then matvekt = True    '9608=matrixtegn
    If InStr(text, VBA.ChrW(9632)) > 0 Then matvekt = True    '9632= andet vektortegn

    If matvekt Then  '9632?
        text = Replace(text, VBA.ChrW(215), "~")    ' vektorprodukt
    Else
        text = Replace(text, VBA.ChrW(215), "*")    ' kryds erstattes med gange
    End If

    'tværvektor \hat
    pos = 1
    Do
        pos = InStr(pos, text, VBA.ChrW(770))
        If pos > 0 Then
            ea.text = text
            If ea.ChrByIndex(pos - 1) = ")" Then
                s = ea.GetPrevBracketContent(pos)
                text = Left(text, ea.pos - 2) & "tvvec(" & s & ")" & right(text, Len(text) - pos)
                pos = pos + 4
            Else
                s = ea.GetPrevVar(pos)
                text = Left(text, pos - Len(s) - 2) & "tvvec(" & Mid(text, pos - Len(s) - 1, Len(s)) & ")" & right(text, Len(text) - pos)
                pos = pos + 5
            End If
        End If
    Loop While pos > 0

    '    text = Replace(text, VBA.ChrW(&H393) & " ", "Gamma")

    '   Fjern problematiske usynlige paranteser foran / ' løses af den efterfølgende også
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
    '        If ea.ValidateBrackets And InStr(ea.text, "=") < 1 Then ' forsøger at løse problem hvor skjulte og alm paranteser er blandet som f.eks.[( ])
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
    If InStr(text, "Sin") > 0 Then
        ea.text = text
        ea.ReplaceVar "Sin", "sin"
        text = ea.text
    End If
    If InStr(text, "Cos") > 0 Then
        ea.text = text
        ea.ReplaceVar "Cos", "cos"
        text = ea.text
    End If
    If InStr(text, "Tan") > 0 Then
        ea.text = text
        ea.ReplaceVar "Tan", "tan"
        text = ea.text
    End If
    If InStr(text, "Log") > 0 Then
        ea.text = text
        ea.ReplaceVar "Log", "log"
        text = ea.text
    End If
    If InStr(text, "Ln") > 0 Then
        ea.text = text
        ea.ReplaceVar "Ln", "ln"
        text = ea.text
    End If

    ''  her var disse før men nu flyttet efter 8289 fjernes  text = Replace(text, "cosh^(-1)", "acosh")


    ' fjern +- symbol
    pos = InStr(text, VBA.ChrW(177))
    If pos > 0 Then
        pm = InputBox(Sprog.A(421), Sprog.A(422), "-")
        If pm = "+" Or VBA.LCase(pm) = "plus" Then
            text = Replace(text, VBA.ChrW(177), "+")
        ElseIf pm = "-" Or VBA.LCase(pm) = "minus" Then
            text = Replace(text, VBA.ChrW(177), "-")
        Else
            MsgBox Sprog.A(423)
            text = Replace(text, VBA.ChrW(177), "+")
        End If
    End If

    ' kommaer og listeseparator
    If Not MaximaSeparator Then
        ea.text = text
        ea.ConvertDecSeparator
        text = ea.text
        '        text = Replace(text, ",", ".")
    End If
    text = Replace(text, ";", ",")


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


'    ' stykkevise funktioner defineret med en tuborg parantes. Ikke færdig. Et problem at der indsættes gangetegn i koden senere hen
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
    ea.text = text
    i = 0
    Do
        pos = InStr(ea.text, VBA.ChrW(124))
        If pos > 0 Then
            posb = pos
            Do
                i = i + 1    ' forkert input kan ellers skabe uendelig løkke
                posb = posb + 1
                c = ea.ChrByIndex(posb)
                If c = "|" Then
                    If Not (ea.IsOperator(ea.ChrByIndex(posb - 1), 4)) Then
                        If matvekt Or InStr(ea.text, VBA.ChrW(9632)) > 0 Then
                            ea.text = Left(ea.text, pos - 1) & "vecmag(" & Mid(ea.text, pos + 1, posb - pos - 1) & ")" & right(ea.text, Len(ea.text) - posb)
                        Else
                            ea.text = Left(ea.text, pos - 1) & "abs(" & Mid(ea.text, pos + 1, posb - pos - 1) & ")" & right(ea.text, Len(ea.text) - posb)
                        End If
                        Exit Do
                    Else
                        pos = posb
                    End If
                ElseIf c = "(" Then
                    ea.pos = posb
                    posb = ea.FindEndOfBracket() - 1
                End If
            Loop While posb < Len(ea.text) And i < 1000
            '        Loop While c <> "|" And c <> "(" And c <> "["

            posb = InStr(ea.text, VBA.ChrW(124))
        End If
    Loop While pos > 0 And i < 1000
    If i = 1000 Then
        MsgBox Sprog.A(424), vbOKOnly, Sprog.A(105)
    End If
    i = 0
    text = ea.text

    ' |x| -> abs(x) ' men der tages ikke højde for paranteser
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

    ' find underforståede paranteser efter ^ og / ' (skal være efter diff og andre funktioner med komma)
    ea.text = text
    ea.InsertBracketAfter ("/")
    ea.InsertBracketBefore ("/")
    text = ea.text
    
    ' d/dx notation
    text = ConvertDifferentials(text, VBA.ChrW(8518)) ' \dd
    If Left(text, 5) = "Fejl!" Then
        text = ea.text
        ConvertErrorText = text
        GoTo Slut
    End If
    text = ConvertDifferentials(text, VBA.ChrW(8706)) ' \partial
    If Left(text, 5) = "Fejl!" Then
        text = ea.text
        ConvertErrorText = text
        GoTo Slut
    End If
    If dAsDiffChr Then
        text = ConvertDifferentials(text, "d") ' normal d
        If Left(text, 5) = "Fejl!" Then
            text = ea.text
            ConvertErrorText = text
            GoTo Slut
        End If
    End If
    
    ' indsæt parentes ved f.eks. log 2 -> log(2)
    ea.text = text
    ea.InsertFunctionBrackets
    ea.text = Replace(ea.text, VBA.ChrW(12310), "")    ' specielle usynlige paranteser fjernes   flyttet fra før trig med store bogstaver for at få insertfunctionbrackets til at virke
    ea.text = Replace(ea.text, VBA.ChrW(12311), "")    ' specielle usynlige paranteser fjernes

    ea.text = Replace(ea.text, "cosh^(-1)", "acosh")
    ea.text = Replace(ea.text, "sinh^(-1)", "asinh")
    ea.text = Replace(ea.text, "tanh^(-1)", "atanh")
    ea.text = Replace(ea.text, "cos^(-1) ", "acos")    ' under sjældne omstændigheder kan der være et mellemrum istedet for et 8289 tegn efter
    ea.text = Replace(ea.text, "sin^(-1) ", "asin")    ' det blive forstået som et gangetegn hvis ikke fjernes
    ea.text = Replace(ea.text, "tan^(-1) ", "atan")
    ea.text = Replace(ea.text, "cos^(-1)", "acos")
    ea.text = Replace(ea.text, "sin^(-1)", "asin")
    ea.text = Replace(ea.text, "tan^(-1)", "atan")

    ea.ConvertTrigPow ("sin")    ' skal være efter inverse trig
    ea.ConvertTrigPow ("cos")
    ea.ConvertTrigPow ("tan")
    text = ea.text

    ' logaritmer med base forskellig fra 10 og e
    Dim base As String
    Dim spos As Integer
    ea.text = text
    spos = 1
    Do
        '    ea.text = text
        pos = ea.FindChr("log_", spos)
        If pos > 0 Then
            pos2 = InStr(pos + 2, ea.text, VBA.ChrW(8289))
            If Not pos2 > 0 Then
                base = ea.GetNextVar(pos)
                base = right(base, Len(base) - 4)
            Else
                base = Mid(ea.text, pos + 4, pos2 - pos - 4)
            End If
            ea.text = Left(ea.text, pos + 2) & "base" & right(ea.text, Len(ea.text) - pos - Len(base) - 4)
            ea.pos = pos + 7
            '        ea.InsertBeforePos ("(")
            '        ea.pos = ea.pos + 3
            ea.GetNextBracketContent
            ea.InsertBeforePos "," & base, ea.pos - 1
            '        text = ea.text
            '        spos = ea.pos + 10
        End If
    Loop While pos > 0
    text = ea.text


    text = Replace(text, VBA.ChrW(8289), "")    ' symbol der definerer funktion fjernes fra word syntaks

    'logaritmer
    If ConvertLnLog Then
        ea.text = text
        ea.ReplaceVar "log", "log10"
        text = ea.text
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
            pos = InStr(text, VBA.ChrW(8728 + rod))
            If pos > 0 Or pos4 > 0 Or pos5 > 0 Then
                ea.text = text
                ea.pos = pos + 1
                If Mid(text, pos + 1, 1) <> "(" Then
                    ea.InsertUnderstoodBracketPair
                End If
                ea.pos = pos
                Call ea.GetNextBracketContent    ' bare for at finde slut parantes
                Call ea.InsertBeforePos("^(1/" & rod & ")")
                ea.text = Replace(ea.text, VBA.ChrW(8728 + rod), "", 1, 1)
                text = ea.text
            End If
        Loop While pos > 0
    Next

    'kvadratrod
    Do
        pos = InStr(text, VBA.ChrW(8730))
        If pos > 0 Then
            If Mid(text, pos + 1, 1) <> "(" Then
                text = Replace(text, VBA.ChrW(8730), "sqrt", 1, 1)
                pos = pos + 4
                ea.text = text
                ea.pos = pos
                ea.InsertUnderstoodBracketPair
                text = ea.text
            Else
                ea.text = text
                ea.pos = pos
                arr = Split(ea.GetNextBracketContent, "&")
                pos2 = ea.pos
                If UBound(arr) = 0 Then
                    text = Replace(text, VBA.ChrW(8730), "sqrt", 1, 1)
                ElseIf UBound(arr) = 1 Then
                    Call ea.InsertBeforePos("^(1/(" & arr(0) & "))")
                    ea.text = Replace(ea.text, VBA.ChrW(8730), "", 1, 1)
                    posog = ea.FindChr("&", 1)
                    text = Left(ea.text, pos) & right(ea.text, Len(ea.text) - posog)
                ElseIf UBound(arr) = -1 Then ' tilføjet v1.09 pga uendelig løkke ved mangl parentes
                    MsgBox Sprog.A(169), vbOKOnly, Sprog.Error
                    GoTo Fejl
                End If
            End If
        End If
    Loop While pos > 0

    ' sumtegn
    Do
        pos = InStr(text, VBA.ChrW(8721))
        If pos > 0 Then
            ea.text = text
            ea.pos = pos
            subs = ea.FindChr("_")
            sups = ea.FindChr("^")
            ends = ea.FindChr(VBA.ChrW(9618))
            arr = Split(Mid(text, subs + 2, sups - subs - 3), "=")
            If UBound(arr) < 1 Then    ' hvis definition er = oversat til :
                arr = Split(Mid(text, subs + 2, sups - subs - 3), ":")
            End If
            If UBound(arr) < 1 Then
                MsgBox Sprog.A(425), vbOKOnly, Sprog.Error '"Notationen for sumtegnet er ikke korrekt. Den nedre grænse skal skrives f.eks. i=1"
                GoTo videre
            End If
            var = arr(0)
            startn = arr(1)
            slutn = Mid(text, sups + 1, ends - sups - 1)
            sluts = ea.FindHierakiBreak(1)
            If ea.ChrByIndex(ea.pos) = "+" Or ea.ChrByIndex(ea.pos) = "-" Or ea.ChrByIndex(ea.pos) = "*" Or ea.ChrByIndex(ea.pos) = "/" Then sluts = sluts - 1
            Expr = Mid(text, ends + 1, sluts - ends)
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
            text = Left(text, pos - 1) & "sum(" & Expr & "," & var & "," & startn & "," & slutn & ")" & right(text, Len(text) - sluts)
        End If
    Loop While pos > 0

    ' produkttegn
    Do
        pos = InStr(text, VBA.ChrW(8719))
        If pos > 0 Then
            ea.text = text
            ea.pos = pos
            subs = ea.FindChr("_")
            sups = ea.FindChr("^")
            ends = ea.FindChr(VBA.ChrW(9618))
            arr = Split(Mid(text, subs + 2, sups - subs - 3), "=")
            If UBound(arr) < 1 Then    ' hvis definition er = oversat til :
                arr = Split(Mid(text, subs + 2, sups - subs - 3), ":")
            End If
            If UBound(arr) < 1 Then
                MsgBox Sprog.A(426)
                GoTo videre
            End If
            var = arr(0)
            startn = arr(1)
            slutn = Mid(text, sups + 1, ends - sups - 1)
            sluts = ea.FindHierakiBreak(1)
            Expr = Mid(text, ends + 1, sluts - ends)
            text = Left(text, pos - 1) & "product(" & Expr & "," & var & "," & startn & "," & slutn & ")" & right(text, Len(text) - sluts)
        End If
    Loop While pos > 0

videre:

    'grænseværdier
    pos = 1
    Do
        pos = InStr(pos, text, "lim")
        '    pos = InStr(text, "lim" & VBA.ChrW(9516))
        '    pos = InStr(text, "lim" & VBA.ChrW(9516))
        If pos > 0 Then
            ea.text = text
            If ea.ChrByIndex(pos + 3) = VBA.ChrW(9516) Or ea.ChrByIndex(pos + 3) = "_" Then
                ea.pos = pos
                var = ea.GetNextBracketContent
                ends = ea.pos

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
                Expr = Mid(text, ends, sluts - ends + 1)

                arr = Split(var, VBA.ChrW(8594))
                var = Trim(arr(0))
                DiffEqVar = var
                app = Trim(arr(1))
                If right(app, 2) = "^+" Then
                    app = Left(app, Len(app) - 2) & ",plus"
                ElseIf right(app, 2) = "^-" Then
                    app = Left(app, Len(app) - 2) & ",minus"
                End If
                text = Left(text, pos - 1) & "limit(" & Expr & "," & var & "," & app & ")" & right(text, Len(text) - sluts)
            Else
                pos = pos + 3
            End If
        End If
    Loop While pos > 0


    'differentiering
    Dim subt As String, vararr() As String, s2 As String, f2 As String
    di = 0    ' hvis der er flere f'(2) skal der bruges flere variable
#If Mac Then
    text = Replace(text, "^('')", "^''")
    text = Replace(text, "^(''')", "^'''")
    text = Replace(text, "^('''')", "^''''")
#End If
    Do
        pos = InStr(text, "^'")
        If pos > 0 Then
            ea.text = text
            ea.pos = pos
            '        f = ea.ChrByIndex(pos - 1)
            f = ea.GetPrevVar(pos - 1)
            antalm = 0
            Do
                antalm = antalm + 1
            Loop Until ea.ChrByIndex(pos + 1 + antalm) <> "'"
            If ea.ChrByIndex(pos + antalm + 1) = " " Then    ' fjern evt. mellemrum efter '
                text = Left(text, pos + 1) & right(text, Len(text) - pos - 2)
                ea.text = text
            End If

            If ea.ChrByIndex(pos + antalm + 1) = "(" Then
                '            var = ea.ChrByIndex(pos + 3 + antalm)
                var = ea.GetNextBracketContent(pos + 1 + antalm)
                varl = Len(var) + 2
            Else
                If DiffEqVar = "" Then
                    var = "x"    ' hvis ingen variabel antages x
                Else
                    var = DiffEqVar    ' hvis diffligning så antages diffeqvar
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

            If ea.ChrByIndex(pos - 1) = ")" Then    ' hvis (...)'(x)
                ' find startparantes?
                udtryk = ea.GetPrevBracketContent(pos)
                If ea.ChrByIndex(ea.pos - 2) = "_" Then    '  f_(...)'(x)
                    f = ea.GetPrevVar(ea.pos - 1) & "(" & udtryk & ")"
                    GoTo msat
                    '                MsgBox "Der er sænket skrift som ikke kan fortolkes", vbOKOnly, "Fejl"
                    '                GoTo fejl
                ElseIf ea.IsLetterPos(ea.pos - 2) Then    ' sin(x)'
                    f = ea.GetPrevVar(ea.pos - 1)
                    text = Left(text, ea.pos - Len(f)) & "subst(" & var & ",qkqg" & di & ",diff(" & f & "(qkqg" & di & "),qkqg" & di & ",round(" & antalm & ")))" & right(text, Len(text) - pos - varl - antalm)
                    di = di + 1
                Else    ' (x^2)'
                    text = Left(ea.text, ea.pos - 2) & "diff(" & udtryk & "," & var & ",round(" & antalm & "))" & right(ea.text, Len(ea.text) - pos - varl - antalm)
                End If
            ElseIf ea.ChrByIndex(pos - 2 - Len(subt)) = ")" And Len(subt) > 0 And Len(subt) = antalm Then ' partiel (...)_xy''
                  s = ea.GetPrevBracketContent(pos - 2)
                  For i = 1 To antalm
                     s = "diff(" & s & "," & Mid(subt, i, 1) & ",1)"
                  Next
                 text = Left(ea.text, ea.pos - 2) & s & right(ea.text, Len(ea.text) - pos - varl - antalm)
            
            ElseIf Not ea.ChrByIndex(pos + antalm + 1) = "(" Then    ' hvis y'  variabel ikke funktion
               If Len(subt) > 0 And Len(subt) = antalm Then ' hvis partiel y'
                  s = Left(f, pos2 - 1)
                  For i = 1 To antalm
                     s = "diff(" & s & "," & Mid(subt, i, 1) & ",1)"
                  Next
                 text = Left(text, pos - 1 - Len(f)) & s & right(text, Len(text) - pos - varl - antalm)
               Else
                text = Left(text, pos - 1 - Len(f)) & "diff(" & f & "," & var & ",round(" & antalm & "))" & right(text, Len(text) - pos - varl - antalm)
               End If
            ElseIf ea.ChrByIndex(pos + antalm + 1) = "(" Then    ' hvis f'(x)
msat:
                If DiffEqVar = var Then    'hvis diffligning så ingen subst
                    text = Left(text, pos - 1 - Len(f)) & "diff(" & f & "(" & var & ")," & var & ",round(" & antalm & "))" & right(text, Len(text) - pos - varl - antalm)
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
                    text = Left(text, pos - 1 - Len(f)) & s & right(text, Len(text) - pos - varl - antalm)
                  Else ' alm ikke partiel f'(x)
                    text = Left(text, pos - 1 - Len(f)) & "subst(" & var & ",qkqg" & di & ",diff(" & f & "(qkqg" & di & "),qkqg" & di & ",round(" & antalm & ")))" & right(text, Len(text) - pos - varl - antalm)
                    di = di + 1
                  End If
                End If
            Else
                GoTo Fejl
            End If
        End If
    Loop While pos > 0

   ' nabla operator, grad og div
    Do
        pos = InStr(text, ChrW(8711) & ChrW(13))
        If pos > 0 Then
            ea.text = text
            ea.pos = pos
        End If
    Loop While pos > 0


    'integraler
    Dim integrand As String
    pos = InStr(text, VBA.ChrW(8747))
    If pos > 0 Then
        ea.text = text
        ea.text = Replace(ea.text, "diff", "qiff")    ' d'et vil give problemer da det vil blive opfattet som afslutning af integrale
        ea.text = Replace(ea.text, "round", "rounq")    ' d'et vil give problemer da det vil blive opfattet som afslutning af integrale
        i = 0
        Do
            ea.pos = 0
            ea.StartBracket = VBA.ChrW(8747)
            ea.EndBracket = VBA.ChrW(8518)    'differential d
            integrand = ea.GetNextBracketContent
            pos = ea.pos - Len(integrand)    'pos sættes til start-integraltegnet
            pos2 = ea.pos
            If Len(integrand) = 0 Then
                ea.EndBracket = "d"
                integrand = ea.GetNextBracketContent(pos)
                pos = ea.pos - Len(integrand) - 2    'pos sættes til start-integraltegnet
                pos2 = ea.pos
                If Not (ea.IsLetterPos(ea.pos, False) And Not (ea.IsLetterPos(ea.pos - 2, False))) Then ' hvis enligt d eller som af variabel/funktion der ikke starter med d, så er det ikke slutningen af integralet. Derfor forsøges med simplere metode til at finde d
                    j = 0
                    Do
                        j = j + 1
                        pos3 = InStr(ea.pos, ea.text, "d")
                        ea.pos = pos3 + 1
                    Loop While j < 100 And Not (ea.IsLetterPos(ea.pos, False) And Not (ea.IsLetterPos(ea.pos - 2, False)))
                    integrand = Mid(ea.text, pos + 1, pos3 - pos - 1)
                    pos2 = ea.pos
                End If
            End If
            var = ea.GetNextVar(ea.pos - 1)
            If Len(integrand) > 0 Then
                If Len(integrand) < 3 And Mid(ea.text, pos2 - Len(var), 2 + Len(var)) = "(" & var & ")" Then    ' hvis speciel notation med \int (dx)/(bla bla)
                    ea.SetNormalBrackets
                    integrand = ea.GetNextBracketContent(pos2 + 1)
                    pos2 = ea.pos - 1
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
                    poso = InStr(integrand, "^")
                    If posn > 0 And poso > 0 And posn < poss And poso < poss Then
                        nedre = Mid(integrand, posn + 1, poso - posn - 1)
                        upper = Mid(integrand, poso + 1, poss - poso - 1)
                        '        text = Left(text, posn - 1) & Mid(text, poss + 1, pos2 - poss - 1) & "," & Mid(text, pos2 + 1, 1) & "," & nedre & "," & upper & ")" & Mid(text, pos2 + 2, Len(text) - pos2 - 1)
                        integrand = right(integrand, Len(integrand) - poss) & "," & var & "," & nedre & "," & upper & ")"
                        If Left(integrand, 1) = VBA.ChrW(9618) Then    'fjern funktiontegn
                            integrand = right(integrand, Len(integrand) - 1)
                        End If
                        If MaximaExact = 2 Then
                            '                        ea.text = Replace(ea.text, VBA.ChrW(8747), "NIntegrate(", 1, 1)
                            ea.text = Left(ea.text, pos - 1) & "NIntegrate(" & integrand & right(ea.text, ea.Length - pos2 - Len(var) + 2)
                        Else
                            ea.text = Left(ea.text, pos - 1) & "Integrate(" & integrand & right(ea.text, ea.Length - pos2 - Len(var) + 2)
                            '                        ea.text = Replace(ea.text, VBA.ChrW(8747), "Integrate(", 1, 1)
                        End If
                    Else    ' ubestemt integrale
                        integrand = integrand & "," & var & ")"
                        If Left(integrand, 1) = VBA.ChrW(9618) Then    'fjern funktionstegn
                            integrand = right(integrand, Len(integrand) - 1)
                        End If
                        ea.text = Left(ea.text, pos - 1) & "integrate(" & integrand & right(ea.text, ea.Length - pos2 - Len(var) + 2)
                    End If
                    i = i + 1
                End If
            End If
        Loop While i < 100 And InStr(ea.text, VBA.ChrW(8747)) > 0 'Len(integrand) > 0
        ea.text = Replace(ea.text, "qiff", "diff")
        ea.text = Replace(ea.text, "rounq", "round")
        ea.SetNormalBrackets
        text = ea.text
    End If

    'dobbeltintegraler
    Dim var2 As String
    pos = InStr(text, VBA.ChrW(8748))
    If pos > 0 Then
        ea.text = text
        ea.text = Replace(ea.text, "diff", "qiff")    ' d'et vil give problemer da det vil blive opfattet som afslutning af integrale
        i = 0
        Do
            ea.pos = 0
            ea.StartBracket = VBA.ChrW(8748)
            ea.EndBracket = VBA.ChrW(8518)    'differential d
            integrand = ea.GetNextBracketContent
            pos = ea.pos - Len(integrand)    'pos sættes til start-integraltegnet
            pos2 = ea.pos
            If Len(integrand) = 0 Then
                ea.EndBracket = "d"
                integrand = ea.GetNextBracketContent
                pos = ea.pos - Len(integrand) - 2    'pos sættes til start-integraltegnet
                pos2 = ea.pos
            End If
            If Len(integrand) > 0 Then
                var = ea.GetNextVar(ea.pos - 1)
                If Len(var) < 2 Then
                    ConvertErrorText = Sprog.A(427)
                    GoTo Slut
                End If
                var = right(var, Len(var) - 1)
                var2 = ea.GetNextVar(ea.pos + 1)
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
                    poso = InStr(integrand, "^")
                    If posn > 0 And poso > 0 And posn < poss And poso < poss Then
                        ConvertErrorText = Sprog.A(429)
                        GoTo Slut
                    Else    ' ubestemt integrale
                        integrand = integrand & "," & var & ")," & var2 & ")"
                        If Left(integrand, 1) = VBA.ChrW(9618) Then    'fjern funktionstegn
                            integrand = right(integrand, Len(integrand) - 1)
                        End If
                        ea.text = Left(ea.text, pos - 1) & "integrate(integrate(" & integrand & right(ea.text, ea.Length - pos2 - Len(var) - Len(var2))
                    End If
                    i = i + 1
                End If
            End If
        Loop While i < 100 And Len(integrand) > 0
        ea.text = Replace(ea.text, "qiff", "diff")
        ea.SetNormalBrackets
        text = ea.text
    End If

    'trippelintegraler
    Dim var3 As String
    pos = InStr(text, VBA.ChrW(8749))
    If pos > 0 Then
        ea.text = text
        ea.text = Replace(ea.text, "diff", "qiff")    ' d'et vil give problemer da det vil blive opfattet som afslutning af integrale
        i = 0
        Do
            ea.pos = 0
            ea.StartBracket = VBA.ChrW(8749)
            ea.EndBracket = VBA.ChrW(8518)    'differential d
            integrand = ea.GetNextBracketContent()
            pos = ea.pos - Len(integrand)    'pos sættes til start-integraltegnet
            pos2 = ea.pos
            If Len(integrand) = 0 Then
                ea.EndBracket = "d"
                integrand = ea.GetNextBracketContent
                pos = ea.pos - Len(integrand) - 2    'pos sættes til start-integraltegnet
                pos2 = ea.pos
            End If
            If Len(integrand) > 0 Then
                var = ea.GetNextVar(ea.pos - 1)
                If Len(var) < 2 Then
                    ConvertErrorText = Sprog.A(427)
                    GoTo Slut
                End If
                var = right(var, Len(var) - 1)
                var2 = ea.GetNextVar(ea.pos + 1)
                If Len(var2) < 2 Or (Left(var2, 1) <> "d" And Left(var2, 1) <> VBA.ChrW(8518)) Then
                    ConvertErrorText = Sprog.A(428) '"Der er problemer med syntaxen. Der skal bruges to dx (eller lign) som afslutning af et dobbletintegrale. Måske har du brugt d i anden sammenhæng. Brug da differential d \dd"
                    GoTo Slut
                End If
                var2 = right(var2, Len(var2) - 1)
                var3 = ea.GetNextVar(ea.pos + 1)
                If Len(var3) < 2 Or (Left(var3, 1) <> "d" And Left(var3, 1) <> VBA.ChrW(8518)) Then
                    ConvertErrorText = Sprog.A(428) '"Der er problemer med syntaxen. Der skal bruges to dx (eller lign) som afslutning af et dobbletintegrale. Måske har du brugt d i anden sammenhæng. Brug da differential d \dd"
                    GoTo Slut
                End If
                var3 = right(var3, Len(var3) - 1)
                integrand = Trim(integrand)
                pos2 = pos2 + Len(var)
                If Len(integrand) > 0 And Len(var) > 0 And Len(var2) > 0 Then
                    IntegrationVars = IntegrationVars & var & ";"
                    poss = InStr(integrand, VBA.ChrW(9618))
                    posn = InStr(integrand, "_")
                    poso = InStr(integrand, "^")
                    If posn > 0 And poso > 0 And posn < poss And poso < poss Then
                        ConvertErrorText = Sprog.A(429) '"Du kan ikke bruge dobbeltintegraltegnet til bestemte integraler. Brug to enkeltintegraler."
                        GoTo Slut
                    Else    ' ubestemt integrale
                        integrand = integrand & "," & var & ")," & var2 & ")," & var3 & ")"
                        If Left(integrand, 1) = VBA.ChrW(9618) Then    'fjern funktionstegn
                            integrand = right(integrand, Len(integrand) - 1)
                        End If
                        ea.text = Left(ea.text, pos - 1) & "integrate(integrate(integrate(" & integrand & right(ea.text, ea.Length - pos2 - Len(var) - Len(var2) - Len(var3) - 2)
                    End If
                    i = i + 1
                End If
            End If
        Loop While i < 100 And Len(integrand) > 0
        ea.text = Replace(ea.text, "qiff", "diff")
        ea.SetNormalBrackets
        text = ea.text
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
    '                If ea.IsLetterPos(pos2 - 1, False) Then    ' hvis bogstav foran d, så indgår det i variabel
    '                    pos2 = pos2 + i
    '                    GoTo findslut
    '                End If
    '                i = 0
    '                Do
    '                    i = i + 1
    '                Loop While ea.IsLetterPos(pos2 + i, False) And i < 50
    '                var = Mid(text, pos2 + 1, i - 1)
    '                If var = "" Or var = "iff" Then    ' der vil være d hvis diff under integraltegn
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


    ' find underforståede paranteser efter ^ og / ' (skal være efter diff og andre funktioner med komma)
    ' flyttet op for at tage højde for skjulte parenteser
    ea.text = text
    ea.SetNormalBrackets
    ea.InsertBracketAfter ("^")
    text = ea.text

    ' konverter a_i - > a[i] a_i,j -> a[i,j]  hvis sat i indstillinger
    If MaximaIndex Then
        pos = 1
        Do
            pos = InStr(pos, ea.text, "_")
            If pos > 0 Then
                If ea.IsLetterPos(pos - 1, False) Or ea.ChrByIndex(pos - 1) = ")" Then
                    ea.RemoveChar pos
                    ea.pos = pos
                    If ea.ChrByIndex(pos) = "(" Then
                        Expr = ea.GetNextBracketContent()
                        ea.text = Left(ea.text, pos - 1) & "[" & Expr & "]" & right(ea.text, Len(ea.text) - pos - Len(Expr) - 1)
                    Else
                        posb = ea.FindHierakiBreak(5)    ' ændret 5 pga f_1^'(x)
                        Expr = ea.GetTextInterval(pos, posb)
                        ea.text = Left(ea.text, pos - 1) & "[" & Expr & "]" & right(ea.text, Len(ea.text) - pos - Len(Expr) + 1)
                    End If
                Else
                    pos = pos + 1
                End If
            End If
        Loop While pos > 0
    End If
spring:
    ea.RemoveSpaceAfterSubSuperscript ("_")
    '    ea.RemoveSpaceAfterSubSuperscript ("^")
    ' fjern dobbelt mellemrum
    Do While InStr(ea.text, "  ") > 0
        ea.text = Replace(ea.text, "  ", " ")
    Loop
    ea.ReplaceSpaceByMult    ' erstatter mellemrum med * men ikke hvis ulovlig
    text = ea.text

    text = Replace(text, " ", "")     ' evt. resterende mellemrum fjernes
    '    text = Replace(text, " ", "*")     ' skal være efter diff og integral ' problematisk ved (x^2 )' -> diff(x^2*,x,1)


    text = Replace(text, "det(", "determinant2(")    ' lidt usikker da der kan være funktioner der slutter på "det"
    ' Matricer
    Do
        pos = InStr(text, VBA.ChrW(9632))
        If pos > 0 Then
            matvekt = True
            stpos = pos - 1
            ea.text = text
            ea.pos = pos
            matrix = ea.GetNextBracketContent
            slpos = ea.pos
            If ea.ChrByIndex(slpos + 1) = "*" Then
                text = Left(text, slpos) & " ## " & right(text, Len(text) - slpos - 1)
            End If
            arr = Split(matrix, "@")
            mtext = "matrix("
            For i = 0 To UBound(arr)
                mtext = mtext & "[" & Replace(arr(i), "&", ",") & "],"
            Next
            mtext = Left(mtext, Len(mtext) - 1) & ")"
            text = Left(text, stpos - 1) & mtext & right(text, Len(text) - slpos)
        End If
    Loop While pos > 0

    'vektorer
    Do
        pos = InStr(text, VBA.ChrW(9608))
        If pos > 0 Then
            matvekt = True
            stpos = pos - 1
            ea.text = text
            ea.pos = pos
            matrix = ea.GetNextBracketContent
            slpos = ea.pos
            arr = Split(matrix, "@")
            mtext = "matrix("
            For i = 0 To UBound(arr)
                mtext = mtext & "[" & arr(i) & "],"
            Next
            mtext = Left(mtext, Len(mtext) - 1) & ")"
            text = Left(text, stpos - 1) & mtext & right(text, Len(text) - slpos)
        End If
    Loop While pos > 0




    ' indsæt underforståede gangetegn mellem bogstav mellemrum og bogstav
    ' virker men så kan man ikke lave funktioner med sænket skrift da der da er mellemrum
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


    ' indsæt underforståede gangetegn ' skal være efter fjern mellemrum
    ea.text = text
    ea.pos = 1
    ea.InsertMultSigns
    text = ea.text
    If Radians Then
        text = Replace(text, VBA.ChrW(8736), " CVinkel ")    ' hvis complex så er vinkel polær notation
    Else
        text = Replace(text, VBA.ChrW(8736), " CVinkelgrad ")    ' hvis complex så er vinkel polær notation
    End If

    text = Replace(text, "gradq*", "gradq ")
    text = Replace(text, "divq*", "divq ")
    text = Replace(text, "curlq*", "curlq ")
    text = Replace(text, "laplaceq*", "laplaceq ")


    If matvekt Then
        '        text = Replace(text, "])^", "])^^")  ' matrix ^ enkelt ^ opløfter kun hvert enkelt element
        text = Replace(text, "^", "^^")  ' definerede matricer blev ikke håndteret korrekt med ovenstående. Forsøger nu at erstatte alle og håber det ikke giver andre problemer
'#If Mac Then ' from 1.18  windows is also sbcl
        text = Replace(text, "*", " ## ")    ' egentlig kun vigtig ved matrix gangetegn og problematisk når * ikke er matrix multiplikation, da der så er problemer med hieraki 2/3 . matrix([2],[3]) her ganges 3 ind først. Dog løst med egen ¤ operator istedet for .
        text = Replace(text, "§", " ## ")    ' egentlig kun vigtig ved matrix gangetegn og problematisk når * ikke er matrix multiplikation, da der så er problemer med hieraki 2/3 . matrix([2],[3]) her ganges 3 ind først. Dog løst med egen ¤ operator istedet for .
'#Else
'        text = Replace(text, "*", " ¤ ")    ' egentlig kun vigtig ved matrix gangetegn og problematisk når * ikke er matrix multiplikation, da der så er problemer med hieraki 2/3 . matrix([2],[3]) her ganges 3 ind først. Dog løst med egen ¤ operator istedet for .
'#End If
        '        text = Replace(text, "*matrix", " . matrix")  ' matrix gangetegn og prikprodukt
        '        text = Replace(text, "*vect", " . vect")
    End If

    'erstat e med %e og i med %i      er flyttet til efter underforståede gangetegn da % ikke gælder som bogstav
    text = CodeConstantPercentSign(text, "e")
    If MaximaComplex Then text = CodeConstantPercentSign(text, "i")


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
    ea.text = text
    If ConvertLnLog Then
        ea.ReplaceVar "ln", "log"
    End If
    ea.ReplaceVar "arg", "carg"
    text = ea.text


    ' trig funktioner konverterer til radianer
    If Not Radians Then
        text = ConvertDegreeToRad(text, "sin")
        text = ConvertDegreeToRad(text, "cos")
        text = ConvertDegreeToRad(text, "tan")
        text = ConvertDegreeToRad(text, "sec")
        text = ConvertDegreeToRad(text, "cot")
        text = ConvertDegreeToRad(text, "csc")
    End If

    ' her var æøå og græske før

    text = Replace(text, VBA.ChrW(960), "%pi")
    text = Replace(text, "-" & VBA.ChrW(8734), "minf")
    text = Replace(text, VBA.ChrW(8734), "inf")


    'enheder
    If MaximaUnits Then
        text = ConvertUnits(text)
    Else
        text = Replace(text, VBA.ChrW(937), "Omega")    ' omega
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
        CodeForMaxima = text
    End If
End Function
Function ConvertDifferentials(text As String, DiffChr As String) As String
    ' d/dx notation
    ' konverterer d/dx til maxima diff(...,x) notation
    ' DiffChr er symbolet for d. Det kan fx være VBA.ChrW(8518) (\dd) eller VBA.ChrW(...)(\partial) eller måske et almindeligt d
    
    Dim num As String, dem As String, vararr(9) As String, varind As Integer, antalm2 As Integer
    Dim udtryk1 As String, udtryk2 As String, pos As Integer, pos3 As Integer, antalm As Integer
    Dim pos2 As Integer, var As String, i As Integer, udtryk As String, posstart As Integer, posslut As Integer
    Dim ea As New ExpressionAnalyser
    ea.SetNormalBrackets
    Do
        pos = InStr(text, "/(" & DiffChr)
        If pos > 0 Then
            ea.text = text
            dem = ea.GetNextBracketContent(pos + 1)
            pos3 = ea.pos
            num = ea.GetPrevBracketContent(pos)
            If Left(num, 1) <> DiffChr Then
                GoTo Slut
            End If
            pos = ea.pos
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
                    text = Sprog.A(430)
                    GoTo Slut
                End If
                ElseIf Left(num, 1) = DiffChr Then
                    antalm = 1
                    vararr(0) = Mid(dem, 2, 1)
                Else
                    text = Sprog.A(430)
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
                text = Left(text, pos - 2) & udtryk1 & udtryk & udtryk2 & right(text, Len(text) - pos3 + 1)
            ElseIf Mid(text, pos3 + 1, 1) = "(" Then ' parentes lige efter d/dx operator
                udtryk = ea.GetNextBracketContent(pos3)
                text = Left(text, pos - 2) & udtryk1 & udtryk & udtryk2 & right(text, Len(text) - ea.pos + 1)
                ' Her mangler notation for df(x)/dx
            Else ' ingen parentes efter d/dx operator så udtrykket må findes via hieraki
                posstart = pos3
                ea.pos = pos3 + 1
                posslut = ea.FindHierakiBreak(2)
                udtryk = Mid(text, pos3, posslut - posstart + 1)
                text = Left(text, pos - 2) & udtryk1 & udtryk & udtryk2 & right(text, Len(text) - posslut)
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
    Loop While pos > 0
Slut:
    ConvertDifferentials = text
End Function
Function ConvertUnits(ByVal text As String) As String
    Dim ea As New ExpressionAnalyser
    text = Replace(text, VBA.ChrW(937), "Ohm")    ' Omega er så ohm
    ea.text = text
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
    ConvertUnits = ea.text

End Function
Function ConvertDegreeToRad(text As String, trigfunc As String) As String
    Dim pos, spos As Integer
    Dim ea As New ExpressionAnalyser
    ea.StartBracket = "("
    ea.EndBracket = ")"
    ea.text = text
    spos = 1

    Do
        pos = ea.FindChr("a" & trigfunc, spos)
        If pos > 0 Then
            If (pos = 1 Or Not (ea.IsLetterPos(pos - 1, False))) And Not (ea.IsLetterPos(pos + Len(trigfunc) + 1, False)) Then
                ea.GetNextBracketContent
                ea.InsertBeforePos (")")
                ea.pos = pos
                ea.InsertBeforePos ("180/%pi*(")
                spos = pos + 13
            Else
                spos = pos + 3
            End If
        End If
    Loop While pos > 0

    spos = 1
    Do
        pos = ea.FindChr(trigfunc, spos)
        If pos > 0 Then
            If (pos = 1 Or Not (ea.IsLetterPos(pos - 1, False))) And Not (ea.IsLetterPos(pos + Len(trigfunc), False)) Then
                'If Not (ea.ChrByIndex(pos - 1) = "a") Then
                ea.GetNextBracketContent
                ea.InsertBeforePos (")")
                ea.pos = pos + Len(trigfunc)
                ea.InsertAfterPos ("%pi/180*(")
                spos = pos + 13
            Else
                spos = pos + 3
            End If
        End If
    Loop While pos > 0

    ConvertDegreeToRad = ea.text

End Function

Function CodeConstantPercentSign(text As String, tegn As String) As String
' indsætter % foran konstant i udtryk hvis der ikke er bogstav foran eller bag
    Dim pos As Integer
    Dim ea As New ExpressionAnalyser

    ea.text = text
    pos = 1
    Do
        pos = InStr(pos, ea.text, tegn)
        If pos > 0 Then
            If Not (ea.IsLetterPos(pos + 1, False) Or ea.ChrByIndex(pos + 1) = "[" Or ea.IsLetterPos(pos - 1, False)) Then
                ea.text = Left(ea.text, pos - 1) & "%" & right(ea.text, Len(ea.text) - pos + 1)
                pos = pos + 2
            Else
                pos = pos + 1
            End If
        End If
    Loop While pos > 0

    CodeConstantPercentSign = ea.text
End Function
Function ConvertToAscii(ByVal text As String) As String
' konverterer fra word input til ascii

    text = Replace(text, VBA.ChrW(12310), "")    ' specielle usynlige paranteser fjernes
    text = Replace(text, VBA.ChrW(12311), "")    ' specielle usynlige paranteser fjernes
    text = Replace(text, VBA.ChrW(9633), "")    ' tegn foran paranteser der gør dem usynlige. Muligvis problematisk da paranteserne ikke fjernes. gør også brøker små
    text = Replace(text, VBA.ChrW(9508), "")    ' tegn for tom plads ved ihvertfald [ ]
    text = Replace(text, VBA.ChrW(11), "")    ' shift-enter
    text = Replace(text, vbLf, "")    ' shift-enter og enter
    text = Replace(text, vbCrLf, "")
    text = Replace(text, vbCr, "")
    text = Replace(text, """", "")    ' apostrof fjernes
'#If Mac Then
'#Else
    text = Replace(text, "^'", "´")    ' alm apostrof kan ikke bruges
    text = Replace(text, "'", "´")    ' alm apostrof kan ikke bruges
'#End If
    text = Replace(text, "+-", "-")    ' plus-minus giver minus
    text = Replace(text, VBA.ChrW(8289), "")    ' symbol der definerer funktion fjernes fra word syntaks
    text = Replace(text, VBA.ChrW(8212), "+")    'dobbbelt minustegn giver plus
    '    text = Replace(text, VBA.ChrW(183), "*") ' prik erstattes med gange
    '    text = Replace(text, VBA.ChrW(215), "*") ' kryds erstattes med gange
    '    text = Replace(text, VBA.ChrW(8901), "*") ' \cdot , ved ikke hvorfor der tilsyneladende er to
    '    text = Replace(text, VBA.ChrW(8729), "*") ' prik \cdot erstattes med gange
    '    text = Replace(text, VBA.ChrW(8226), "*") ' tyk prik erstattes med gange
    '    text = Replace(text, VBA.ChrW(180), "^'") ' der er tre forskellige '´` apostrof lignende de forståes alle som diff
    '    text = Replace(text, VBA.ChrW(96), "^'") ' der er tre forskellige '´`

    text = Replace(text, ChrW(183), "*")  ' Siden 1.18 (SBCL) kan alm gangetegn ikke vises
    text = Replace(text, ChrW(215), "*")  '
    text = Replace(text, "  ", " ")
    '    text = Replace(text, " ", "") ' mellemrum fjernes ' problematisk når , bruges som ; separator f(0 , 1) -> f(0,1)
    text = Replace(text, """", "")    ' apostrof fjernes
    text = Replace(text, "+-", "-")    ' plus-minus giver minus
    text = Replace(text, "%e", "e")    '
    '    text = Replace(text, "æ", "ae") '
    '    text = Replace(text, "ø", "oe") '
    '    text = Replace(text, "å", "aa") '


    ConvertToAscii = text
End Function
Function ConvertToWordSymbols(ByVal text As String) As String
    On Error GoTo Fejl
    Dim gtegn As String
    Dim ea As New ExpressionAnalyser
    Dim pos As Integer
    Dim udtryk As String
    Dim possep As Integer
    Dim arr As Variant
    Dim mtext As String
    Dim slutp As Integer
    Dim s As String, s2 As String

    ea.SetNormalBrackets
    gtegn = MaximaGangeTegn

    text = Replace(text, " . ", "*")    ' prikprodukt
    text = Replace(text, " ", "")    ' mellemrum fjernes
    text = Replace(text, """", "")    ' apostrof fjernes
    text = Replace(text, "+-", "-")    ' plus-minus giver minus
    text = Replace(text, VBA.ChrW(11), "")    ' shift-enter
    text = Replace(text, vbCrLf, "")    ' enter
    text = Replace(text, VBA.ChrW(10), "")    ' shift-enter
    text = Replace(text, VBA.ChrW(13), "")    ' shift-enter
    text = Replace(text, "'lgog", "lgog")    '


    ea.text = text

    ea.RemoveFaultySeparator


    ' fx e^x^-1 rettes
    pos = InStr(ea.text, "^-")
    Do While pos > 0
        ea.InsertTextAt "(", pos + 1
        ea.pos = pos + 3
        ea.FindHierakiBreak 3
        ea.InsertTextAt ")", ea.pos + 1
        pos = InStr(pos + 1, ea.text, "^-")
    Loop


    ea.InsertBracketAfter ("^")


    ' eksponentiel notation positiv   ' før insertbracketbefore da ellers 4E-13/x -> 4E-(13)/x men c^5.0E-1 -> c^0.5*10^-1 duer ikke
    pos = 0
    Do
'#If Mac Then ' SBCL
        pos = InStr(pos + 1, ea.text, "e+")
'#Else 'GCL
'        pos = InStr(pos + 1, ea.text, "E+")
'#End If
        If pos > 1 Then
            If ea.IsNumber(pos - 1) Then
                ea.pos = pos + 1
                Do
                    ea.pos = ea.pos + 1
                Loop While ea.IsNumber(ea.pos)
                ' mellemrum fjernet for at få insertbracket til at virke
                ea.text = Left(ea.text, pos - 1) & gtegn & "10^" & Mid(ea.text, pos + 2, ea.pos - pos - 2) & " " & right(ea.text, Len(ea.text) - ea.pos + 1)
                'ea.text = Left(ea.text, pos - 1) & gtegn & "10^" & Mid(ea.text, pos + 2, ea.pos - pos - 2) & Right(text, Len(ea.text) - ea.pos + 1)
            End If
        End If
    Loop While pos > 0

    ' eksponentiel notation negativ
    pos = 0
    Do
'#If Mac Then ' SBCL
        pos = InStr(pos + 1, ea.text, "e-")
'#Else 'GCL
'        pos = InStr(pos + 1, ea.text, "E-")
'#End If
        If pos > 1 Then
            If ea.IsNumber(pos - 1) Then
                ea.pos = pos + 1
                Do
                    ea.pos = ea.pos + 1
                Loop While ea.IsNumber(ea.pos)
                ea.text = Left(ea.text, pos - 1) & gtegn & "10^" & Mid(ea.text, pos + 1, ea.pos - pos - 1) & " " & right(ea.text, Len(ea.text) - ea.pos + 1)
                '    ea.text = Left(ea.text, pos - 1) & gtegn & "10^" & Mid(ea.text, pos + 1, ea.pos - pos - 1) & Right(ea.text, Len(ea.text) - ea.pos + 1)
            End If
        End If
    Loop While pos > 1

    ' efter eksponentiel da 1E-1*x går galt
    ' 1* fjernes
    pos = InStr(ea.text, "1*")
    Do While pos > 0
        If pos = 1 Then
            ea.text = Left(ea.text, pos - 1) & right(ea.text, Len(ea.text) - pos - 1)    ' fjern 1*
        ElseIf Mid(ea.text, pos - 1, 1) = "*" Then  ' gange foran så fjern altid
            ea.text = Left(ea.text, pos - 1) & right(ea.text, Len(ea.text) - pos - 1)    ' fjern 1*
        ElseIf Mid(ea.text, pos - 1, 1) = "(" Then  ' parentes foran så fjern altid
            ea.text = Left(ea.text, pos - 1) & right(ea.text, Len(ea.text) - pos - 1)    ' fjern 1*
        ElseIf Mid(ea.text, pos - 1, 1) = ")" Then  ' parentes foran så fjern altid
            ea.text = Left(ea.text, pos - 1) & right(ea.text, Len(ea.text) - pos - 1)    ' fjern 1*
        ElseIf Mid(ea.text, pos - 1, 1) = "=" Then  ' ligmed foran så fjern altid
            ea.text = Left(ea.text, pos - 1) & right(ea.text, Len(ea.text) - pos - 1)    ' fjern 1*
        ElseIf Mid(ea.text, pos - 1, 1) = VBA.ChrW(8776) Then  ' ca. ligmed foran så fjern altid
            ea.text = Left(ea.text, pos - 1) & right(ea.text, Len(ea.text) - pos - 1)    ' fjern 1*
        ElseIf ea.IsNumber(pos - 1) Then    ' tal foran så ingenting

        ElseIf ea.IsLetterPos(pos - 1, True) Then    ' bogstav foran så ingenting

        ElseIf Mid(ea.text, pos - 1, 1) = "." Then    ' .1* så ingenting

        ElseIf pos > 2 Then
            If Mid(ea.text, pos - 2, 2) = "^-" Then    ' ^-1 fjern ikke
            Else
                ea.text = Left(ea.text, pos - 1) & right(ea.text, Len(ea.text) - pos - 1)    ' fjern 1*
            End If
        Else
            ea.text = Left(ea.text, pos - 1) & right(ea.text, Len(ea.text) - pos - 1)    ' fjern 1*
        End If
        pos = InStr(pos + 1, ea.text, "1*")
    Loop

    ' 1,0* fjernes og ,0* ->*
    pos = InStr(ea.text, ".0*")
    Do While pos > 0
        If ea.ChrByIndex(pos - 1) = "1" Then
            If Not (ea.IsNumber(pos - 2)) Or pos = 2 Then
                ea.text = Left(ea.text, pos - 2) & right(ea.text, Len(ea.text) - pos - 2)
            Else
                ea.text = Left(ea.text, pos - 1) & right(ea.text, Len(ea.text) - pos - 1)    ',0* ->*
            End If
        Else
            ea.text = Left(ea.text, pos - 1) & right(ea.text, Len(ea.text) - pos - 1)    ',0* ->*
        End If
        pos = InStr(pos + 1, ea.text, ".0*")
    Loop
    ' tal,ikketal fjernes  f.eks. 240, -> 240  samt x*,45  ->x*0,45
    ' test: 2,0·10^3·g
    
'#If Mac Then
'** MAC   .0 på alle tal, fordi der bruges float istedet for dectalallnum
    pos = InStr(ea.text, ".0")
    Do While pos > 0
            If Not (ea.IsNumber(pos + 2)) Then
                ea.text = Left(ea.text, pos - 1) & right(ea.text, Len(ea.text) - pos - 1)
            End If
        pos = InStr(pos + 1, ea.text, ".0")
    Loop
'#End If

    ' konverter a[i,j] -> a_(i,j)
    If MaximaIndex Then
        pos = 1
        Do
            pos = InStr(pos + 1, ea.text, "[")
            If pos > 1 And ea.IsLetterPos(pos - 1, False) Then
                ea.pos = pos
                ea.SetSquareBrackets
                udtryk = ea.GetNextBracketContent(pos)
                '       if ea.ChrByIndex(ea.pos)=
                ea.text = Left(ea.text, pos - 1) & "_(" & udtryk & ")" & right(ea.text, Len(ea.text) - ea.pos + 1)

            End If
        Loop While pos > 1
        ea.SetNormalBrackets
    End If


    ' fjern gangetegn der er underforståede
    ' giver desværre problemer bla. foran inverse trig  180sin^-1(12)/\pi og (1+x)=2 sin^(-1)(1/2)/\pi konvertes forkert til prof
    ' virker hverken hvis * erstattes med mellemrum eller fjernes
    ' rækkefølgen på disse to skal muligvis byttes om da det giver problemer ved identifikation af identitiske udtryk for beregn auto

    '    ea.InsertBracketBefore ("/") ' skal bruges for at removemultsign virker, men problem ved f.eks. matrix([5/6];[13/10]) -> matrix([5/6];([13)/10])
    If MaximaUnits Then ' version 1.07 forsøges når enheder er slået til
        ea.RemoveMultSign ' stadig problematisk ved f.eks. e^-x^2*(e^(x^2)+2·c)/2
    End If
    ea.AddSpaceAfterSubSuperscript ("_")
    text = ea.text

    If Not MaximaSeparator Then
        text = Replace(text, ",", ";")
        text = Replace(text, ".", ",")
    End If
    text = Replace(text, "%pi", VBA.ChrW(960))
    text = Replace(text, "%e", "e")
    text = Replace(text, "%c", "c")
    text = Replace(text, "%k", "c_")
    text = Replace(text, "%r", "k_")
    text = Replace(text, "%z", VBA.ChrW(8484) & "_")
    text = Replace(text, "%Z", VBA.ChrW(8484))
    text = Replace(text, "#", VBA.ChrW(8800))    ' ikke ligmed

    If MaximaUnits Then
        text = Replace(text, VBA.ChrW(937), "Ohm")    ' Omega er så ohm
        text = Replace(text, "%%m", "m")
        text = Replace(text, "%mu", VBA.ChrW(956))
        ea.text = text
        ea.ReplaceVar "degC", VBA.ChrW(176) & "C"  ' grad tegn og C
        ea.ReplaceVar "%timer", "timer"
        ea.ReplaceVar "%min", "min"
        ea.ReplaceVar "aar", "år"
        text = ea.text
    End If


    text = Replace(text, "*", gtegn)    ' VBA.ChrW(183)


    '    Text = Replace(Text, "e", VBA.ChrW(183) & "10^ ")
    text = Replace(text, "exp(", "e^(")
    text = Replace(text, "log", "ln" & VBA.ChrW(8289))
    text = Replace(text, "lgog(", "log" & VBA.ChrW(8289) & "(")
    text = Replace(text, "sqrt", VBA.ChrW(8730))

    '    text = Replace(text, "E+", gtegn & "10^")
    '    text = Replace(text, "E-", gtegn & "10^-")


    ' notequal(x,y) -> x=y
    Do
        pos = InStr(text, "notequal(")
        If pos > 0 Then
            ea.text = text
            udtryk = ea.GetNextBracketContent(pos + 7)
            text = Left(text, pos - 1) & " " & Replace(udtryk, ListSeparator, VBA.ChrW(8800)) & " " & right(text, Len(text) - ea.pos + 1)
        End If
    Loop While pos > 0

    ' equal(x,y) -> x=y
    Do
        pos = InStr(text, "equal(")
        If pos > 0 Then
            ea.text = text
            udtryk = ea.GetNextBracketContent(pos + 4)
            text = Left(text, pos - 1) & " " & Replace(udtryk, ListSeparator, "=") & " " & right(text, Len(text) - ea.pos + 1)
        End If
    Loop While pos > 0

    '    text = Replace(text, "if", vbCrLf & "hvis  ")
    '    text = Replace(text, "then", "  så  ")
    '    text = Replace(text, "else", vbCrLf & "ellers  ")

    ' abs(x) -> |x|
    Do
        pos = InStr(text, "abs(")
        If pos > 0 Then
            ea.text = text
            udtryk = ea.GetNextBracketContent(pos)
            text = Left(text, pos - 1) & VBA.ChrW(124) & udtryk & VBA.ChrW(124) & right(text, Len(text) - ea.pos + 1)
        End If
    Loop While pos > 0

    ' 'vecmag(x) -> |x|
    Do
        pos = InStr(text, "'vecmag(")
        If pos > 0 Then
            ea.text = text
            udtryk = ea.GetNextBracketContent(pos + 6)
            text = Left(text, pos - 1) & VBA.ChrW(124) & udtryk & VBA.ChrW(124) & right(text, Len(text) - ea.pos + 1)
        End If
    Loop While pos > 0

    'atan2(x,y) til atan(x/y)
    pos = InStr(text, "atan2")
    Do While pos > 0
        ea.text = text
        ea.GetNextBracketContent (pos)
        text = Left(text, ea.pos - 1) & ")" & right(text, Len(text) - ea.pos + 1)    ' slutparantes i nævner
        text = Left(text, pos + 3) & "(" & right(text, Len(text) - pos - 4)    ' , atan2 til atan(
        possep = InStr(pos + 4, text, ListSeparator)
        text = Left(text, possep - 1) & ")/(" & right(text, Len(text) - possep)    ' , til /

        pos = InStr(pos, text, "atan2")
    Loop

    ' integrate til integraltegn
    pos = InStr(text, "integrate")
    Do While pos > 0
        If Mid(text, pos - 1, 1) = "'" Then
            text = Left(text, pos - 2) & right(text, Len(text) - pos + 1)
            pos = pos - 1
        End If
        Dim udt As String
        ea.text = text
        ea.pos = pos + 8
        udt = ea.GetNextBracketContent()
        arr = Split(udt, ListSeparator)
        If UBound(arr) > 1 Then
            text = Left(text, pos - 1) & VBA.ChrW(8747) & "_" & arr(2) & "^" & arr(3) & VBA.ChrW(9618) & arr(0) & "  d" & arr(1) & " " & right(text, Len(text) - pos - Len(udt) - 10)
        Else
            text = Left(text, pos - 1) & VBA.ChrW(8747) & VBA.ChrW(9618) & arr(0) & "  d" & arr(1) & " " & right(text, Len(text) - pos - Len(udt) - 10)
        End If
        pos = InStr(pos, text, "integrate")
    Loop

    ' differentier
    pos = InStr(text, "diff")
    Do While pos > 0
        If Mid(text, pos - 1, 1) = "'" Then
            text = Left(text, pos - 2) & right(text, Len(text) - pos + 1)
            pos = pos - 1
        End If
        ea.text = text
        ea.pos = pos + 4
        udt = ea.GetNextBracketContent()
        arr = Split(udt, ListSeparator)
        If UBound(arr) > 1 Then
            If arr(2) = "1" Then
                text = Left(text, pos - 1) & VBA.ChrW(8518) & "/(" & VBA.ChrW(8518) & arr(1) & ") (" & arr(0) & ")" & " " & right(text, Len(text) - Len(udt) - 5 - pos)
            Else
                text = Left(text, pos - 1) & VBA.ChrW(8518) & "/(" & VBA.ChrW(8518) & "^" & arr(2) & " " & arr(1) & ") (" & arr(0) & ")" & " " & right(text, Len(text) - Len(udt) - 6)
            End If
        Else
        End If
        pos = InStr(pos, text, "diff")
    Loop

    ' Limit
    pos = InStr(text, "limit")
    Do While pos > 0
        If Mid(text, pos - 1, 1) = "'" Then
            text = Left(text, pos - 2) & right(text, Len(text) - pos + 1)
            pos = pos - 1
        End If
        ea.text = text
        ea.pos = pos + 4
        udt = ea.GetNextBracketContent()
        arr = Split(udt, ListSeparator)
        'lim" & VBA.ChrW(9516) & "(x" & VBA.ChrW(8594) & "1)" & VBA.ChrW(8289) & "(x+1)
        text = Left(text, pos - 1) & "lim" & VBA.ChrW(9516) & "(" & arr(1) & VBA.ChrW(8594) & arr(2) & ")" & VBA.ChrW(8289) & "(" & arr(0) & ") " & right(text, Len(text) - pos - Len(udt) - 6)
        pos = InStr(pos, text, "limit")
    Loop


    If Not Radians Then
        text = Replace(text, "sing(", "sin(")
        text = Replace(text, "asing(", "asin(")
        text = Replace(text, "cosg(", "cos(")
        text = Replace(text, "acosg(", "acos(")
        text = Replace(text, "tang(", "tan(")
        text = Replace(text, "atang(", "atan(")

        '        text = ConvertRadToDegree(text, "sin")
        '        text = ConvertRadToDegree(text, "cos")
        '        text = ConvertRadToDegree(text, "tan")
        text = ConvertRadToDegree(text, "sec")
        text = ConvertRadToDegree(text, "cot")
        text = ConvertRadToDegree(text, "csc")
    End If

    '    text = Replace(text, "acos", "cos^(-1) ")
    '    text = Replace(text, "asin", "sin^(-1) ")
    '    text = Replace(text, "atan", "tan^(-1) ")
    ea.text = text
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
    text = ea.text

    '        text = Replace(text, "acos(", "cos^(-1)" & VBA.ChrW(8289) & "(") ' parentes forsøg ved invers også
    '        text = Replace(text, "asin(", "sin^(-1)" & VBA.ChrW(8289) & "(")
    '        text = Replace(text, "atan(", "tan^(-1)" & VBA.ChrW(8289) & "(")
    '        text = Replace(text, "cos(", "cos" & VBA.ChrW(8289) & "(")
    '        text = Replace(text, "sin(", "sin" & VBA.ChrW(8289) & "(")
    '        text = Replace(text, "tan(", "tan" & VBA.ChrW(8289) & "(")



    ' matricer
    pos = InStr(text, "matrix")
    Do While pos > 0

        ea.text = text
        mtext = ea.GetNextBracketContent(pos + 5)
        mtext = Replace(mtext, "]" & ListSeparator & "[", "@")
        mtext = Replace(mtext, ListSeparator, "&")

        mtext = matrixstartbracket & VBA.ChrW(9632) & "(" & Mid(mtext, 2, Len(mtext) - 2) & ")" & matrixendbracket
        text = Left(text, pos - 1) & mtext & right(text, Len(text) - ea.pos + 1)

        pos = InStr(pos + 1, text, "matrix")
    Loop

    ' 'and i %if
    pos = InStr(text, "'and(")
    If pos > 0 Then
        ea.text = text
        ea.pos = pos + 3
        udt = ea.GetNextBracketContent()
        udt = Replace(udt, ListSeparator, "  " & VBA.ChrW(8743) & "  ")
        text = Left(text, pos - 1) & udt & right(text, Len(text) - ea.pos + 1)
    End If

    ' %if
    pos = InStr(text, "%if")
    Do While pos > 0
        ea.text = text
        mtext = ea.GetNextBracketContent(pos + 2)
        slutp = ea.pos
        arr = Split(mtext, ListSeparator)
        text = Left(text, pos - 1) & arr(1) & "   for   " & arr(0) & right(text, Len(text) - slutp + 1)
        pos = InStr(text, "%if")
    Loop

    'vecta_ -> vektorpil
    pos = InStr(text, "vecta_")
    Do While pos > 0
        ea.text = text
        s = ea.GetNextVar(pos + 6)
        text = Left(text, pos - 1) & "(" & s & ")" & VBA.ChrW(160) & VBA.ChrW(8407) & right(text, Len(text) - ea.pos)
        pos = InStr(text, "vecta_")
    Loop
    'vectb_ -> ikke hel vektorpil
    pos = InStr(text, "vectb_")
    Do While pos > 0
        ea.text = text
        s = ea.GetNextVar(pos + 6)
        text = Left(text, pos - 1) & "(" & s & ")" & VBA.ChrW(160) & VBA.ChrW(8401) & right(text, Len(text) - ea.pos)
        pos = InStr(text, "vectb_")
    Loop
    'vectc_ -> matrixstreg
    pos = InStr(text, "vectc_")
    Do While pos > 0
        ea.text = text
        s = ea.GetNextVar(pos + 6)
        text = Left(text, pos - 1) & "(" & s & ")" & VBA.ChrW(160) & VBA.ChrW(773) & right(text, Len(text) - ea.pos)
        pos = InStr(text, "vectc_")
    Loop

    'græske bogstaver
    text = Replace(text, "Delta", VBA.ChrW(916))
    text = Replace(text, "delta", VBA.ChrW(948))
    text = Replace(text, "alpha", VBA.ChrW(945))
    text = Replace(text, "beta", VBA.ChrW(946))
    text = Replace(text, "gamma_incomplete", "\Gamma_incomplete ")
    text = Replace(text, "gamma", VBA.ChrW(915))
    text = Replace(text, "gammaLB", VBA.ChrW(947))
    text = Replace(text, "theta", VBA.ChrW(952))
    text = Replace(text, "Theta", VBA.ChrW(920))
    text = Replace(text, "lambda", VBA.ChrW(955))
    text = Replace(text, "Lambda", VBA.ChrW(923))
    text = Replace(text, "rho", VBA.ChrW(961))
    text = Replace(text, "varphi", VBA.ChrW(966))
    text = Replace(text, "phi", VBA.ChrW(981))
    text = Replace(text, "Phi", VBA.ChrW(934))
    text = Replace(text, "varepsilon", VBA.ChrW(949))
    text = Replace(text, "epsilon", VBA.ChrW(1013))
    text = Replace(text, "psi", VBA.ChrW(968))
    text = Replace(text, "Psi", VBA.ChrW(936))
    text = Replace(text, "sigma", VBA.ChrW(963))
    text = Replace(text, "Sigma", VBA.ChrW(931))
    text = Replace(text, "mu", VBA.ChrW(956))
    text = Replace(text, "Ohm", VBA.ChrW(937))
    text = Replace(text, "Omega", VBA.ChrW(937))
    text = Replace(text, "omega", VBA.ChrW(969))
    text = Replace(text, "Xi", VBA.ChrW(926))
    text = Replace(text, "xi", VBA.ChrW(958))
    text = Replace(text, "Chi", VBA.ChrW(935))
    text = Replace(text, "chi", VBA.ChrW(967))
    text = Replace(text, "tau", VBA.ChrW(964))
    text = Replace(text, "Pi", VBA.ChrW(928))
    text = Replace(text, "greek-nu", VBA.ChrW(957))
    text = Replace(text, "kappa", VBA.ChrW(954))
    text = Replace(text, "zeta", VBA.ChrW(950))
    text = Replace(text, "eta", VBA.ChrW(951))    ' skal være sidst da eta indgår i andre
    text = Replace(text, "increment", VBA.ChrW(8710))  ' specielt delta increment
    text = Replace(text, "Symhalf", VBA.ChrW(189))    ' ½
    text = Replace(text, "degC", VBA.ChrW(8451))   ' specielt oC tegn
    text = Replace(text, "<=", VBA.ChrW(8804))  '
    text = Replace(text, ">=", VBA.ChrW(8805))    '
    text = Replace(text, "CVinkelO", VBA.ChrW(8736))    '
    text = Replace(text, "CVinkel", VBA.ChrW(8736))    '
    text = Replace(text, "Symangle", VBA.ChrW(8736))  '

    ' konverter vektor og matrixpile
    '    text = Replace(text, "SymVecta", VBA.ChrW(8407))   ' vektorpil
    '    text = Replace(text, "SymVectb", VBA.ChrW(8401)) ' vektorpil men ikke hel pil
    '    text = Replace(text, "SymVectc", VBA.ChrW(773)) ' matrixtegn
    ''    text = Replace(text, "SymVectc", VBA.ChrW(175))  ' matrixtegn \overbar

    ea.text = text
    ' SymVecta til vektorpil
    pos = 1
    Do
        pos = InStr(pos, ea.text, "SymVecta")
        If pos > 0 Then
            ea.pos = pos
            s = ea.GetPrevVar(pos)
            s = Left(s, Len(s) - (ea.pos - pos) - 1)
            If pos + 7 = Len(ea.text) Then
                ea.text = Left(ea.text, pos - 1 - Len(s)) & "(" & s & ")" & VBA.ChrW(8407)
            Else
                ea.text = Left(ea.text, pos - 1 - Len(s)) & "(" & s & ")" & VBA.ChrW(8407) & right(ea.text, Len(ea.text) - pos - 7)
            End If
            pos = pos + 1
        End If
    Loop While pos > 0

    ' SymVectb til halv vektorpil
    pos = 1
    Do
        pos = InStr(pos, ea.text, "SymVectb")
        If pos > 0 Then
            ea.pos = pos
            s = ea.GetPrevVar(pos)
            s = Left(s, Len(s) - (ea.pos - pos) - 1)
            If pos + 7 = Len(ea.text) Then
                ea.text = Left(ea.text, pos - 1 - Len(s)) & "(" & s & ")" & VBA.ChrW(8401)
            Else
                ea.text = Left(ea.text, pos - 1 - Len(s)) & "(" & s & ")" & VBA.ChrW(8401) & right(ea.text, Len(ea.text) - pos - 7)
            End If
            pos = pos + 1
        End If
    Loop While pos > 0

    ' SymVectc til halv vektorpil
    pos = 1
    Do
        pos = InStr(pos, ea.text, "SymVectc")
        If pos > 0 Then
            ea.pos = pos
            s = ea.GetPrevVar(pos)
            s = Left(s, Len(s) - (ea.pos - pos) - 1)
            If pos + 7 = Len(ea.text) Then
                ea.text = Left(ea.text, pos - 1 - Len(s)) & "(" & s & ")" & VBA.ChrW(773)
            Else
                ea.text = Left(ea.text, pos - 1 - Len(s)) & "(" & s & ")" & VBA.ChrW(773) & right(ea.text, Len(ea.text) - pos - 7)
            End If
            pos = pos + 1
        End If
    Loop While pos > 0

    ' lambert w
    Do
        pos = InStr(ea.text, "generalized_lambert_w")
        If pos > 0 Then
        s = ea.GetNextBracketContent(pos + 20)
        arr = Split(s, ListSeparator)
        If UBound(arr) = 0 Then
            arr = Split(s, ",")
        End If
        s2 = Trim(arr(0))
        If (s2 = "0" Or s2 = "-1") And UBound(arr) = 1 Then
            ea.text = Left(ea.text, pos - 1) & "W_(" & s2 & ") (" & arr(1) & ")" & right(ea.text, Len(ea.text) - ea.pos + 1)
        End If
        End If
    Loop While pos > 0


    ' Disse kan være resultat fra grænseværdi
    ea.ReplaceVar "inf", VBA.ChrW(8734)
    ea.ReplaceVar "minf", "-" & VBA.ChrW(8734)
    ea.ReplaceVar "infinity", VBA.ChrW(8734)
    ea.ReplaceVar "und", Sprog.A(431) '"ikke defineret"
    ea.ReplaceVar "ind", Sprog.A(432) ' "ikke defineret men begrænset"
    ea.ReplaceVar "tgrader", VBA.ChrW(176)     ' gradtegn
    text = ea.text

    '
    text = Replace(text, "SPC", " ")


    'ingen komplekse udtryk
    If Not MaximaComplex Then
        If InStr(text, "%i") > 0 Then
            text = Sprog.A(433) & " " & VBA.ChrW(8477)
            GoTo Slut
        End If
    Else
        text = Replace(text, "%i", "i")
    End If

    '    text = Replace(text, "", "\")
    GoTo Slut
Fejl:
    MsgBox Sprog.A(100) & vbCrLf & vbCrLf & text, vbOKOnly, Sprog.Error
Slut:
    ConvertToWordSymbols = text
End Function
Function ConvertRadToDegree(text As String, trigfunc As String) As String
    Dim pos, spos As Integer
    Dim ea As New ExpressionAnalyser
    ea.StartBracket = "("
    ea.EndBracket = ")"
    ea.text = text
    spos = 1

    Do
        pos = ea.FindChr("a" & trigfunc, spos)
        If pos > 0 Then
            ea.GetNextBracketContent
            ea.InsertBeforePos (")")
            ea.pos = pos
            ea.InsertBeforePos (VBA.ChrW(960) & "/180" & MaximaGangeTegn & "(")
            spos = pos + 11
        End If
    Loop While pos > 0

    spos = 1
    Do
        pos = ea.FindChr(trigfunc, spos)
        If pos > 0 And Not (ea.ChrByIndex(pos - 1) = "a") Then
            ea.GetNextBracketContent
            ea.InsertBeforePos (")")
            ea.pos = pos + Len(trigfunc)
            ea.InsertAfterPos ("180/" & VBA.ChrW(960) & MaximaGangeTegn & "(")
            spos = pos + 11
        Else
            spos = spos + 3
        End If
    Loop While pos > 0

    ConvertRadToDegree = ea.text

End Function
Sub RemoveComplexSolutions()
'kun reelle løsninger ' problem at selvom der er i i kan det måske reduceres til reelt
    Dim arr As Variant
    Dim i As Integer
    If Not MaximaComplex Then
        '        If InStr(MaximaOutput, "%i") Then
        arr = Split(MaximaOutput, ListSeparator)
        MaximaOutput = ""
        For i = 0 To UBound(arr)
            If InStr(arr(i), "%i") < 1 Then
                If MaximaOutput <> "" Then MaximaOutput = MaximaOutput & "    " & VBA.ChrW(8744) & "    "
                MaximaOutput = MaximaOutput & arr(i)
            End If
        Next
        '        End If
    Else
        MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8744) & "    ")    ' komma erstattes med eller
    End If

End Sub
Function ConvertToLatex(text As String) As String
    Dim pos As Integer, pos1 As Integer, pos2 As Integer, pos3 As Integer, pos4 As Integer
    Dim ea As New ExpressionAnalyser
    Dim ea2 As New ExpressionAnalyser
    Dim s As String, t As String
    Dim arr As Variant
    Dim ib As Boolean
    Dim rod As Integer
    Dim ends As Integer
    Dim subs As Integer, sups As Integer
    
    If ConvertTexWithMaxima Then
        ConvertToLatex = ConvertToLatexMaxima(text)
        Exit Function
    End If
    ea.SetNormalBrackets
    ea2.SetTuborgBrackets


    text = Replace(text, "-" & VBA.ChrW(8734), "\minfty ")
    text = Replace(text, VBA.ChrW(8734), "\infty ")

    'diverse symboler
'    text = Replace(text, VBA.ChrW(183), "* ")    ' prik erstattes med gange
    text = Replace(text, VBA.ChrW(8901), VBA.ChrW(183))     ' \cdot , ved ikke hvorfor der tilsyneladende er to
    text = Replace(text, VBA.ChrW(8729), VBA.ChrW(183))    ' prik \cdot erstattes med gange (denne er \bullet)
    text = Replace(text, VBA.ChrW(8226), VBA.ChrW(183))    ' tyk prik erstattes med gange
    
    text = Replace(text, VBA.ChrW(247), "\div ")    ' divisionstegn med prikker over/under


    text = Replace(text, vbCrLf, "")
    text = Replace(text, vbCr, "")
    text = Replace(text, vbLf, "")

    text = Replace(text, VBA.ChrW(9633), "")    ' tegn foran paranteser der gør dem usynlige. Muligvis problematisk da paranteserne ikke fjernes. gør også brøker små
    text = Replace(text, VBA.ChrW(9508), "")    ' tegn for tom plads ved ihvertfald [ ]
    text = Replace(text, VBA.ChrW(11), "")    ' shift-enter
    text = Replace(text, """", "")    ' apostrof fjernes
    text = Replace(text, VBA.ChrW(180) & VBA.ChrW(180), "^''")    ' der er tre forskellige '´` apostrof lignende de forståes alle som diff
    text = Replace(text, VBA.ChrW(96) & VBA.ChrW(96), "^''")    ' der er tre forskellige '´`
    text = Replace(text, VBA.ChrW(180), "^'")    ' der er tre forskellige '´` apostrof lignende de forståes alle som diff
    text = Replace(text, VBA.ChrW(96), "^'")    ' der er tre forskellige '´`
    '    text = Replace(text, VBA.ChrW(176), "\circ ") ' gradtegn
    text = Replace(text, VBA.ChrW(189), "\frac{1}{2} ")    ' ½
    text = Replace(text, VBA.ChrW(8800), "\neq ")    ' ikkeligmed
    text = Replace(text, VBA.ChrW(177), "\pm ")    ' plus/minus
    text = Replace(text, VBA.ChrW(8723), "\mp ")    ' minus-plus
    text = Replace(text, VBA.ChrW(8804), "\leq ")    ' <=
    text = Replace(text, VBA.ChrW(8805), "\geq ")    ' >=
    text = Replace(text, VBA.ChrW(8776), "\approx ")    ' cirka lig
    text = Replace(text, VBA.ChrW(8773), "\cong ")    ' cirka lig
    text = Replace(text, VBA.ChrW(8656), "\Leftarrow ")    '
    text = Replace(text, VBA.ChrW(8658), "\Rightarrow ")    '
    text = Replace(text, VBA.ChrW(8660), "\Leftrightarrow ")    ' ensbetydende tegn
    text = Replace(text, VBA.ChrW(8661), "\Updownarrow ")    ' ensbetydende tegn op/ned
    text = Replace(text, VBA.ChrW(8657), "\Uparrow ")    ' dobbeltpil op
    text = Replace(text, VBA.ChrW(8659), "\Downarrow ")    ' dobbeltpil ned
    text = Replace(text, VBA.ChrW(8593), "\uparrow ")    '
    text = Replace(text, VBA.ChrW(8595), "\downarrow ")    '
    text = Replace(text, VBA.ChrW(8596), "\leftrightarrow ")    '
    text = Replace(text, VBA.ChrW(10187), "\partial ")    '
    text = Replace(text, VBA.ChrW(8469), "\mathbb{N} ")    ' symbol for naturlige tal
    text = Replace(text, VBA.ChrW(8484), "\mathbb{Z} ")    ' symbol for hele tal
    text = Replace(text, VBA.ChrW(8474), "\mathbb{Q} ")    ' symbol for rationelle tal
    text = Replace(text, VBA.ChrW(8477), "\mathbb{R} ")    ' symbol for reelle tal
    text = Replace(text, VBA.ChrW(8450), "\mathbb{C} ")    ' symbol for komplekse tal
    text = Replace(text, VBA.ChrW(8463), "\hbar ")    ' planck konstant/2pi
    text = Replace(text, VBA.ChrW(8704), "\forall ")    ' for alle tegn
    text = Replace(text, VBA.ChrW(8712), "\in ")    ' tilhører tegn
    text = Replace(text, VBA.ChrW(8733), "\propto ")    ' propotionel tegn
    text = Replace(text, VBA.ChrW(8810), "\ll ")    ' <<
    text = Replace(text, VBA.ChrW(8811), "\gg ")    ' >>
    text = Replace(text, VBA.ChrW(8704), "\exists ")    ' eksisterer tegn
    text = Replace(text, VBA.ChrW(8708), "\exists ")    ' eksisterer tegn
    text = Replace(text, VBA.ChrW(8713), "\notin ")    ' tilhører ikke
    text = Replace(text, VBA.ChrW(8709), "\emptyset ")    ' Ø tomme mængde
    text = Replace(text, VBA.ChrW(8711), "\nabla ")    ' nabla operator
    text = Replace(text, VBA.ChrW(8801), "\equiv ")    ' ligmed 3 streger
    text = Replace(text, VBA.ChrW(8710), "\inc ")    ' specielt delta increment
    text = Replace(text, VBA.ChrW(8943), "\ldots ")    ' ...
    text = Replace(text, VBA.ChrW(8942), "\vdots ")    ' 3 prikker lodret
    text = Replace(text, VBA.ChrW(8944), "\rddots ")    ' 3 prikker skråt op til højre
    text = Replace(text, VBA.ChrW(8945), "\ddots ")    ' 3 prikker skråt ned til højre
    text = Replace(text, VBA.ChrW(8756), "\therefore ")    ' 3 prikker i trekant
    text = Replace(text, VBA.ChrW(8757), "\because ")    ' 3 prikker i trekant
    text = Replace(text, VBA.ChrW(8834), "\subset ")    ' tilhører mængde (ægte delmængde)
    text = Replace(text, VBA.ChrW(8838), "\subseteq ")    ' tilhører lig mængde
    text = Replace(text, VBA.ChrW(8745), "\cap ")    ' foreningsmængde
    text = Replace(text, VBA.ChrW(8746), "\cup ")    ' fællesmængde
    text = Replace(text, VBA.ChrW(8764), "\sim ")    'ca ligmed en bølge
    text = Replace(text, VBA.ChrW(8771), "\simeq ")    ' ca ligmed en bølge en lige
    text = Replace(text, VBA.ChrW(8855), "\otimes ")    ' mængdeprodukt
    text = Replace(text, VBA.ChrW(8614), "\mapsto ")    ' pil til funktionsdef
    text = Replace(text, VBA.ChrW(8640), "\rightharpoonup ")    ' pil højre med kun halv pil
    text = Replace(text, VBA.ChrW(8651), "\lrhar ")    ' harpun højre/venstre
    text = Replace(text, VBA.ChrW(8598), "\nwarrow ")    ' pil op til venstre
    text = Replace(text, VBA.ChrW(8599), "\nearrow ")    ' pil op til højre
    text = Replace(text, VBA.ChrW(8600), "\nwarrow ")    ' pil ned højre
    text = Replace(text, VBA.ChrW(8601), "\swarrow ")    ' pil ned venstre
    text = Replace(text, VBA.ChrW(8728), "\circ ")    ' grad-tegn
    text = Replace(text, VBA.ChrW(8451), "^\circ\text{C} ")    ' specielt oC tegn
    text = Replace(text, VBA.ChrW(8457), "^\circ\text{F} ")    ' specielt oF tegn
    text = Replace(text, VBA.ChrW(176), "^\circ ")    '
    text = Replace(text, VBA.ChrW(8736), "\angle ")    '
    text = Replace(text, VBA.ChrW(160), "")    ' foran accent f.eks. vector pil 160 8407
    text = Replace(text, VBA.ChrW(8743), "\wedge ")    ' og
    text = Replace(text, VBA.ChrW(8744), "\vee ")    ' eller

'    text = Replace(text, VBA.ChrW(8721), "\sum ")    ' sumtegn. Skal måske nok laves mere avanceret
'    text = Replace(text, VBA.ChrW(8719), "\prod ")    ' produkttegn. Skal måske nok laves mere avanceret
    text = Replace(text, VBA.ChrW(8748), "\iint ")    ' dobbeltintegraltegn. Skal måske nok laves mere avanceret
    text = Replace(text, VBA.ChrW(8749), "\iiint ")    ' tripelintegraltegn. Skal måske nok laves mere avanceret
    text = Replace(text, VBA.ChrW(8750), "\oint ")    ' kurveintegraltegn. Skal måske nok laves mere avanceret


    '   Fjern problematiske usynlige paranteser efter og foran / måske ikke nødvendig for latex? ihvertfald problematisk for x/10^2 -> x/(10)^2
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
    '        If ea.ValidateBrackets And InStr(ea.text, "=") < 1 Then ' forsøger at løse problem hvor skjulte og alm paranteser er blandet som f.eks.[( ])
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
    ea.text = text

    ea.text = Replace(ea.text, VBA.ChrW(12310), "")    ' specielle usynlige paranteser fjernes   flyttet fra før trig med store bogstaver for at få insertfunctionbrackets til at virke
    ea.text = Replace(ea.text, VBA.ChrW(12311), "")    ' specielle usynlige paranteser fjernes

'lim
    pos = -1
    Do
        pos = InStr(pos + 2, ea.text, "lim")
        If pos > 0 And ea.ChrByIndex(pos - 1) <> "\" Then
            ea.pos = pos
            If ea.ChrByIndex(pos + 3) = VBA.ChrW(9516) Then
                ea.ReplaceStringAt "_", pos + 3
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
            ea.text = Left(ea.text, pos - 1) & "\lim " & right(ea.text, Len(ea.text) - pos - 2)
        End If
    Loop While pos > 0

'    ea.text = Replace(ea.text, "cosh^(-1)", "cosh^{-1}")
'    ea.text = Replace(ea.text, "sinh^(-1)", "sinh^{-1}")
'    ea.text = Replace(ea.text, "tanh^(-1)", "tanh^{-1}")
'    ea.text = Replace(ea.text, "cos^(-1) ", "cos^{-1}")    ' under sjældne omstændigheder kan der være et mellemrum istedet for et 8289 tegn efter
'    ea.text = Replace(ea.text, "sin^(-1) ", "sin^{-1}")    ' det blive forstået som et gangetegn hvis ikke fjernes
'    ea.text = Replace(ea.text, "tan^(-1) ", "tan^{-1}")
'    ea.text = Replace(ea.text, "cos^(-1)", "cos^{-1}")
'    ea.text = Replace(ea.text, "sin^(-1)", "sin^{-1}")
'    ea.text = Replace(ea.text, "tan^(-1)", "tan^{-1}")

    ea.text = ConvertTrigPowLatex("sin", ea.text) ' skal være før funktionstegn fjernes
    ea.text = ConvertTrigPowLatex("cos", ea.text) ' skal være før funktionstegn fjernes
    ea.text = ConvertTrigPowLatex("tan", ea.text) ' skal være før funktionstegn fjernes

    ea.text = Replace(ea.text, VBA.ChrW(8289), "") ' funktionstegn

    'fbox
    ea.text = ConvertPreAccent(ea.text, VBA.ChrW(9645), "\fbox{", "}")

    'overbrace
    ea.text = ConvertPreAccent(ea.text, VBA.ChrW(9182), "\overbrace{", "}")

    'underbrace
    ea.text = ConvertPreAccent(ea.text, VBA.ChrW(9183), "\underbrace{", "}")


    'vektor-notation med pil over bogstav
    ea.text = ConvertAccent(ea.text, VBA.ChrW(8407), "\overrightarrow{", "}")

    'tværvektor \hat
    ea.text = ConvertAccent(ea.text, VBA.ChrW(770), "\widehat{", "}")
    
    '\overline
    ea.text = ConvertAccent(ea.text, VBA.ChrW(773), "\overline{", "}")
    'dobbelt overline
    ea.text = ConvertAccent(ea.text, VBA.ChrW(831), "\overline{\overline{", "}}")
    
    'halvpil
    ea.text = ConvertAccent(ea.text, VBA.ChrW(8401), "\overset{\rightharpoonup}{", "}")

    'dobbeltpil over
    ea.text = ConvertAccent(ea.text, VBA.ChrW(8417), "\overleftrightarrow{", "}")
    
    'venstrepil over
    ea.text = ConvertAccent(ea.text, VBA.ChrW(8406), "\overleftarrow{", "}")

    '\widetilde
    ea.text = ConvertAccent(ea.text, VBA.ChrW(771), "\widetilde{", "}")

    '\dot
    ea.text = ConvertAccent(ea.text, VBA.ChrW(775), "\dot{", "}")
    '\ddot
    ea.text = ConvertAccent(ea.text, VBA.ChrW(776), "\ddot{", "}")
    '\ddot
    ea.text = ConvertAccent(ea.text, VBA.ChrW(8411), "\dddot{", "}")

    '\underset
    Do
        pos = InStr(ea.text, VBA.ChrW(9516))
        If pos > 0 Then
            If ea.ChrByIndex(pos + 1) = "(" Then
                ea.SetNormalBrackets
                s = ea.GetNextBracketContent(pos + 1)
                pos2 = ea.pos - 1
            Else
                s = ea.GetNextVar(pos)
                pos2 = ea.pos
            End If
            If ea.ChrByIndex(pos - 1) = ")" Then
                ea.SetNormalBrackets
                t = ea.GetPrevBracketContent(pos)
                pos1 = ea.pos - 2
            ElseIf ea.ChrByIndex(pos - 1) = "}" Then
                ea.SetTuborgBrackets
                t = ea.GetPrevBracketContent(pos)
                pos1 = ea.pos - 2
                Do While ea.IsLatexLetterPos(pos1)
                    pos1 = pos1 - 1
                Loop
                If ea.ChrByIndex(pos1) = "\" Then
                    t = Mid(ea.text, pos1, pos - pos1)
                    pos1 = pos1 - 1
                End If
            Else
                t = ea.GetPrevVar(pos)
                pos1 = ea.pos - Len(t)
            End If
            ea.text = Left(ea.text, pos1) & "\underset{" & t & "}{" & s & "}" & right(ea.text, Len(ea.text) - pos2)
        End If
    Loop While pos > 0
    
    

    ' sumtegn
    Do
        pos = InStr(ea.text, VBA.ChrW(8721))
        If pos > 0 Then
            ea.pos = pos
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
            ea.text = Left(ea.text, pos - 1) & "\sum " & right(ea.text, Len(ea.text) - pos)
        End If
    Loop While pos > 0

    ' produkttegn
    Do
        pos = InStr(ea.text, VBA.ChrW(8719))
        If pos > 0 Then
            ea.pos = pos
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
            ea.text = Left(ea.text, pos - 1) & "\prod " & right(ea.text, Len(ea.text) - pos)
        End If
    Loop While pos > 0

    ' integrale
    pos = InStr(ea.text, VBA.ChrW(8747))
    Do While pos > 0
        pos1 = InStr(pos, ea.text, "_")
        pos2 = InStr(pos, ea.text, "^")
        pos3 = InStr(pos, ea.text, VBA.ChrW(9618))
        If pos3 = 0 Then pos3 = pos
        If pos1 > 0 Then
            Dim ul As String, ll As String
            ll = Mid(ea.text, pos1 + 1, pos2 - pos1 - 1)
            ul = Mid(ea.text, pos2 + 1, pos3 - pos2 - 1)
            If Left(ll, 1) = "(" And right(ll, 1) = ")" Then ll = Mid(ll, 2, Len(ll) - 2) ' Word alm indsætter parenteser som ikke skal ses. Hvis der skal være parenteser så er der dobbelt
            If Left(ul, 1) = "(" And right(ul, 1) = ")" Then ul = Mid(ul, 2, Len(ul) - 2)
            ea.text = Left(ea.text, pos - 1) & "\int_{" & ll & "}^{" & ul & "} " & right(ea.text, Len(ea.text) - pos3)
        Else    ' ubestemt
            ea.text = Left(ea.text, pos - 1) & "\int " & right(ea.text, Len(ea.text) - pos3)
        End If
        pos = InStr(ea.text, VBA.ChrW(8747))
    Loop

    ' ^
    pos = InStr(ea.text, "^")
    Do While pos > 0
        If ea.ChrByIndex(pos + 1) = "(" Then
            s = Trim(ea.GetNextBracketContent(pos))
            If Not (Left(s, 1) = "{" And right(s, 1) = "}") Then s = "{" & s & "}"
            ea.text = Left(ea.text, pos) & s & right(ea.text, Len(ea.text) - ea.pos + 1)
        ElseIf ea.ChrByIndex(pos + 1) = "{" Then ' hvis allerede er konverteret
        Else
            ea.pos = pos + 1
            ea.FindHierakiBreak 4, False
            s = ea.SubString(pos + 1, ea.pos)
            If Not (Left(s, 1) = "{" And right(s, 1) = "}") Then s = "{" & s & "}"
            ea.text = Left(ea.text, pos) & s & right(ea.text, NotZero(ea.Length - ea.pos))
        End If
        pos = InStr(pos + 1, ea.text, "^")
    Loop



    '_
    pos = InStr(ea.text, "_")
    Do While pos > 0
        If ea.ChrByIndex(pos + 1) = "(" Then
            s = ea.GetNextBracketContent
            If Not (Left(s, 1) = "{" And right(s, 1) = "}") Then s = "{" & s & "}"
            ea.text = Left(ea.text, pos) & s & right(ea.text, Len(ea.text) - ea.pos + 1)
        Else
            ea.pos = pos + 1
            ea.FindHierakiBreak 4, False
            s = ea.SubString(pos + 1, ea.pos)
            If Not (Left(s, 1) = "{" And right(s, 1) = "}") Then s = "{" & s & "}"
            ea.text = Left(ea.text, pos) & s & right(ea.text, NotZero(ea.Length - ea.pos))
        End If
        pos = InStr(pos + 1, ea.text, "_")
    Loop

    ' /
    pos = InStr(ea.text, "/")
    Dim n As String
    Do While pos > 0
        If ea.ChrByIndex(pos - 1) = ")" Then
            ea.SetNormalBrackets
            t = ea.GetPrevBracketContent(pos)
            If ea.IsLetterPos(ea.pos - 2) Or ea.IsNumber(ea.pos - 2) Or ea.IsSquareRoot(ea.pos - 2) Or ea.ChrByIndex(ea.pos - 2) = "^" Then
                ea.pos = ea.pos - 2
                '                pos1 = ea.FindPrevBreak
                pos1 = ea.FindPrevHierakiBreak(4, False)
                t = Mid(ea.text, pos1, pos - pos1)
            Else
                pos1 = ea.pos - 1
            End If
        ElseIf ea.ChrByIndex(pos - 1) = "}" Then
            ea.SetTuborgBrackets
            t = ea.GetPrevBracketContent(pos)
            If ea.IsLetterPos(ea.pos - 2) Or ea.IsNumber(ea.pos - 2) Or ea.IsSquareRoot(ea.pos - 2) Or ea.ChrByIndex(ea.pos - 2) = "^" Then
                ea.pos = ea.pos - 2
                '                pos1 = ea.FindPrevBreak
                pos1 = ea.FindPrevHierakiBreak(3, False)
                t = Mid(ea.text, pos1, pos - pos1)
            Else
                pos1 = ea.pos - 1
            End If
        Else
            ea.pos = pos - 1
            '            pos1 = ea.FindPrevBreak
            pos1 = ea.FindPrevHierakiBreak(4, False)
            t = Mid(ea.text, pos1, pos - pos1)
        End If

        '        If ea.ChrByIndex(pos + 1) = "(" Then
        '            n = ea.GetNextBracketContent(pos)
        '            pos2 = ea.pos
        '        Else
        ea.pos = pos + 1
        ''            pos2 = ea.FindNextBreak
        pos2 = ea.FindHierakiBreak(3, False)
        
        n = Mid(ea.text, pos + 1, pos2 - pos)
        If Left(n, 1) = "(" And right(n, 1) = ")" Then
            n = Mid(n, 2, Len(n) - 2)
        End If
        pos2 = pos2 + 1
        '        End If
        ea.text = Left(ea.text, pos1 - 1) & "\frac{" & t & "}{" & n & "}" & right(ea.text, NotZero(Len(ea.text) - pos2 + 1))
        pos = InStr(ea.text, "/")
    Loop


    ' kubikrod
    '    ea.text = Replace(ea.text, VBA.ChrW(8731) & "(", VBA.ChrW(8730) & "(3&")
    '    ea.text = Replace(ea.text, VBA.ChrW(8731), VBA.ChrW(8730) & "(3&")

    ' 3 og 4 rod
    For rod = 3 To 4
        Do
            pos = InStr(ea.text, VBA.ChrW(8728 + rod))
            If pos > 0 Then
                ea.pos = pos + 1
                If Mid(ea.text, pos + 1, 1) <> "(" Then
                    ea.SetTuborgBrackets
                    ea.InsertUnderstoodBracketPair
                    ea.text = Replace(ea.text, VBA.ChrW(8728 + rod), "\sqrt[" & rod & "]", 1, 1)
                Else
                    ea.SetNormalBrackets
                    s = ea.GetNextBracketContent()
                    ea.text = Left(ea.text, pos - 1) & "\sqrt[" & rod & "]{" & s & "}" & right(ea.text, NotZero(ea.Length - ea.pos + 1))
                End If
                ea.pos = pos
            End If
        Loop While pos > 0
    Next

    'kvadratrod
    Do
        pos = InStr(ea.text, VBA.ChrW(8730))
        If pos > 0 Then
            If Mid(ea.text, pos + 1, 1) <> "(" Then
                ea.text = Replace(ea.text, VBA.ChrW(8730), "\sqrt", 1, 1)
                pos = pos + 5
                ea.pos = pos
                ea.SetTuborgBrackets
                ea.InsertUnderstoodBracketPair
            Else
                ea.pos = pos
                arr = Split(ea.GetNextBracketContent, "&")
                pos2 = ea.pos
                If UBound(arr) = 0 Then
'                    ea.text = Replace(ea.text, VBA.ChrW(8730), "\sqrt ", 1, 1)
                    ea.text = Left(ea.text, pos - 1) & "\sqrt{" & arr(0) & "}" & right(ea.text, NotZero(ea.Length - ea.pos + 1))
                ElseIf UBound(arr) = 1 Then
                    ea.text = Left(ea.text, pos - 1) & "\sqrt[" & arr(0) & "]{" & arr(1) & "}" & right(ea.text, NotZero(ea.Length - ea.pos + 1))
                End If
            End If
        End If
    Loop While pos > 0

    ' matricer
    ea.text = Replace(ea.text, VBA.ChrW(9608), VBA.ChrW(9632))
    pos = InStr(ea.text, VBA.ChrW(9632))
    Do While pos > 0
        ea.SetNormalBrackets
        s = ea.GetNextBracketContent(pos)
        ea.text = Left(ea.text, pos - 1) & "\begin{matrix}" & Replace(s, "@", "\\") & "\end{matrix}" & right(ea.text, NotZero(ea.Length - pos - Len(s) - 2))
        pos = InStr(ea.text, VBA.ChrW(9632))
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
    ea.pos = 1
    s = ea.GetNextVar()
    Do
        ea.pos = ea.pos + 1
        If ea.IsUnit(s) Then
            If ea.IsNumber(ea.pos - Len(s) - 1) Or ea.ChrByIndex(ea.pos - Len(s) - 1) = " " Then
                t = "\,\mathrm{" & s & "}"
            Else
                t = "\mathrm{" & s & "}"
            End If
            ea.ReplaceVar s, t, ea.pos - Len(s), 1
            ea.pos = ea.pos + Len(t) - Len(s)
        End If
        ea.text = Replace(ea.text, VBA.ChrW(8289), "")    ' funktionstegn
        s = ea.GetNextVar()
    Loop While s <> ""
    End If

    'græske bogstaver
    ea.text = Replace(ea.text, VBA.ChrW(916), "\Delta ")
    ea.text = Replace(ea.text, VBA.ChrW(948), "\delta ")
    ea.text = Replace(ea.text, VBA.ChrW(945), "\alpha ")
    ea.text = Replace(ea.text, VBA.ChrW(946), "\beta ")
    ea.text = Replace(ea.text, VBA.ChrW(947), "\gammaLB ")
    ea.text = Replace(ea.text, VBA.ChrW(952), "\theta ")
    ea.text = Replace(ea.text, VBA.ChrW(920), "\Theta ")
    ea.text = Replace(ea.text, VBA.ChrW(955), "\lambda ")
    ea.text = Replace(ea.text, VBA.ChrW(923), "\Lambda ")
    ea.text = Replace(ea.text, VBA.ChrW(956), "\mu ")
    ea.text = Replace(ea.text, VBA.ChrW(961), "\rho ")
    ea.text = Replace(ea.text, VBA.ChrW(963), "\sigma ")
    ea.text = Replace(ea.text, VBA.ChrW(931), "\Sigma ")
    ea.text = Replace(ea.text, VBA.ChrW(981), "\phi ")
    ea.text = Replace(ea.text, VBA.ChrW(934), "\Phi ")
    ea.text = Replace(ea.text, VBA.ChrW(949), "\varepsilon ")
    ea.text = Replace(ea.text, VBA.ChrW(1013), "\epsilon ")
    ea.text = Replace(ea.text, VBA.ChrW(968), "\psi ")
    ea.text = Replace(ea.text, VBA.ChrW(936), "\Psi ")
    ea.text = Replace(ea.text, VBA.ChrW(926), "\Xi ")
    ea.text = Replace(ea.text, VBA.ChrW(958), "\xi ")
    ea.text = Replace(ea.text, VBA.ChrW(935), "\Chi ")
    ea.text = Replace(ea.text, VBA.ChrW(967), "\chi ")
    ea.text = Replace(ea.text, VBA.ChrW(928), "\Pi ")
    ea.text = Replace(ea.text, VBA.ChrW(964), "\tau ")
    ea.text = Replace(ea.text, VBA.ChrW(957), "\greek-nu ")
    ea.text = Replace(ea.text, VBA.ChrW(954), "\kappa ")
    ea.text = Replace(ea.text, VBA.ChrW(951), "\eta ")
    ea.text = Replace(ea.text, VBA.ChrW(950), "\zeta ")
    ea.text = Replace(ea.text, VBA.ChrW(960), "\pi ")

    ea.text = Replace(ea.text, VBA.ChrW(8592), "\leftarrow ")    '
    ea.text = Replace(ea.text, VBA.ChrW(8594), "\rightarrow ")    '
    ea.text = Replace(ea.text, VBA.ChrW(8652), "\rightleftharpoons ")    ' egentlig \leftrightharpoons   but not supported by default Latex
    ea.text = Replace(ea.text, VBA.ChrW(8651), "\rightleftharpoons ")    '
    ea.text = Replace(ea.text, VBA.ChrW(8592), "\leftarrow ")    '
    ea.text = Replace(ea.text, VBA.ChrW(10236), "\longmapsto ")    '
    ea.text = Replace(ea.text, VBA.ChrW(8640), "\rightharpoonup ")    '
    ea.text = Replace(ea.text, VBA.ChrW(8641), "\rightharpoondown ")    '
    ea.text = Replace(ea.text, VBA.ChrW(183), "\cdot ")    ' prik erstattes med gange
    ea.text = Replace(ea.text, VBA.ChrW(215), "\times ")    ' x gangetegn
    ea.text = Replace(ea.text, " \", "\") ' i Moodle bliver mellemrum lavet. Derfor skal de helst fjernes. mellemrum før \ er ihvertfald ikke nødvendigt

    ' parenteser
    ea.text = Replace(ea.text, "(", "\left(")    '
    ea.text = Replace(ea.text, ")", "\right)")    '
    
    'kantede parenteser ' ikke så simpelt fordi de indgår i latex syntaks
    ea.SetSquareBrackets
    pos1 = 1
    Do
        pos = InStr(pos1, ea.text, "[")
        If pos > 0 And ea.ChrByIndex(pos - 1) <> "}" Then
            pos1 = pos + 6
            pos2 = pos
            ea.GetNextBracketContent pos - 1
            pos3 = ea.pos
            If Not ea.ChrByIndex(pos3) = "{" Then
                ea.text = Left(ea.text, pos2 - 1) & "\left" & right(ea.text, Len(ea.text) - pos2 + 1)
                ea.text = Left(ea.text, pos3 + 3) & "\right" & right(ea.text, Len(ea.text) - pos3 - 3)
            End If
        Else
            pos1 = pos + 1
        End If
    Loop While pos > 0

    ConvertToLatex = ea.text
End Function
Function ConvertTrigPowLatex(trig As String, text As String) As String
' omdanner f.eks. sin^2(x) -> sin^2(x)
Dim ea As ExpressionAnalyser
Dim pow As String
Dim pos2 As Integer, pos As Integer
pos = 1
Do
    pos = InStr(pos, text, trig & "^")
    If pos > 0 Then
        pos = pos + Len(trig) + 1
        pos2 = InStr(pos, text, VBA.ChrW(8289))
        If pos2 = 0 Then
            Set ea = New ExpressionAnalyser
            ea.SetNormalBrackets
            ea.text = text
            ea.pos = pos
            pos2 = ea.FindEndOfBracket
        End If
        pow = Mid(text, pos, pos2 - pos)
        If pow = "(-1)" Then pow = "-1"
        text = Left(text, pos - 1) & "{" & pow & "}" & right(text, Len(text) - pos2)
'        InsertTextAt "^" & pow, FindEndOfBracket
        
    End If
Loop While pos > 0
ConvertTrigPowLatex = text
End Function
Function ConvertAccent(text As String, Accent As String, aString, bString) As String
' ConvertAccent(ea.text,VBA.ChrW(8407), "\overrightarrow{","}")
Dim ea As New ExpressionAnalyser, pos As Integer, pos1 As Integer, s As String
    ea.text = text
    pos = InStr(ea.text, Accent)
    Do While pos > 0
        If ea.ChrByIndex(pos + 1) = " " Then ea.RemoveChar (pos + 1) ' der kan være mellemrum efter pilen som kan forstyrre
        If Mid(ea.text, pos - 1, 1) = ")" Then
            ea.SetNormalBrackets
            s = ea.GetPrevBracketContent(pos)
            ea.text = Left(ea.text, pos - Len(s) - 3) & aString & s & bString & right(ea.text, Len(ea.text) - pos)
        ElseIf Mid(ea.text, pos - 1, 1) = "}" Then
            ea.SetTuborgBrackets
            s = ea.GetPrevBracketContent(pos)
            pos1 = ea.pos - 2
            Do While ea.IsLetterPos(pos1, False)
                pos1 = pos1 - 1
            Loop
            If ea.ChrByIndex(pos1) = "\" Then
                ea.text = Left(ea.text, pos1 - 1) & aString & Mid(ea.text, pos1, pos - pos1) & bString & right(ea.text, Len(ea.text) - pos - Len(Accent) + 1)
            End If
        Else
            ea.text = Left(ea.text, pos - 2) & aString & Mid(ea.text, pos - 1, 1) & bString & right(ea.text, Len(ea.text) - pos - Len(Accent) + 1)
        End If
        pos = InStr(pos + 1, ea.text, Accent)
    Loop
    ConvertAccent = ea.text
End Function
Function ConvertPreAccent(text As String, Accent As String, aString, bString) As String
    Dim ea As New ExpressionAnalyser
    Dim pos As Integer, s As String
    ea.text = text
    Do
        pos = InStr(ea.text, Accent)
        If pos > 0 Then
            If ea.ChrByIndex(pos + 1) = "(" Then
                ea.SetNormalBrackets
                s = ea.GetNextBracketContent(pos)
                ea.text = Left(ea.text, pos - 1) & aString & s & bString & right(ea.text, Len(ea.text) - ea.pos + 1)
            Else
                s = ea.GetNextVar(pos)
                ea.text = Left(ea.text, pos - 1) & aString & s & bString & right(ea.text, Len(ea.text) - ea.pos)
            End If
        End If
    Loop While pos > 0
    ConvertPreAccent = ea.text
End Function
Function ConvertToLatexMaxima(Optional text As String) As String
' converts to Latex using Maxima
    If text <> "" Then Kommando = text
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
Function ConvertLatexToWord(text As String) As String
    Dim ea As New ExpressionAnalyser
    Dim p As Integer
    Dim t As String, n As String

    text = Replace(text, VBA.ChrW(11), "")    ' shift-enter
    text = Replace(text, vbCrLf, "")    ' enter
    text = Replace(text, VBA.ChrW(10), "")    ' shift-enter
    text = Replace(text, VBA.ChrW(13), "")    ' shift-enter
    text = Replace(text, "\infty", VBA.ChrW(8734))    '
    text = Replace(text, "\cdot", VBA.ChrW(183))    'prik gange

    'græske bogstaver
    text = Replace(text, "\Delta", VBA.ChrW(916))
    text = Replace(text, "\delta", VBA.ChrW(948))
    text = Replace(text, "\alpha", VBA.ChrW(945))
    text = Replace(text, "\beta", VBA.ChrW(946))
    text = Replace(text, "\gamma", VBA.ChrW(915))
    text = Replace(text, "\theta", VBA.ChrW(952))
    text = Replace(text, "\Theta", VBA.ChrW(920))
    text = Replace(text, "\lambda", VBA.ChrW(955))
    text = Replace(text, "\Lambda", VBA.ChrW(923))
    text = Replace(text, "\rho", VBA.ChrW(961))
    text = Replace(text, "\phi", VBA.ChrW(981))
    text = Replace(text, "\Phi", VBA.ChrW(934))
    text = Replace(text, "\varepsilon", VBA.ChrW(949))
    text = Replace(text, "\epsilon", VBA.ChrW(1013))
    text = Replace(text, "\psi", VBA.ChrW(968))
    text = Replace(text, "\Psi", VBA.ChrW(936))
    text = Replace(text, "\sigma", VBA.ChrW(963))
    text = Replace(text, "\Sigma", VBA.ChrW(931))
    text = Replace(text, "\mu", VBA.ChrW(956))
    text = Replace(text, "\Ohm", VBA.ChrW(937))
    text = Replace(text, "\Omega", VBA.ChrW(937))
    text = Replace(text, "omega", VBA.ChrW(969))
    text = Replace(text, "\Xi", VBA.ChrW(926))
    text = Replace(text, "\xi", VBA.ChrW(958))
    text = Replace(text, "\Chi", VBA.ChrW(935))
    text = Replace(text, "\chi", VBA.ChrW(967))
    text = Replace(text, "\tau", VBA.ChrW(964))
    text = Replace(text, "\Pi", VBA.ChrW(928))
    '    text = Replace(text, "\greek-nu", VBA.ChrW(957))
    text = Replace(text, "\kappa", VBA.ChrW(954))
    text = Replace(text, "\zeta", VBA.ChrW(950))
    text = Replace(text, "\eta", VBA.ChrW(951))    ' skal være sidst da eta indgår i andre
    text = Replace(text, "\inc", VBA.ChrW(8710))  ' specielt delta increment
    '    text = Replace(text, "degC", VBA.ChrW(8451))   ' specielt oC tegn
    '    text = Replace(text, "Symangle", VBA.ChrW(8736))  '

    ea.SetTuborgBrackets
    ea.text = text
    p = InStr(ea.text, "\frac")
    Do While p > 0
        t = ea.GetNextBracketContent
        n = ea.GetNextBracketContent
        ea.text = Left(ea.text, p - 1) & "(" & t & ")/(" & n & ")" & right(ea.text, ea.Length - ea.pos + 1)
        p = InStr(ea.text, "\frac")
    Loop

    ConvertLatexToWord = ea.text

End Function
Sub ReadSelection()
' indsætter det markerede som kommando
    Dim mtext As String, copytext As String
    Dim antal As Integer
    Dim sstart As Long
    Dim sslut As Long
    Dim j As Integer, antalgenstart As Integer
    On Error GoTo Slut

    ' hvis bold giver det buildup problemer
    For j = 1 To Selection.OMaths.Count
        Selection.OMaths(j).Range.Font.Bold = False
    Next
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
        Selection.OMaths(1).Range.Select    ' ovenstående var ikke nok, men hvorfor var det der?

        antalgenstart = 0
genstart:
        Kommando = ReadEquationCopy()
        If Kommando = "" And antalgenstart = 0 Then    ' forsøg på at rette sjældent problem hvor alt læses som "" måske skyldes det tempdoc
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
        '        Selection.OMaths(j).Range.Select ' ovenstående var ikke nok, men hvorfor var det der?
        '        Kommandoer(j - 1) = Replace(ReadEquationCopy(), Chr(11), "") ' fjern shift-enter
        '    Next

        '    GoTo hop ' forsøg

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
            Kommandoer(j) = Replace(mathcol(j + 1).Range.text, ChrW(11), "")    ' fjern shift-enter
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
' Oversætter selection der er omath til streng
' kan gøres hurtigere hvis der kopieres indenfor samme document, men virker ikke 100%
' autotextentries kan kopiere istedet for clipboard, og er måske hurtigere, men det er meget tvivlsomt.
' har dog oplevet at det giver problemer. Ligningssystem tan(30)=h/x og tan(25)=h/(x+34) på løsningen beregnes i enden. Det går helt galt, men virker når autotextentries ikke anvendes

    Dim sr As Range, r As Range
    If ir Is Nothing Then
        Set ir = Selection.Range
    End If

    Set sr = Selection.Range
    Set at = ActiveDocument.AttachedTemplate.AutoTextEntries.Add(Name:="WordMat", Range:=ir)

    Set r = tempDoc.Range
    at.Insert r, True

    With r
        .OMaths.BuildUp
        .OMaths.Linearize
        .OMaths(1).ConvertToNormalText
    End With

    ReadEquation = r.text
    r.Delete    ' dette afhjælper muligvis at word bliver langsom

    sr.Select

End Function
Function ReadEquationCopy(Optional ir As Range) As String
' bruger selection.copy som nok er langsommere end autotextentries og fylder clipboard
' Oversætter selection der er omath til streng
    Dim sr As Range, r As Range
    '    Dim t As Variant
    '    t = Timer
    If ir Is Nothing Then

    End If

    '    copytext = Selection.Range.text ' forsøg med ikke at klippe og kopiere problemet er at man er nødt til at konvertere til text for at læse ordentlig
    Set sr = Selection.Range

    Selection.Copy

    Set r = tempDoc.Range


    With r
        .Paste
        .OMaths.BuildUp
        .OMaths.Linearize
        .OMaths(1).ConvertToNormalText
    End With

    ReadEquationCopy = r.text
    sr.Select

    r.Delete
    '    ClearClipboard

    '    tempDoc.Range.Delete ' dette afhjælper muligvis at word bliver langsom
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
        .OMaths.BuildUp
        .OMaths.Linearize
        .OMaths(1).ConvertToNormalText
    End With

    ReadEquation2 = r.text
    sr.Select

    tempDoc.Range.Delete    ' dette afhjælper muligvis at word bliver langsom

End Function
Sub ReadSelectionOld()
' indsætter det markerede som kommando
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
        If Len(Selection.text) < 2 Then
            Kommando = ""
            MsgBox "Marker det der skal evalueres af Maxima.", vbOKOnly, "Vælg"
            Exit Sub
        End If
        Kommando = Selection.text
        Selection.Collapse (wdCollapseEnd)
    ElseIf antal = 1 Then
        If Len(Selection.text) < 3 Then
            Selection.OMaths(1).Range.Select
        End If
        If Selection.OMaths(1).Range.End < Selection.Range.End Then
            Selection.OMaths(1).Range.Select
        End If

        Selection.OMaths.BuildUp
        Selection.OMaths.Linearize
        Selection.OMaths(1).ConvertToNormalText
        Kommando = Selection.text
        Selection.text = Replace(Selection.text, VBA.ChrW(8727), MaximaGangeTegn)    'fjern *
        Selection.text = Replace(Selection.text, "*", MaximaGangeTegn)    'fjern *
        If DecSeparator = "," Then
            Selection.text = Replace(Selection.text, ".", ",")    'Når komma er decseparator kan . ændres til komma
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
            Kommandoer(i) = Replace(mathcol(i).Range.text, ChrW(11), "")
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
    mc(mc.Count).ParentOMath.Range.Select    ' parentomath nødvendig hvis cursor står i lyseblå firkant
    On Error GoTo Slut
    '        mc(mc.Count).Range.Select  ' virker med word 2010, parentomath giver tilgengæld problemer. Hmm problem med valgt del af udtryk og reducer
Slut:
    Selection.Collapse wdCollapseEnd
    Dim r As Range
    Set r = Selection.Range
    r.MoveStart wdCharacter, -1
    If r.text = VBA.ChrW(11) Then    ' hvis der er shift-enter i slutningen erstattes med alm. retur
        r.text = VBA.ChrW(13)
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
        mc(mc.Count).Range.Select  ' virker med word 2010, parentomath giver tilgengæld problemer. Hmm problem med valgt del af udtryk og reducer
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
    If r.text = VBA.ChrW(11) Then    ' hvis der er shift-enter i slutningen erstattes med alm. retur
        r.text = VBA.ChrW(13)
    End If
End Sub
Sub InsertMaximaOutput()
' indsætter maxima output i word document
    Dim mo As Range

    If MaximaKommando Then
        Set mo = Selection.OMaths.Add(Selection.Range)
        Selection.TypeText text:=Kommando
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
    Selection.TypeText MaximaOutput
    mo.OMaths.BuildUp

    InsertAssumptions

    If Len(KommentarOutput) > 3 Then
        '        Selection.Collapse (wdCollapseEnd)
        Selection.TypeParagraph
        Selection.InsertAfter TranslateReplaceComment(KommentarOutput)
        '        Selection.TypeText KommentarOutput
        Selection.Collapse (wdCollapseEnd)
    End If

    '        Selection.TypeParagraph

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



    'comm = Replace(comm, "solve: using arc-trig functions to get a solution.", "Da det er en trigonometrisk ligning er der uendeligt mange løsninger. Kun en af disse løsninger er angivet.")
    'comm = Replace(comm, "Some solutions will be lost.", " Alle løsninger er ikke fundet.")
    'comm = Replace(comm, "A number was found where a variable was expected -`solve'", "Du har bedt om at løse en ligning for en variabel der er defineret til en værdi. Lav en 'slet def:'  kommando før du løser for denne variabel.")
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
        Dim pos As Integer
        Dim ea As New ExpressionAnalyser
        pos = InStr(KommentarOutput, "Assumptions:")
        ea.text = KommentarOutput
        ea.pos = pos
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
    Selection.TypeText "Løsningsbetingelser:  "
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

Function FindMaximaOutput(text As String) As String
' plukker output2 fra hele maxima output
    Dim opos, ipos As Integer
    Dim text2 As String
    On Error GoTo Fejl

    opos = InStr(text, "(%o2)")
    ipos = InStr(text, "(%i3)")
    'MsgBox text
    If opos < 1 Then GoTo Fejl

    text2 = Trim(Mid(text, opos + 5, ipos - opos - 5))
    text2 = Replace(text2, "[", "")
    text2 = Replace(text2, "]", "")
    'text2 = Mid(text2, 2, Len(text2) - 3) ' fjern kantede paranteser

    FindMaximaOutput = text2

    GoTo Slut
Fejl:
    opos = InStr(text, "(%i2)")
    text2 = right(text, Len(text) - opos - 4)
    text2 = Replace(text2, "(%i2)", "")
    text2 = Sprog.A(108) & vbCrLf & vbCrLf & text2
    MsgBox text2, vbOKOnly, Sprog.A(109)
    FindMaximaOutput = ""
Slut:

End Function
Sub FindLastMaximaOutput(Optional ConvertToWord As Boolean = True)
' plukker sidste output fra hele maxima output
Dim same As Boolean
    GenerateOutputArray
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
    '    KommentarOutput = "Der kom følgende fejlmelding fra Maxima:" & vbCrLf & vbCrLf & KommentarOutput
    'End If

    GoTo Slut
Fejl:
Slut:
End Sub
Sub GenerateOutputArray()
    Dim text As String
    Dim Output As String
    Dim nr As Integer
    Dim AntalCifre As Integer
    Dim pos As Long
    Dim pos2 As Long
    Dim posb As Long
    Dim ArrIndex As Integer 'mac
    ArrIndex = 1 'mac

    On Error GoTo Fejl
    text = MaxProc.LastMaximaOutput
    'text = Right(text, Len(text) - 270)
    ReDim Minput(1 To 100) As String
    ReDim moutput(1 To 100) As String
    Dim slutnu As Boolean
    'MsgBox text
    
'    pos = InStr(1, text, vbLf & "(%")
'    pos2 = InStr(pos + 2, text, "(%") - 1
   pos = FindNextio(1, text) - 1
   pos2 = FindNextio(pos + 2, text) - 1

    'pos2 = InStr(pos + 2, text, vbLf & "(%")
    Do
        If pos < 1 Then
            GoTo Slut
        End If
        If pos2 < 1 Then
            pos2 = Len(text)
            slutnu = True
        End If
        posb = InStr(pos, text, ")")

        nr = CInt(Mid(text, pos + 4, posb - pos - 4))
        If nr < 10 Then
            AntalCifre = 1
        ElseIf nr < 100 Then
            AntalCifre = 2
        ElseIf nr < 1000 Then
            AntalCifre = 3
        Else
            AntalCifre = 4
        End If

        Output = Mid(text, posb + 1, pos2 - pos - 4 - AntalCifre)
        If Mid(text, pos + 3, 1) = "o" Then Output = RemovePrintetBool(Output)
        Output = Replace(Trim(Output), VbCrLfMac, "")
'#If Mac Then ' SBCL
        ArrIndex = ArrIndex + 1 'mac
        If Mid(text, pos + 3, 1) = "i" Then
            ' næste er input
            Minput(ArrIndex) = Minput(ArrIndex) & Output
        Else
            ' næste er output
            moutput(ArrIndex) = moutput(ArrIndex) & Output
        End If
'#Else ' GCL
'        If Mid(text, pos + 3, 1) = "i" Then
'            ' næste er input
'            Minput(nr) = Minput(nr) & Output
'        Else
'            ' næste er output
'            moutput(nr) = moutput(nr) & Output
'        End If
'#End If
        pos = pos2
'        pos2 = InStr(pos + 3, text, vbLf & "(%")
      pos2 = FindNextio(pos + 3, text) - 1
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
Function FindNextio(StartPos As Long, text As String) As Long

On Error GoTo Fejl
Dim pos As Long, pos2 As Long, io As String
Dim ea As New ExpressionAnalyser

FindNextio = 0
ea.text = text
Do
   pos = InStr(StartPos, text, "(%")
   If pos <= 0 Then Exit Function
   io = ea.ChrByIndex(pos + 2)
   If io = "o" Or io = "i" Then
      If ea.IsNumber(pos + 3) Then
         FindNextio = pos
         Exit Function
      End If
   End If
   StartPos = pos + 2
Loop While StartPos < Len(text)
   
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
    Dim arr As Variant
    Dim t As String, ant As Integer, i As Integer
    On Error GoTo Fejl
    s = TrimR(Trim(s), vbLf)
    s = TrimR(Trim(s), vbCr)
'#If Mac Then ' vbcrlfmac can ikke bruges her, da det er lf på mac
    arr = Split(s, vbLf)
'#Else
'    arr = Split(s, vbLf)
'#End If
    ant = UBound(arr)
    If ant > 0 Then
        t = Trim(arr(ant))
        If t = "true" Or t = "false" Or t = "unknown" Then
            s = vbNullString
            For i = 0 To ant - 1
                s = s & arr(i)
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
    Dim text As String
    Dim Output As String
    Dim nr As Integer
    Dim AntalCifre As Integer
    Dim pos As Integer
    Dim pos2 As Integer
    Dim posb As Integer
    
    Dim MinputU() As String
    Dim moutputU() As String
    

    On Error GoTo Fejl
#If Mac Then
    text = MaxProc.LastMaximaOutput
#Else
    text = MaxProcUnit.LastMaximaOutput
#End If
    ReDim MinputU(1 To 100) As String
    ReDim moutputU(1 To 100) As String
    Dim slutnu As Boolean
'    MsgBox text

    pos = InStr(1, text, vbLf & "(%")
    pos2 = InStr(pos + 2, text, "(%") - 1
    Do
        If pos < 1 Then
            Exit Function
        End If
        If pos2 < 1 Then
            pos2 = Len(text)
            slutnu = True
        End If
        posb = InStr(pos, text, ")")

        nr = CInt(Mid(text, pos + 4, posb - pos - 4))
        If nr < 10 Then
            AntalCifre = 1
        ElseIf nr < 100 Then
            AntalCifre = 2
        Else
            AntalCifre = 3
        End If
        Output = Replace(Trim(Mid(text, posb + 1, pos2 - pos - 4 - AntalCifre)), vbCrLf, "")

        If Mid(text, pos + 3, 1) = "i" Then
            ' næste er input
            MinputU(nr) = MinputU(nr) & Output
        Else
            ' næste er output
            moutputU(nr) = moutputU(nr) & Output
        End If
        pos = pos2
        pos2 = InStr(pos + 3, text, vbLf & "(%")
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
Sub FindVariable(Optional text As String, Optional MaximaCode As Boolean = True)
' finder variable i texten og putter dem i Vars-string
    Dim var As String
    Dim pos As Integer
    Dim i As Integer
    Dim j As Integer
    Dim arrint As Variant
    AntalVars = 0
    vars = ""
    IntegrationVars = ""
    If text = "" Then
        text = Kommando
        If text = "" Then
            On Error GoTo Fejl
            For j = 0 To UBound(Kommandoer)
                text = text & "+" & Kommandoer(j)    ' der skal bare være tegn mellem som ikke kan indgå i variabel
            Next
        End If
    End If
    Dim ea As New ExpressionAnalyser
    If MaximaCode Then
        ea.text = CodeForMaxima(text)
    Else
        ea.text = text
    End If
    ea.pos = 1
    Dim vararr(30) As String    ' gemmer variablene kun til sammenligning
    Dim varind As Integer
    Dim varexists As Boolean

    ' først tilsæt integrationsvariable
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
        If ea.ChrByIndex(ea.pos + 1) <> "(" And ea.ChrByIndex(ea.pos + 1) <> "'" And ea.ChrByIndex(ea.pos + 1) <> VBA.ChrW(8289) And Not (InStr(defstring, var & ":") = 1 Or InStr(defstring, "$" & var & ":") > 0) And InStr(var, "qkqg") < 1 Then
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
        ea.pos = ea.pos + 1
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
        s = ActiveDocument.Range.OMaths(i).Range.text
        '        If ActiveDocument.Range.OMaths(i).Range.start > start Then GoTo slut

        If InStr(s, VBA.ChrW(8788)) > 0 Or InStr(s, VBA.ChrW(8797)) > 0 Or InStr(s, ":") > 0 Or InStr(s, VBA.ChrW(8801)) > 0 Then    ' kun se på felter med :  tripelligmed def= og specielt := tegn
#If Mac Then
            s = ReadEquation(ActiveDocument.Range.OMaths(i).Range) ' kan ikke huske hvad forskel på readeq1 og 2 er. Ihvert et tilfælde på mac numerisk grafisk ligningsløsning med foregående definition fejler. så derfor alm. readeq på mac.
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
    defstring = FindDefinitions
    Selection.start = start
    Selection.End = sslut
    '    Selection.Collapse (wdCollapseStart)

End Function
Function AddDefinition(s As String) As String
   Dim t As String, i As Integer, Expr As String, c As String, j As Integer, pos As Integer, pos2 As Integer, k As Integer
   Dim ea As New ExpressionAnalyser
   Dim ea2 As New ExpressionAnalyser
   Dim defenk As String, ass As String
   Dim arr As Variant, arr2() As String, arr3() As String
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
            AddDefinition = "kill(" & CodeForMaxima(Replace(s, ",", ";")) & ")$" & AddDefinition
         End If
      Else    ' gem definitioner
         '            FindDefinitions = CodeForMaxima(s) & "," & FindDefinitions
         defstringtext = defstringtext & s & "$"
         ea.text = Replace(s, ";", " , ")    ' forsøg med at fjerne semikolon
         ea.ConvertDecSeparator
         ea.pos = 1
         Do
            defenk = Trim(ea.GetNextListItem(ea.pos, ","))
            If Len(defenk) = 0 Then GoTo hop
            If InStr(defenk, "block") > 0 Then
               defenk = Replace(defenk, ";", ",")
               defenk = Replace(defenk, VBA.ChrW(9633), "")    ' tegn foran paranteser der gør dem usynlige. Muligvis problematisk da paranteserne ikke fjernes. gør også brøker små
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
            
               ea2.text = defenk
               ea2.SetNormalBrackets
               i = 0
               Do
                  pos = InStr(ea2.text, "{" & VBA.ChrW(9608))
                  If pos > 0 Then
                     ea2.text = Replace(ea2.text, ")=", "):=", 1, pos)
                     pos = InStr(ea2.text, "{" & VBA.ChrW(9608))
                     pos2 = InStr(pos, ea2.text, VBA.ChrW(9508))
                     
                     Expr = Mid(ea2.text, pos + 2, pos2 - pos - 1)
                     Expr = ea2.GetNextBracketContent(pos)
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
                           Do While k <= UBound(arr3) ' hvis delt op forkert så repareres
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
                           t = t & sqbrAssump(SplitAssump(Trim(ass))) & " then " & CodeForMaxima(arr3(0))
                        Else
                           t = t & " 0=0 then " & CodeForMaxima(arr3(0))
                        End If
                        If j < UBound(arr2) Then t = t & " elseif "
                     Next
                     ea2.text = Left(ea2.text, pos - 1) & t & ")"  ' & Right(ea2.Text, Len(ea2.Text) - pos - Len(expr) - 6)
                     i = i + 1
                  End If
               Loop While pos > 0 And i < 1000
               defenk = ea2.text
            
            
            Else
               defenk = Replace(defenk, VBA.ChrW(34), "")    ' apostrof
               defenk = CodeForMaxima(defenk)
'#If Mac Then ' på mac bruges ## som multiplikator ved vektorer, så ikke nødvendigvis 'ikke ligmed'  nu også på windows 1.18
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
            arr = Split(defenk, ":")
            If UBound(arr) = 0 Then
               '                    If InStr(defenk, "setunits") > 0 Or InStr(defenk, "assume") > 0 Then
               GoTo hop2
               '                    End If
            End If
            If Not ValidateDef(arr(0), arr(1)) Then GoTo hop

            If Left(arr(1), 1) = "=" Then arr(1) = right(arr(1), Len(arr(1)) - 1)
            If arr(1) = "" Then GoTo hop
            If UBound(ldefname) < defindex Then
               ReDim Preserve ldefname(defindex)
               ReDim Preserve ldefvalue(defindex)
            End If
            ldefname(defindex) = arr(0)
            ldefvalue(defindex) = arr(1)
            defindex = defindex + 1
hop2:                    ' hvis setunit eller assume eller declare
            ea2.text = defenk
            ea2.pos = 1
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
         Loop Until ea.pos >= Len(ea.text)
      End If



      ' find var navne til sletning næste kørsel


      '        arr = Split(s, ListSeparator)
      '        For j = 0 To UBound(arr)
      '            s = arr(j)
      '            ea.text = CodeForMaxima(s)
      '            ea.pos = 1
      '            s = ea.GetNextVar
      '            If s = "e" Then
      '                MsgBox "bogstavet e er reserveret og kan ikke defineres.", vbOKOnly, "Fejl"
      '            GoTo fejl
      '            End If
      '            If s <> "" Then
      '                KillDef = KillDef & s & ","
      '            End If
      '        Next

      '    Selection.Collapse (wdCollapseEnd)
   End If

End Function
Function SplitAssump(s As String) As String
   Dim ea As New ExpressionAnalyser
   Dim v As String, p As Integer, p2 As Integer
   Dim pre As String, pst As String
   
   ea.SetNormalBrackets
   
    s = Replace(s, VBA.ChrW(8804), "<=")    '
    s = Replace(s, VBA.ChrW(8805), ">=")    '

   ea.text = s
   
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
   Dim pre As String, pst As String, arr() As String, l1 As String, l2 As String
   t = s
   p = InStr(s, VBA.ChrW(8712))
   If p <= 0 Then
      sqbrAssump = s
      Exit Function
   End If
   Set ea = New ExpressionAnalyser
   ea.text = s
   v = ea.GetNextVar()
   If v = "" Then
      sqbrAssump = s
      Exit Function
   End If
   ea.pos = p
   s = right(s, Len(s) - p)
   ea.text = s
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
   ea.pos = 1
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
   arr = Split(s, c)
   If UBound(arr) = 1 Then
      sqbrAssump = v & l1 & arr(0) & " and " & arr(1) & l2 & v
   Else
      sqbrAssump = t
   End If

End Function
Sub ResetDefinitions()
    ReDim ldefname(100)
    ReDim ldefvalue(100)
    defindex = 0

End Sub
Function ValidateDef(ByVal lhs As String, ByVal RHS As String) As Boolean
' returnerer false hvis der er fejl
    Dim ea As New ExpressionAnalyser
    Dim pos As Integer
    On Error GoTo Fejl
    ValidateDef = False

    pos = InStr(lhs, "[")    ' Man kan skrive alt muligt i subscript/index
    If pos > 0 Then
        lhs = Left(lhs, pos - 1)
    End If

    ' If Not ea.IsAllText Then ikke god nok da , og ; ikke bliver taget højde for
    If InStr(lhs, "+") > 0 Or InStr(lhs, "-") > 0 Or InStr(lhs, "*") > 0 Or InStr(lhs, "/") > 0 Then
        MsgBox Sprog.A(117) & vbCrLf & vbTab & ConvertToAscii(lhs & "=" & RHS) & vbCrLf & Sprog.A(118), vbOKOnly, Sprog.Error
        Exit Function
    End If

    If right(RHS, 1) = "+" > 0 Then
        MsgBox Sprog.A(117) & vbCrLf & vbTab & ConvertToAscii(lhs & "=" & RHS) & vbCrLf & Sprog.A(118), vbOKOnly, Sprog.Error
        Exit Function
    End If


    ValidateDef = True
Fejl:
End Function
Function GetAssumeText(defenk) As String
' tager text fra word og omdanner til assume samt indsætter til forget
    Dim ea As ExpressionAnalyser
    Dim ea2 As ExpressionAnalyser
    Dim s As String
    Dim arr As Variant
    Dim p1 As Integer, p2 As Integer
    If Len(defenk) > 0 Then
        If Left(defenk, 6) = "assume" Then
            GetAssumeText = defenk & "$"
            Set ea = New ExpressionAnalyser
            ea.text = defenk
            ea.SetNormalBrackets
            defenk = ea.GetNextBracketContent(6)
            ForgetList = ForgetList & defenk & ","
        Else
            Set ea = New ExpressionAnalyser
            Set ea2 = New ExpressionAnalyser
            ea.text = defenk
            defenk = ""
            ea.SetNormalBrackets
            Do
                s = ea.GetNextListItem
                If InStr(s, "#") > 0 Then
                    arr = Split(s, "#")
                    s = "notequal(" & arr(0) & "," & arr(1) & ")"
                ElseIf Len(s) > 0 Then
                    ea2.text = s
                    p1 = ea2.FindNextGreaterOrLower(1)
                    p2 = ea2.FindNextGreaterOrLower(p1 + 1)
                    If p2 > 0 Then    ' 2<=x<=3
                        If ea.ChrByIndex(p1 + 1) = "=" Then
                            s = Left(s, p2 - 1) & "," & right(s, Len(s) - p1 - 1)
                        Else
                            s = Left(s, p2 - 1) & "," & right(s, Len(s) - p1)
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


Sub AddToKillDefs(defs As String)
    Dim arr As Variant
    Dim j As Integer
    Dim s As String
    If defs <> "" Then
        arr = Split(defs, ListSeparator)
        For j = 0 To UBound(arr)
            s = arr(j)
            s = Split(s, "=")(0)
            s = Split(s, "(")(0)
            KillDef = KillDef & s & ","
        Next
    End If

End Sub

Function IsSolved(ByVal Ligning As String, variabel As String) As Boolean
' checker om ligningen er helt løst for variabelen
' starter den med variabel= ?
' indgår variabel på højresiden af =tegn
    On Error GoTo Slut
    Dim NotSolved As Boolean
    Dim ea As New ExpressionAnalyser
    Dim var As String
    Dim arr As Variant
    Ligning = RemoveBrackets(Ligning)
    arr = Split(Ligning, "=")

    If UBound(arr) = 1 Then    ' der er præcis et =tegn
        ea.text = arr(1)    'højresiden
        IsSolved = True
        Do
            var = ea.GetNextVar
            ea.pos = ea.pos + 1
            If var = variabel Then
                IsSolved = False    'variabel fundet på højresiden så er det ihvertfald ikke løst
                GoTo Slut
            End If
        Loop Until var = ""

        If Left(Ligning, 1 + Len(variabel)) = variabel & "=" Then  '  ikke nok alene
            IsSolved = True
        ElseIf Len(variabel) > 1 And Mid(Ligning, Len(variabel) + 1, 1) = "=" Then    ' hvis længde af variabel samme som venstresiden. Nødvendig for variable som f.eks. 1+r da det kan ændres til r+1. Men nok ikke 100% sikker
            IsSolved = True
        Else    ' det må antages at den ikke er løst hvis den ikke passer med noget ovenfor
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

Function IsAllSolved(ByVal ligninger As String, ByVal variabel As String, Optional ByVal sep As String = ",") As String
' checker om en liste af ligninger er løst helt
' true hvis en er løst, false hvis ingen. All hvis alle.
    Dim oneSolved As Boolean
    Dim allsolved As Boolean
    Dim arr As Variant
    Dim i As Integer
    oneSolved = False
    allsolved = True

    If Not LmSet Then
        If InStr(ligninger, "%if(") > 0 Then GoTo Slut

        ligninger = RemoveBrackets(ligninger)
        arr = Split(ligninger, sep)

        For i = 0 To UBound(arr)
            If IsSolved(arr(i), variabel) Then
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
    Dim pos As Integer, pos2 As Integer, pos3 As Integer, pos4 As Integer
    Dim t As String
    Dim deffejl As Boolean
    Dim ndeferror As Integer

    If defstring <> "" Then
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
        pos = InStr(MaximaOutput, "incorrectsyntax:Missing")
        fejltekst = Sprog.SyntaxError & ". " & vbCrLf & Sprog.Missing & " " & Mid(MaximaOutput, pos + 22, 1)
        deffejl = True
        CheckForError = True
    ElseIf InStr(MaximaOutput, "incorrectsyntax:Toomany") > 0 Then
        pos = InStr(KommentarOutput, "incorrectsyntax:Toomany")
        fejltekst = Sprog.SyntaxError & ". " & vbCrLf & Sprog.TooMany & " " & Mid(MaximaOutput, pos + 29, 1)
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
        fejltekst = Sprog.A(133) '"Du har bedt om at løse ligningen for en variabel der allerede er defineret. Indsæt en 'slet def:' kommando før ligningen"
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
            UFerror.TextBox_definitioner.text = DefinitionsNice
        Else
            UFerror.TextBox_definitioner.visible = False
            UFerror.Label_definitioner.visible = False
        End If
        UFerror.Label_fejltekst.Caption = fejltekst
        UFerror.Show
    End If

End Function

Function GetErrorText(text As String) As String
    Dim pos As Integer, pos2 As Integer, pos3 As Integer, pos4 As Integer
    Dim t As String
    Dim l As Integer
    On Error Resume Next
    l = Len(text)
    pos = InStr(MaximaOutput, "incorrectsyntax")
    pos2 = InStr(pos, MaximaOutput, text)
    pos4 = InStr(pos2 + l, MaximaOutput, "^")
    If pos4 < 1 Then
        pos4 = Len(MaximaOutput)
    End If
    t = Mid(MaximaOutput, pos2 + l, pos4 - pos2 - l + 1)
    t = Replace(t, "^", vbCrLf & "    ^", 1, 1)
    GetErrorText = Sprog.SyntaxError & vbCrLf & Sprog.IllegalSymbol & ":" & vbCrLf & t

End Function
Function DefinitionsNice() As String
    Dim defs As String
    defs = defstring
    If Len(defs) > 3 Then
        '    defs = Mid(defs, 2, Len(defs) - 3)
        defs = ConvertToAscii(defs)
        defs = Replace(defs, "$", vbCrLf)
        defs = Replace(defs, ":=", vbTab & "= ")
        defs = Replace(defs, ":", vbTab & "= ")
        If DecSeparator = "," Then
            defs = Replace(defs, ",", ";")
            defs = Replace(defs, ".", ",")
        End If
    End If

    DefinitionsNice = defs
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


