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
Public Kommando As String    ' command that maxima should execute
Public StartKommando As String
Public KommandoForklaring As String
Private Kommandoer() As String    ' commands that maxima should execute. KommandoerString must be used public
Public MaximaOutput As String    ' result from Maxima
Public KommentarOutput As String    ' any comment from Maxima
Public Vars As String    ' variables in expressions separated by ;
Public AntalVars As Integer    ' number of variables found by findvariable
Private IntegrationVars As String    ' variables used for integration should not be included in the var list
Public DiffEqVar As String    ' variable for differential equations. ensures that subst is not used in diff
Public MaximaInstalled As Boolean
Public CmdTask As Task
Public StopNow As Boolean
Public TempDefs As String    ' temporary definitions for this command only
Public DefFejl As Boolean    ' set if there are errors in definitions
Public Definitions As Boolean    ' set if there are definitions in the document
Public DefString As String    ' contains the output from finddefinition
Public defstringtext As String    ' contains definitions from finddefinitions before ConvertToMaxima
Public DefColl As Collection
Private ldefname() As String    ' array with all definitions
Private ldefvalue() As String    ' array with all definitions
Private MathOArr() As OMath
Public defindex As Integer
Public KillDef As String    'contains the list separator list of last definitions
Public ForgetList As String    ' contains list of assumes to be deleted
Private Minput() As String
Private moutput() As String
Public ConvertErrorText As String    ' if an error occurs while trying to translate syntax, this contains the error text
Public MaximaInputStreng As String
Public MaximaInputStrengSec As String    ' secondary inputstring to be run after question
Private matvekt As Boolean    ' is set if matrices are registered. Then dot and cross product are activated.
Public matrixstartbracket As String
Public matrixendbracket As String
Public ConvertLnLog As Boolean    ' whether converttomaxima converts ln(x) -> log(x) and log(x)-> log(x)/log(10)
Public prevspr As String    ' used when Maxima asks repeatedly about the same thing
Public prevsvar As String
Public MaximaSti As String

Private Sub Class_Initialize()
' Finds Maximapath. seeks in Appdata and programfiles. If Maxima exists both places then use newest version

    On Error Resume Next
    ConvertLnLog = True
#If Mac Then
    MaximaInstalled = True
#Else
    
    MaximaSti = GetMaximaPath()
    
    If MaximaSti = vbNullString Then
        MaximaInstalled = False
        MsgBox TT.A(888), vbOKOnly, TT.Error
        GoTo slut
    Else
        MaximaInstalled = True
    End If
    
    MaxProc.SetMaximaPath MaximaSti
    
#End If

slut:
End Sub
Public Function PrepareNewCommand(Optional FindDef As Boolean = True) As Boolean
    Dim DefS As String
    On Error Resume Next
    StopNow = False
    DefFejl = False
    DiffEqVar = vbNullString
    Kommando = vbNullString
    matvekt = False
    matrixstartbracket = "("
    matrixendbracket = ")"
    ReDim Kommandoer(1 To 1)
    ' kommandoer?
    MaximaOutput = vbNullString

    KommentarOutput = vbNullString
    Vars = vbNullString
    IntegrationVars = vbNullString
    AntalVars = 0
    TempDefs = vbNullString
    ReDim Kommandoer(0 To 0)
    ReDim Minput(1 To 100) As String
    ReDim moutput(1 To 100) As String
    MaximaInputStreng = vbNullString
    MaximaInputStrengSec = vbNullString
    ConvertErrorText = vbNullString

    If FindDef Then
        Definitions = False
        InsertKillDef
        DefS = FindDefinitions
        If DefS = "error" Then
            PrepareNewCommand = False
            GoTo slut
        End If
        If Len(DefS) > 0 Then
'            MaximaInputStreng = MaximaInputStreng & "[" & left$(Replace(DefS, "$", ","), Len(DefS) - 1) & "]$"
            MaximaInputStreng = Replace(DefS, "$", ";")
        End If
        MaximaInputStreng = MaximaInputStreng & "fejl;"
    End If

    PrepareNewCommand = True
    GoTo slut

fejl:
    MsgBox TT.ErrorGeneral & vbCrLf & " preparing maxima", vbOKOnly, TT.Error
    PrepareNewCommand = False
slut:

End Function
Private Sub Class_Terminate()
'Closes commandprompt
    On Error Resume Next
End Sub

Public Sub ResetMaxima()
' Preparemaxima also ensures variables are forgotten, but there is no time penalty for this extra call, and killdef/forgetlist are set to "" hence the next call is smaller
    Dim kd As String
    
    If DllConnType = 2 Then Exit Sub ' wsh starts anew every time so reset is not necessary
    
#If Mac Then
    KillDef = ""
    ForgetList = ""
    Exit Sub
#Else
    If Len(KillDef) > 0 Then
        KillDef = Left$(KillDef, Len(KillDef) - 1)    'remove last comma
        kd = kd & "errcatch(kill(" & KillDef & "))"
        KillDef = ""
    End If
    If Len(ForgetList) > 0 Then
        ForgetList = Left$(ForgetList, Len(ForgetList) - 1)    'remove last comma
        If Len(kd) > 0 Then kd = kd & ","
        kd = kd & "forget([" & ForgetList & "])"
        ForgetList = ""
    End If
    MaxProc.Reset kd
#End If
End Sub
Public Sub RunMaxima(Optional StopTime As Integer, Optional Reset As Boolean = True, Optional ConvertWordSymbols As Boolean = True)
' starts maxima sends input and receives output

    IncreaseCalcCounter
#If Mac Then
    RunMaximaFile StopTime, ConvertWordSymbols
#Else
    Dim i As Integer, InputStreng As String
    If StopTime = 0 Then StopTime = 200

    '    MaximaInputStreng = "mu:2;sigma:1;Integrate(1/(sqrt(2*%pi)*sigma)*%e^(-1/2*((y-mu)/sigma)^2),y,minf,1000);"
    '     MsgBox (MaximaInputStreng)
    '    Selection.InsertAfter MaximaInputStreng
    
    If DllConnType = 2 Then
        RunMaximaWSH StopTime
        Exit Sub
    End If

    If MaxProc Is Nothing Then
        Set MaxProc = GetMaxProc() 'CreateObject("MaximaProcessClass")
        MaxProc.StartMaximaProcess
    End If

    InputStreng = GetMaximaSettingsString & MaximaInputStreng

    MaxProc.AntalCifre = MaximaCifre

    If MaximaComplex Then
        MaxProc.Complex = 1
    Else
        MaxProc.Complex = 0
    End If
    i = 0
    Dim UFwait2 As UserFormWaitForMaxima

    If StopTime < -2 Then    ' by graphs and some other nsolve
        '        MsgBox MaximaInputStreng
        MaxProc.ExecuteMaximaCommand InputStreng, 0
        WaitForMaximaUntil (-1 * StopTime)
        '        Wait (2)
        If StopNow Then GoTo slut
        If MaxProc.Question = 1 Then
            AnswerQuestion StopTime, False ' possibly problematic after changed to form can be displayed with answerquestion
            WaitForMaximaUntil (-1 * StopTime)
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
        Set UFwait2 = New UserFormWaitForMaxima
        UFwait2.Show vbModeless
        DoEvents
        MaxProc.ExecuteMaximaCommand InputStreng, 1
        FindLastMaximaOutput
        ResetMaxima
        Unload UFwait2
        '        MaxProc.CloseProcess ' causes it to hang a little, maybe necessary
        ShowDebug
        Exit Sub
        GoTo slut
    Else ' normal calculation
        If MaxProc.Finished = 0 Then
            WaitForMaximaUntil 100
            If MaxProc.Finished = 0 Then RestartMaxima
'            UFwait2.Show vbModeless
'            DoEvents
'            Do While MaxProc.Finished = 0 And i < 50
'                If Round(i / 8) = i / 8 Then
'                    UFwait2.Label_progress.Caption = UFwait2.Label_progress.Caption & "*"
'                End If
'                Wait (0.1)
'                i = i + 1
'            Loop
'            UFwait2.Label_tip.Font.Size = 10
'            UFwait2.Label_tip.Font.Italic = False
'            UFwait2.Label_tip.Caption = GetRandomTip
        End If
        If StopNow Then
            If Not UFwait2 Is Nothing Then Unload UFwait2
            MaxProc.CloseProcess
'            MaxProc.Finished = 1
            MaxProc.StartMaximaProcess
            Exit Sub
        End If

        MaxProc.ExecuteMaximaCommand InputStreng, 0
        If Not UFwait2 Is Nothing Then UFwait2.Label_progress.Caption = UFwait2.Label_progress.Caption & "*"
    End If
    
    WaitForMaximaForm StopTime

    AnswerQuestion StopTime

    ShowDebug
    '    MsgBox (MaxProc.LastMaximaOutput)

    If Reset Then
        FindLastMaximaOutput ConvertWordSymbols
        If MaximaOutput = vbNullString And KommentarOutput = vbNullString Then ' Something must be wrong with Maxima. Try restart
            If MaximaUnits Then
                MaxProcUnit.CloseProcess
                MaxProcUnit.StartMaximaProcess
            Else
                MaxProc.CloseProcess
                MaxProc.StartMaximaProcess
            End If
        Else
            ResetMaxima ' deletes definitions. does not slow WordMat
        End If
    End If

    If FreshMaxima Then
        MaxProc.ExecuteMaximaCommand "quit()$", 0
    End If
    On Error Resume Next
    If Not UFwait2 Is Nothing Then Unload UFwait2
    GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
#End If
End Sub
Sub RunMaximaWSH(Optional StopTime As Integer)
    Dim Moutp As String, cmd As String
    
    If Not QActivePartnership(False, True) Then
        If MsgBox2("Connecting to Maxima using WSH requires Partnership with Eduap" & vbCrLf & vbCrLf & "Do you want to change the setting to use the default option: registered dll?", vbYesNo, "Error") = vbYes Then
            DllConnType = 0
        End If
        Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    cmd = GetMaximaSettingsString & MaximaInputStreng
    If MaximaUnits Then
        Application.Run macroname:="ExecuteMaximaViaWSH", varg1:=cmd, varg2:=10, varg3:=True, varg4:=True
        If Err.Number = 513 Then
            Moutp = Err.Description
        End If
    Else
        Application.Run macroname:="ExecuteMaximaViaWSH", varg1:=cmd, varg2:=10, varg3:=False, varg4:=True
        If Err.Number = 513 Then
            Moutp = Err.Description
        End If
    End If
    If Err.Number = 513 Then
        FindLastMaximaOutput True, Moutp
        ShowDebug False, cmd, Moutp
    Else
        MsgBox2 TT.A(889), vbOKOnly, TT.Error
    End If
    Err.Clear
End Sub
Function ConvertMaximaUnits(Expr As String, Optional basedim As Boolean = True) As String
' Converts a maxima expression, using the maxima-image with units loaded
'#If Mac Then ' any of the following may work, but try this first
'    ConvertMaximaUnits = Expr
'#Else
    Dim i As Integer, k As String, arr As Variant, s As String
    
    Expr = Replace(Expr, """", "") ' The answer will typically be run through the num function which adds "" to numbers
    
    If basedim Then ' all units are reverted to base SI-units
        k = Expr & ",numer;applyb1(%,unitrule),numer;dectalallNum(%),numer;"
    Else
        k = "numer:true$uforget(append(globalbaseunitlisting,globalderivedunitlisting))$"
        If Not OutUnits = vbNullString Then
            arr = Split(OutUnits, ",")
            For i = 0 To UBound(arr) ' It has to be done this way because you cannot run setunits on a list where the same type of unit appears twice. e.g. eV and J
                k = k & "setunits(" & omax.ConvertUnits((arr(i))) & ")$"
            Next
        End If
'        k = k & expr & ",numer;applyb1(%,unitrule),numer;scanmap(processunits,%);dectalallNum(%),numer;"
        k = k & Expr & ",numer;ConvertUnit(%);"
        If MaximaDecOutType = 3 Then
            k = k & "scinotall(%);"
        Else
'            k = k & "dectalallNum(%),numer;" ' causes problems for units with % prefix as the units are then displayed first.
            k = k & "%,numer;"
        End If
    End If
'    MsgBox k
#If Mac Then
    k = GetMaximaSettingsString & "fpprintprec:" & MaximaCifre & "$" & k ' there are problems with dectalallNum, so this fppprintprec method is used here with units
    s = ExecuteMaximaViaFile(k, 3, True)

    ConvertMaximaUnits = GetMaximaUnitOutput(s) 'MaxProcUnit.LastMaximaOutput
#Else
    MaxProcUnit.ExecuteMaximaCommand k, 0
    WaitForMaximaUnitUntil
    ConvertMaximaUnits = GetMaximaUnitOutput() 'MaxProcUnit.LastMaximaOutput
        
'    MsgBox MaxProcUnit.LastMaximaOutput
#End If
    
'    ConvertMaximaUnits = ConvertToWordSymbols(ConvertMaximaUnits)
'    ConvertMaximaUnits = CodeForMaxima(ConvertMaximaUnits)
    
    If matvekt Then ' if there are vectors in the expression, the first pass will have removed the special operators and functions that were introduced
        ConvertMaximaUnits = Replace(ConvertMaximaUnits, ChrW$(215), "~")    ' vector produkt
'        ConvertMaximaUnits = left$(ConvertMaximaUnits, Pos - 1) & "vecmag(" & mid$(ea.text, Pos + 1, posb - Pos - 1) & ")" & right$(ea.text, Len(ea.text) - posb)
        ConvertMaximaUnits = Replace(ConvertMaximaUnits, "^", "^^")
        ConvertMaximaUnits = Replace(ConvertMaximaUnits, "*", " ## ")
        Dim ea As New ExpressionAnalyser
        ea.text = ConvertMaximaUnits
        ea.ReplaceVar "abs", "vecmag"
        ConvertMaximaUnits = ea.text
    End If

End Function
Sub WaitForMaximaForm(StopTime As Integer)
#If Mac Then
#Else
    Dim i As Integer
    Dim UFwait2 As New UserFormWaitForMaxima
    
    i = 0
    Do While MaxProc.Finished = 0 And i < 110
        Wait (0.01)
        i = i + 1
    Loop
    If MaxProc.Finished = 0 Then
        '    ufwait2.omax = omax
        On Error Resume Next
        UFwait2.Show vbModeless
        DoEvents
        On Error GoTo fejl
        i = 0
        Do While MaxProc.Finished = 0 And i < StopTime
            Wait (0.1)
            i = i + 1
            If InStr(MaxProc.LastMaximaOutput, "Lisp error:") > 0 Then 'MaxProc.Finished virker ikke altid
                MaxProc.CloseProcess
                Exit Do
            End If
            If Round(i / 8) = i / 8 Then
                UFwait2.Label_progress.Caption = UFwait2.Label_progress.Caption & "*"
            ElseIf Round(i / 60) = i / 60 Then
                UFwait2.Label_tip.Font.Size = 10
                UFwait2.Label_tip.Font.Italic = False
                UFwait2.Label_tip.Caption = GetRandomTip
            ElseIf Len(UFwait2.Label_progress.Caption) > 33 Then
                UFwait2.Label_progress.Caption = "*"
            End If
            If i = StopTime - 200 Then
                UFwait2.Label_tip.Caption = TT.A(890) & " 20s"
            ElseIf i = StopTime - 150 Then
                UFwait2.Label_tip.Caption = TT.A(890) & " 15s"
            ElseIf i = StopTime - 100 Then
                UFwait2.Label_tip.Caption = TT.A(890) & " 10s"
            ElseIf i = StopTime - 50 Then
                UFwait2.Label_tip.Caption = TT.A(890) & " 5s"
            ElseIf i = StopTime Then
                UFwait2.Label_tip.Caption = TT.A(891)
            End If
            If StopNow Then
                Unload UFwait2
                ShowDebug
                Dim tmis As String
                tmis = MaximaInputStreng
                RestartMaxima
'                MaxProc.CloseProcess
'                MaxProc.StartMaximaProcess
                Exit Sub
            End If
        Loop
        If i >= StopTime Then
            StopNow = True
            Unload UFwait2
            RestartMaxima
'            MaxProc.CloseProcess
            Exit Sub
        Else
            Unload UFwait2
        End If
    End If
    GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
#End If
End Sub
Sub ShowDebug(Optional vis As Boolean = False, Optional MInputString As String, Optional OutputString As String)
    If DebugWM Or vis Then
        If MInputString = vbNullString Then MInputString = MaximaInputStreng
        UserFormDebug.TextBox_input = MInputString
        If OutputString = vbNullString Then
#If Mac Then
#Else
            UserFormDebug.TextBox_lastmaximaoutput = MaxProc.LastMaximaOutput
            UserFormDebug.TextBox_maximaoutput = MaxProc.MaximaOutput
#End If
        Else
            UserFormDebug.TextBox_lastmaximaoutput = OutputString
            UserFormDebug.TextBox_maximaoutput = OutputString
        End If
        
        UserFormDebug.Label_time.Caption = UserFormDebug.Label_time.Caption & "Samlet tid: " & Timer - tid & vbCrLf
        UserFormDebug.Show
    End If
End Sub
#If Mac Then
#Else
Sub AnswerQuestion(StopTime As Integer, Optional ShowWaitForm As Boolean = True)
    Dim svar As String
    Dim answered As Boolean

    Do While MaxProc.Question = 1
        Dim pos As Integer
        Dim pos2 As Integer
        pos = InStr(MaxProc.LastMaximaOutput, "positive, negative or zero?")
        pos2 = InStr(MaxProc.LastMaximaOutput, "positive or negative?")
        If pos > 0 Or pos2 > 0 Then
            Dim UF As New UserFormAskSign
            If pos <= 0 And pos2 > 0 Then pos = pos2
            pos2 = InStr(MaxProc.LastMaximaOutput, "Is ")
            UF.Label_udtryk.Caption = Mid$(MaxProc.LastMaximaOutput, pos2 + 3, pos - pos2 - 4)
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
            svar = InputBox(MaxProc.LastMaximaOutput, "spr", "")    ' MaxProc.Question must be adjusted
        ElseIf InStr(MaxProc.LastMaximaOutput, "variables to solve for?") Then
            Dim antalvar As Integer, Var As String, ea As New ExpressionAnalyser, sprtext As String
            pos = InStr(MaxProc.LastMaximaOutput, " variables to solve for?")
            antalvar = val(Mid$(MaxProc.LastMaximaOutput, pos - 2, 2))
            sprtext = TT.A(416) & " " & antalvar & " " & TT.A(417)
            ea.text = Right$(MaxProc.LastMaximaOutput, Len(MaxProc.LastMaximaOutput) - pos - 24)
            ea.SetSquareBrackets
            ea.pos = 0
            Var = ea.GetNextBracketContent
igen:
            svar = InputBox(sprtext, "spr", Var)
            If svar = "" Then
                svar = "end;%;%;%"
            Else
                If UBound(Split(svar, ",")) = antalvar - 1 Then
                    svar = "[" & svar & "];%;%;%"
                Else
                    sprtext = TT.A(418) & " " & antalvar & " " & TT.A(419)
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
        defindex = 0 ' When answering the initial definition is lost from lastmaximaoutput. definition errors will not be detected. If this variable is not reset. WordMat will think a defition error has occured.
    Loop

    If answered And MaximaInputStrengSec <> "" Then
        MaximaInputStrengSec = "%th(2);" & MaximaInputStrengSec    ' because of slutprut must be skipped.
        If Right$(MaximaInputStrengSec, 1) = ";" Then    ' at answerquestion is set automatically;
            MaximaInputStrengSec = Left$(MaximaInputStrengSec, Len(MaximaInputStrengSec) - 1)
        End If
        MaxProc.AnswerQuestion (MaximaInputStrengSec)
        WaitForMaximaForm StopTime
    End If

End Sub
#End If

Public Sub RunMaximaFile(Optional StopTime As Integer, Optional ConvertWordSymbols As Boolean = True)
    ' for mac. Starts maxima, sends command, receives output. closes maxima
#If Mac Then

    Dim s As String, WT As String, Var As String
    Dim pos As Integer, pos2 As Integer, Assume As String, InputString As String

    '    Dim FN As String
    '    Dim UfWait2 As UserFormWaitForMaxima
    If StopTime > 0 Then
        WT = StopTime
    Else
        WT = "3"
    End If
    '    Set UfWait2 = New UserFormWaitForMaxima
    '    On Error Resume Next ' if a form is already open it gives an error
    '    UfWait2.Show vbModeless
    '    UfWait2.Label_tip.Font.Size = 10
    '    UfWait2.Label_tip.Font.Italic = False
    '    On Error GoTo 0
start:
    '    UfWait2.Label_tip.Caption = GetRandomTip
    DoEvents
    '  FN = DataFolder & "output.txt"
    InputString = GetMaximaSettingsString & Assume & MaximaInputStreng
    s = ExecuteMaximaViaFile(InputString, WT, False)
    If Left$(s, 5) = "Fejln" Then
        WT = InputBox(TT.A(892) & WT & "s." & TT.A(893), "Stopped", 20)
        If Trim$(WT) <> "" Then
            WT = Trim$(WT)
            GoTo start
        Else
            StopNow = True
        End If
    End If
    If Assume = vbNullString Then
        pos = InStr(s, "positive, negative or zero?")
        pos2 = InStr(s, "positive or negative?")
    
        If pos > 0 Or pos2 > 0 Then
            Dim UF As New UserFormAskSign
            If pos <= 0 And pos2 > 0 Then pos = pos2
            pos2 = InStr(s, "Is ")
            Var = Mid$(s, pos2 + 3, pos - pos2 - 4)
            UF.Label_udtryk.Caption = Var
            UF.Show
            If UF.OptionButton_positiv.Value = True Then
                Assume = "assume(" & Var & ">0)$"
            ElseIf UF.OptionButton_negativ.Value = True Then
                Assume = "assume(" & Var & "<0)$"
            Else
                Assume = "assume(" & Var & "=0)$"
            End If
            Unload UF
            WT = 2
            GoTo start
            '        p = InStrRev(s, vbCr)
            '        p = InStrRev(s, vbCr, p - 2)
            '        If p > 0 Then
            '            p2 = InStr(p + 4, s, " ")
            '            If p2 > 0 Then
            '                Var = mid$(s, p + 3, p2 - p - 3)
            '            Else
            '                Var = "x"
            '            End If
            '        Else
            '            Var = "x"
            '        End If
            '        MsgBox "The problem could not be solved, because the sign of a variable is unknown. You can probably solve the problem by inserting a definition. Example:" & vbCr & " 'Define: " & Var & ">0'" & vbCrLf & "Maxima Asked:" & vbCrLf & right$(s, Len(s) - p)
        ElseIf InStr(s, " integer?") And MaximaComplex = False Then
            Assume = "assume(" & Var & ", integer)$"
            GoTo start
        End If
    End If
    '    i = 0
    '    Do Until Dir(FN) > "" Or i > 80
    '        i = i + 1
    '        Wait (0.1)
    '    Loop
    
    ShowDebug False, InputString, s
    
    FindLastMaximaOutput ConvertWordSymbols, s
    '    GenerateOutputArray (s)
    '    Unload UfWait2
#End If
End Sub


Sub InsertKillDef()
    If Len(KillDef) > 0 Then
        KillDef = Left$(KillDef, Len(KillDef) - 1)    'remove last comma
        MaximaInputStreng = MaximaInputStreng & "errcatch(kill(" & KillDef & "))$"
        KillDef = ""
    End If
    If Len(ForgetList) > 0 Then
        ForgetList = Left$(ForgetList, Len(ForgetList) - 1)    'remove last comma
        MaximaInputStreng = MaximaInputStreng & "forget(" & ForgetList & ")$reset(linenum,%,features,%rnum)$"
        ForgetList = ""
    End If

End Sub
Sub AddStandardSecondary()
' typically this should be run after the maxima line is set up to run the command
' this then ensures that the output comes as specified in the settings

    If MaximaDecOutType = 3 Then     ' scientific notation
        MaximaInputStrengSec = MaximaInputStrengSec & "scinotall(%),numer"    ' without ,numer the expression sqrt(2) will not be reduced
    ElseIf MaximaExact = 2 Or MaximaUnits Then    ' numerical
        MaximaInputStrengSec = MaximaInputStrengSec & "dectalall(%),numer"
    Else
        If MaximaExact = 0 Then    ' only exact if the expression is not too long
            MaximaInputStrengSec = MaximaInputStrengSec & "if length(charlist(string(%)))>70 then ev(dectalall(%),numer) else dectalall(%)"
        Else    ' eksakt
            MaximaInputStrengSec = MaximaInputStrengSec & "if length(charlist(string(%)))>150 then ev(dectalall(%),numer) else dectalall(%)"
        End If
    End If

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec & ";"
End Sub
Public Sub ExecuteMaximaCommand(Optional MaxCmd As String, Optional TranslateToMaxima As Boolean = False)
' executes the command that is in the command string or specified as a parameter
    Dim text As String
    If MaxCmd <> "" Then
        Kommando = MaxCmd    ' LCase maybe
    End If

    Application.ScreenUpdating = False
    If TranslateToMaxima Then
        MaximaInputStreng = MaximaInputStreng & CodeForMaxima(Kommando) & ";"
    Else
        text = Kommando
        text = Replace(text, ChrW$(9633), "")    ' characters before parentheses that make them invisible. Possibly problematic as the parentheses are not removed. also makes fractions small
        text = Replace(text, ChrW$(9508), "")    ' blank space character at least [ ]
        text = Replace(text, ChrW$(11), "")    ' shift-enter
        text = Replace(text, vbLf, "")    ' shift-enter and enter
        text = Replace(text, vbCrLf, "")
        text = Replace(text, vbCr, "")
        text = Replace(text, ChrW$(183), "*")    ' dot is replaced with times
        text = Replace(text, ChrW$(8901), "*")    ' \cdot , don't know why there seem to be two
        text = Replace(text, ChrW$(8729), "*")    '  \cdot replaced with times
        text = Replace(text, ChrW$(8226), "*")    ' thick dot replaced with times
        text = Replace(text, ChrW$(247), "/")    '
        text = Replace(text, ChrW$(8800), "#")    ' not equal to
        text = Replace(text, ChrW$(8804), "<=")    '
        text = Replace(text, ChrW$(8805), ">=")    '
        text = Replace(text, ChrW$(160), "")    ' before accent e.g. vector arrow 160 8407
        text = Replace(text, ChrW$(8289), "") ' function symbol
        text = Replace(text, ChrW$(9618), "") ' function symbol
        text = Replace(text, ChrW$(12310), "(") ' hidden brackets
        text = Replace(text, ChrW$(12311), ")") ' hidden brackets
        
        Kommando = text
        Kommando = Replace(Kommando, ChrW$(8289), "")
        MaximaInputStreng = MaximaInputStreng & Kommando & ";"
    End If
    Call RunMaxima(200, True, False)
    If StopNow Then GoTo slut
    
    Dim nr As Integer, i As Integer
    nr = UBound(Minput)

    If defindex > 0 And nr <= 1 + defindex Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf ' changed from moutput(1) in v.1.29 as errors were not included
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(defindex - nr) Is Nothing Then MathOArr(defindex - nr).Range.Select
        ElseIf defindex - nr - 1 >= 0 Then
            If Not MathOArr(defindex - nr - 1) Is Nothing Then MathOArr(defindex - nr - 1).Range.Select
        End If
        DefFejl = True
'        ActiveWindow.ScrollIntoView MathOArr(defindex - nr - 1).Range, True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    End If
slut:
End Sub

Sub MaximaSolve(variabel As String)
    Dim nr As Integer
    Dim i As Integer
    Dim gemMaxExct As Integer
    gemMaxExct = -1
    'ti = Timer

    If variabel = "" Then Exit Sub
    Dim rkommando As String, ikommando As String, ckommando As String

    ikommando = Kommando    ' command is saved. This is not tampered with
start:
    KommandoForklaring = "Solve(" & Kommando & "," & variabel & ")"    ' only for output of command explanation

    ckommando = CodeForMaxima(ikommando)    ' coded command
'    variabel = ReplaceFakeVarNames(variabel)
    '    kommando = "CSolve(" & ckommando & "," & variabel & ")"
    If TempDefs <> "" Then
        '        MaximaInputStreng = MaximaInputStreng & "[" & CodeForMaxima(Replace(TempDefs, "=", ":")) & "]$"
        '        AddToKillDefs (TempDefs)

        MaximaInputStreng = AddDefinition("definer:" & TempDefs) & MaximaInputStreng
    End If
    If AllTrig Then
        MaximaInputStreng = MaximaInputStreng & "AllTrig:true$"
    Else
        MaximaInputStreng = MaximaInputStreng & "AllTrig:false$"
    End If

    If MaximaExact = 2 Then
        rkommando = "ev(%,numer)"    ' attempt to solve: 6^2=5^2+3^2-2*5*3*cos(C) solve for C with number, gives partially exact
    Else
        rkommando = "ev(%)"    ' command where log and sin are converted
    End If
    '    rkommando = "ev(%th(2))" ' command where log and sin are converted to th2 due to slutprut when using answerquestion


    '    If Not Radians And Not MaximaExact = 2 Then ' can't remember why not num, but it's a problem when expressions with variables come in
    If Not Radians Then
        rkommando = "ConvertToDegr(" & rkommando & ")"
    End If
    '    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
    If MaximaLogOutput = 0 Then
        If (InStr(ikommando, "log" & ChrW$(8289)) > 0 Or InStr(ikommando, "10" & ChrW$(12311) & "^") > 0) Then
            rkommando = "ConvertLog(" & rkommando & ")"
        End If
        '    ElseIf MaximaLogOutput = 2 And Not MaximaExact = 2 Then
    ElseIf MaximaLogOutput = 2 Then
        rkommando = "ConvertLog(" & rkommando & ")"
    End If

    ' inputstring beeing build **************

    ' first a simple reduction of the equation. probably not necessary. Causes problems with questions as the command with question is the last one executed
    '    If MaximaExact = 2 Then
    '        MaximaInputStreng = MaximaInputStreng & "fullratsimp(" & ckommando & "),numer;"
    '    Else
    '        MaximaInputStreng = MaximaInputStreng & ckommando & ";"
    '    End If
    
    If MaximaUnits Then 'First reduce the units of the expression before the command. Can sometimes give a more complicated expression.
        If Not (InStr(ckommando, "diff") > 0 And InStr(ckommando, "subst") > 0) Then ' if any f' is included, it must not be reduced before
            ckommando = Replace(ConvertMaximaUnits(ckommando), """", "")
        End If
    End If
    
    If MaximaExact = 2 Then    ' necessary so that trig equations with sing are not solved exactly
        MaximaInputStreng = MaximaInputStreng & "Solve(ev(" & ckommando & ",numer)," & variabel & ");"
    Else
        MaximaInputStreng = MaximaInputStreng & "Solve(" & ckommando & "," & variabel & ");"
    End If

    '    RunMaxima reset:=False
    '    MaximaInputStreng = ""

    MaximaInputStrengSec = MaximaInputStrengSec & "fullratsimp(" & rkommando & ");"    'so log and trig conversions

    AddStandardSecondary

    Call RunMaxima

    'MsgBox MaximaInputStreng
    '    MaxProc.AnswerQuestion (MaximaInputStreng)
    '    FindLastMaximaOutput

    If StopNow Then
        If MaximaExact < 2 Then
            If MsgBox(TT.A(839), vbRetryCancel, "Stop") = vbRetry Then
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
    
    If defindex > 0 And nr <= 1 + defindex Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(defindex - nr) Is Nothing Then MathOArr(defindex - nr).Range.Select
        ElseIf defindex - nr - 1 >= 0 Then
            If Not MathOArr(defindex - nr - 1) Is Nothing Then MathOArr(defindex - nr - 1).Range.Select
        End If
        DefFejl = True
'        ActiveWindow.ScrollIntoView MathOArr(defindex - nr - 1).Range, True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    Else
        KommentarOutput = Minput(1) & vbLf & moutput(1)
        If nr >= 2 Then KommentarOutput = KommentarOutput & vbLf & moutput(2)
        For i = 0 To nr - 1
            KommentarOutput = KommentarOutput & Minput(nr - i)
        Next
        KommentarOutput = Replace(KommentarOutput, vbCrLf, "")
    End If

    If MaximaOutput = "[]" Then
        If LmSet Then
            MaximaOutput = "L=" & ChrW$(8709)
        Else
            MaximaOutput = variabel & ChrW$(8712) & ChrW$(8709)
        End If
    ElseIf MaximaOutput = "all" Or MaximaOutput = "[all]" Or MaximaOutput = "[" & variabel & "=k_1]" Then
        If LmSet Then
            MaximaOutput = "L=" & ChrW$(8477)
        Else
            MaximaOutput = variabel & ChrW$(8712) & ChrW$(8477)
        End If
    ElseIf MaximaOutput = "false" Then
        MaximaOutput = ""
    Else
        If InStr(omax.KommentarOutput, "solving system of equations") > 0 Then
            ConvertOutputToSolved2 (variabel)
        ElseIf LmSet Then
            ' notation with L={}

            If MaximaOutput <> "" Then
                MaximaOutput = Replace(MaximaOutput, "[", "")
                MaximaOutput = Replace(MaximaOutput, "]", "")
                MaximaOutput = Replace(MaximaOutput, "=", "")
                MaximaOutput = Replace(MaximaOutput, variabel, "")
                MaximaOutput = "L={" & MaximaOutput & "}"
            End If

        Else    ' notation with x=...
            ConvertOutputToSolved2
            If InStr(MaximaOutput, ";") Then ' if normal equation
                MaximaOutput = RemoveBrackets(MaximaOutput)
                MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & ChrW$(8744) & "    ")    ' Comma replaced by or
            ElseIf ListSeparator = "," And InStr(MaximaOutput, ",") Then
                MaximaOutput = RemoveBrackets(MaximaOutput)
                MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & ChrW$(8744) & "    ")    ' Comma replaced by or
            End If
        End If

        '    MaximaOutput = Replace(MaximaOutput, "%i", "i")
    End If    ' end of L=empty
    ResetMaxima

    If gemMaxExct >= 0 Then MaximaExact = gemMaxExct
slut:
End Sub
Sub MaximaSolveInequality(variabel As String)
    Dim nr As Integer
    Dim arr As Variant
    Dim i As Integer
    If variabel = "" Then Exit Sub

    Kommando = "SolveIneq(" & CodeForMaxima(Kommando) & "," & variabel & ")"
    MaximaInputStreng = MaximaInputStreng & "fullratsimp(" & Kommando & "),numer:false;%,numer;"

    AddStandardSecondary

    Call RunMaxima
    If StopNow Then GoTo slut

    nr = UBound(Minput)
    
    If defindex > 0 And nr <= 1 + defindex Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(defindex - nr) Is Nothing Then MathOArr(defindex - nr).Range.Select
        ElseIf defindex - nr - 1 >= 0 Then
            If Not MathOArr(defindex - nr - 1) Is Nothing Then MathOArr(defindex - nr - 1).Range.Select
        End If
        DefFejl = True
'        ActiveWindow.ScrollIntoView MathOArr(defindex - nr - 1).Range, True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    Else
        If Not MaximaExact = 2 Then
            MaximaOutput = ConvertToWordSymbols(moutput(nr - 2))
        End If
        ConvertOutputToSolved2 (variabel)
    End If
    
slut:
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
        Kommentar = "There must be an equal sign in an equation."
        MaximaOutput = ""
        Exit Sub
    End If
    Kommando = Replace(Kommando, "=", "-(") & ")"    ' must be an expression not an equation. This is set equal to zero
    Kommando = "newton(" & CodeForMaxima(Kommando) & "," & variabel & "," & guess & "," & eps & ")"
    MaximaInputStreng = MaximaInputStreng & "[ratsimpexpons:false,load(""newton1"")]$" & Kommando & ";"
    ' for inexplicable reasons, exact solutions may emerge

    If MaximaDecOutType = 3 Then     ' scientific notation
        MaximaInputStrengSec = "scinotall(%),numer"  ' without ,numer the expression sqrt(2) will not be reduced, e.g.
    Else
        MaximaInputStrengSec = "dectalall(%),numer"
    End If

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec & ";"

    Call RunMaxima
    If StopNow Then GoTo slut

    Dim nr As Integer, i As Integer
    nr = UBound(Minput)

    If defindex > 0 And nr <= 1 + defindex Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf '
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(defindex - nr) Is Nothing Then MathOArr(defindex - nr).Range.Select
        ElseIf defindex - nr - 1 >= 0 Then
            If Not MathOArr(defindex - nr - 1) Is Nothing Then MathOArr(defindex - nr - 1).Range.Select
        End If
        DefFejl = True
'        ActiveWindow.ScrollIntoView MathOArr(defindex - nr - 1).Range, True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    Else
        If LmSet Then
            ' notation med L={}
            MaximaOutput = Trim$(MaximaOutput)
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
                MaximaOutput = variabel & ChrW$(&H2248) & RemoveBrackets(MaximaOutput)
            End If
            MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & ChrW$(8744) & "    " & variabel & ChrW$(&H2248))    ' Comma replaced by or
        End If
    End If
    
slut:
End Sub
Sub MaximaFindRoot(ByVal variabel As String, ByVal xmin As String, ByVal xmax As String, Optional ByVal eps As String)
    If variabel = "" Then Exit Sub
    If eps = "" Then
        eps = "10^-" & MaximaCifre
    End If
    Kommando = Replace(Kommando, "=", "-(") & ")"    ' must be an expression not an equation. This is set equal to zero
    Kommando = "find_root(" & CodeForMaxima(Kommando) & "," & variabel & "," & xmin & "," & xmax & ")"
    MaximaInputStreng = MaximaInputStreng & "[ratsimpexpons:false]$" & Kommando & ";"

    If MaximaDecOutType = 3 Then     ' scientific notation
        MaximaInputStrengSec = "scinotall(%),numer"  ' without ,numer the expression sqrt(2) will not be reduced, e.g.
    Else
        MaximaInputStrengSec = "dectalall(%),numer"
    End If
    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec & ";"

    Call RunMaxima
    If StopNow Then GoTo slut

    MaximaOutput = Trim$(MaximaOutput)
    If InStr(MaximaOutput, "find_root") > 0 Then
        MaximaOutput = "[]"
    End If
    If LmSet Then
        ' notation L={}
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
            MaximaOutput = variabel & ChrW$(&H2248) & RemoveBrackets(MaximaOutput)
        End If
        MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & ChrW$(8744) & "    " & variabel & ChrW$(&H2248))    ' Comma replaced by or
    End If

slut:
End Sub
Sub Nsolve(ByVal variabel As String, ByVal qn As Integer, ByVal qm As Integer, ByVal qtid As Integer, ByVal maxsol As Integer, Optional ByVal qnn As Integer = 30, Optional ByVal qmm As Integer = 30, Optional newt As Boolean = False)
    If variabel = "" Then Exit Sub
    '    eps = "10^-" & MaximaCifre
    If TempDefs <> "" Then
        MaximaInputStreng = AddDefinition("definer:" & TempDefs) & MaximaInputStreng
    End If

    Kommando = "nsolve(" & CodeForMaxima(Kommando) & "," & variabel & "," & qn & "," & qm & "," & qtid & "," & maxsol & "," & qnn & "," & qmm & "," & LCase$(newt) & ")"
    MaximaInputStreng = MaximaInputStreng & "[ratsimpexpons:false]$" & Kommando & ";"

    If MaximaDecOutType = 3 Then     ' scientific notation
        MaximaInputStrengSec = "scinotall(%),numer"  ' without ,numer the expression sqrt(2) will not be reduced, e.g.
    Else
        MaximaInputStrengSec = "dectalall(%),numer"
    End If
    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec & ";"

#If Mac Then
    RunMaxima 2
#Else
    RunMaxima -150
#End If
    If StopNow Then GoTo slut

    Dim nr As Integer, i As Integer
    nr = UBound(Minput)

    If defindex > 0 And nr <= 1 + defindex Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf '
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(defindex - nr) Is Nothing Then MathOArr(defindex - nr).Range.Select
        ElseIf defindex - nr - 1 >= 0 Then
            If Not MathOArr(defindex - nr - 1) Is Nothing Then MathOArr(defindex - nr - 1).Range.Select
        End If
        DefFejl = True
'        ActiveWindow.ScrollIntoView MathOArr(defindex - nr - 1).Range, True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    Else
        If LmSet Then
            ' notation med L={}
            MaximaOutput = Trim$(MaximaOutput)
            If MaximaOutput <> "" Then
                MaximaOutput = Replace(MaximaOutput, "[", "")
                MaximaOutput = Replace(MaximaOutput, "]", "")
                MaximaOutput = Replace(MaximaOutput, "=", "")
                MaximaOutput = Replace(MaximaOutput, variabel, "")
                MaximaOutput = "L={" & MaximaOutput & "}"
            End If
        Else    ' notation med x=...
            MaximaOutput = RemoveBrackets(MaximaOutput)
            MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & ChrW$(8744) & "    ")    ' Comma replaced by or
            MaximaOutput = Replace(MaximaOutput, "=", ChrW$(&H2248))    ' cirka equal
            '    If MaximaOutput <> "" Then
            '        MaximaOutput = variabel & chrw$(&H2248) & RemoveBrackets(MaximaOutput)
            '    End If
            '    MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & chrw$(8744) & "    " & variabel & chrw$(&H2248)) ' Comma replaced by or
        End If
    End If


slut:
End Sub

Sub SolveSystem(Variable As String)
    Dim Kommentar As String
    Dim i As Integer
    Dim rkommando As String, ikommando As String

    ikommando = Kommando    ' command is saved. This is not tampered with

    If Variable = "" Then Exit Sub
    If InStr(Kommandoer(0), "=") < 1 Then
        Kommentar = TT.A(139)
        MaximaOutput = ""
        Exit Sub
    End If
    If TempDefs <> "" Then
        MaximaInputStreng = AddDefinition("definer:" & TempDefs) & MaximaInputStreng
        '        MaximaInputStreng = MaximaInputStreng & "[" & CodeForMaxima(Replace(TempDefs, "=", ":")) & "]$"
        '        AddToKillDefs (TempDefs)
    End If

    rkommando = "ev(%)"    ' command where log and sin are converted
    '    If Not Radians And Not MaximaExact = 2 Then ' can't remember why not num, but it's a problem when expressions with variables come in
    If Not Radians Then
        rkommando = "ConvertToDegr(" & rkommando & ")"
    End If
    '    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
    If MaximaLogOutput = 0 Then
        If (InStr(ikommando, "log" & ChrW$(8289)) > 0 Or InStr(ikommando, "10" & ChrW$(12311) & "^") > 0) Then
            rkommando = "ConvertLog(" & rkommando & ")"
        End If
        '    ElseIf MaximaLogOutput = 2 And Not MaximaExact = 2 Then
    ElseIf MaximaLogOutput = 2 Then
        rkommando = "ConvertLog(" & rkommando & ")"
    End If


'    MaximaInputStreng = MaximaInputStreng & "autonsolve:false$"
    Kommando = "[" & CodeForMaxima(Kommandoer(0))
    For i = 1 To UBound(Kommandoer)
        Kommando = Kommando & "," & CodeForMaxima(Kommandoer(i))
    Next

    ' not sure this is good for systems of equations
'    If MaximaUnits Then 'First reduce the units of the expression before the command. Can sometimes give a more complicated expression.
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

    If MaximaExact = 2 Then    ' necessary so that trig equations with sing are not solved exactly
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
'    If MaximaExact = 1 Then    ' exact
        MaximaInputStreng = MaximaInputStreng & ";"
'    Else
'        MaximaInputStreng = MaximaInputStreng & ",autonsolve=true;"
'    End If

    MaximaInputStrengSec = MaximaInputStrengSec & "fullratsimp(" & rkommando & ");"    'then log and trig convertions

    AddStandardSecondary

    Call RunMaxima
    If StopNow Then GoTo slut

    Dim nr As Integer
    nr = UBound(Minput)

    If defindex > 0 And nr <= 1 + defindex Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf '
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(defindex - nr) Is Nothing Then MathOArr(defindex - nr).Range.Select
        ElseIf defindex - nr - 1 >= 0 Then
            If Not MathOArr(defindex - nr - 1) Is Nothing Then MathOArr(defindex - nr - 1).Range.Select
        End If
        DefFejl = True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    Else
        KommentarOutput = moutput(1)
        ConvertOutputToSolved2 (Variable)
    End If
slut:
End Sub
Sub Eliminate(Variable As String)
    Dim Kommentar As String
    Dim i As Integer
    Dim arr As Variant
    If Variable = "" Then Exit Sub
    If InStr(Kommandoer(0), "=") < 1 Then
        Kommentar = TT.A(139)
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

    Kommando = Kommando & "]"
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"
    Variable = "[" & Variable & "]"
    Kommando = "eliminate(%," & Variable & ")"
    If Not Radians Then
        Kommando = "ratsimp(ConvertToDegr(" & Kommando & "))"
    End If

    MaximaInputStreng = MaximaInputStreng & Kommando & ";"

    AddStandardSecondary

    Call RunMaxima
    If StopNow Then GoTo slut

    Dim nr As Integer
    nr = UBound(Minput)

    If defindex > 0 And nr <= 1 + defindex Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf '
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(defindex - nr) Is Nothing Then MathOArr(defindex - nr).Range.Select
        ElseIf defindex - nr - 1 >= 0 Then
            If Not MathOArr(defindex - nr - 1) Is Nothing Then MathOArr(defindex - nr - 1).Range.Select
        End If
        DefFejl = True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    Else
        MaximaOutput = Mid$(MaximaOutput, 2, Len(MaximaOutput) - 2)
        arr = Split(MaximaOutput, ListSeparator)
        If Not (InStr(arr(0), "=") > 0) Then
            MaximaOutput = arr(0) & "=0"
        Else
            MaximaOutput = arr(0)
        End If
        For i = 1 To UBound(arr)
            If Not (InStr(arr(i), "=") > 0) Then
                MaximaOutput = MaximaOutput & "    " & ChrW$(8743) & "    " & arr(i) & "=0"
            Else
                MaximaOutput = MaximaOutput & "    " & ChrW$(8743) & "    " & arr(i)
            End If
        Next
    End If
    
slut:
End Sub
Sub SolveSystemNumeric(ByVal Variable As String, ByVal guess As String, Optional ByVal eps As String)
    Dim Kommentar As String
    Dim i As Integer
    If Variable = "" Then Exit Sub
    If InStr(Kommandoer(0), "=") < 1 Then
        Kommentar = TT.A(139)
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

    If MaximaDecOutType = 3 Then     ' scientific notation
        MaximaInputStrengSec = "scinotall(%),numer;"
    Else
        MaximaInputStrengSec = "dectalall(%),numer;"
    End If

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec


    Call RunMaxima
    If StopNow Then GoTo slut

    Dim nr As Integer
    nr = UBound(Minput)

    If defindex > 0 And nr <= 1 + defindex Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf '
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(defindex - nr) Is Nothing Then MathOArr(defindex - nr).Range.Select
        ElseIf defindex - nr - 1 >= 0 Then
            If Not MathOArr(defindex - nr - 1) Is Nothing Then MathOArr(defindex - nr - 1).Range.Select
        End If
        DefFejl = True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    Else
        ConvertOutputToSolved2 (Variable)
    End If

slut:
End Sub
Sub beregn()
    Dim inkom As String, pkommando As String, StartKommando3 As String
    Dim posca As Integer, posligmed As Integer, possumtegn As Integer
    Dim gemMaxExct As Integer
    gemMaxExct = -1
    Kommando = Replace(Kommando, vbLf, "")
    Kommando = Replace(Kommando, vbCrLf, "")
    Kommando = Replace(Kommando, vbCr, "")
    Kommando = Trim$(Kommando)

    If Right$(Kommando, 1) = "=" Then
        Kommando = Left$(Kommando, Len(Kommando) - 1)
    End If

start:
    StartKommando = Kommando    ' command before it is tampered with
    StartKommando3 = Kommando
    
    pkommando = Kommando    ' command before it is tampered with
    pkommando = Replace(pkommando, ChrW$(12310), "") ' special invisible brackets are removed
    pkommando = Replace(pkommando, ChrW$(12311), "")
    Kommando = CodeForMaxima(Kommando)
    If Kommando = vbNullString And ConvertErrorText <> vbNullString Then
        GoTo slut
    End If
    inkom = "%"    ' save command

    If MaximaLogOutput = 0 Then
        If (InStr(pkommando, "log" & ChrW$(8289)) > 0 Or InStr(pkommando, "10" & ChrW$(12311) & "^") > 0) Then
            inkom = "ConvertLog(" & inkom & ")"
        End If
    ElseIf MaximaLogOutput = 2 Then
        inkom = "ConvertLog(" & inkom & ")"
    End If

    If Not Radians Then
        inkom = "ConvertToDegr(" & inkom & ")"
    End If

    If MaximaComplex And PolarOutput Then
        Kommando = "ConvertToPolarAngleNotation(" & Kommando
        If Radians Then
            Kommando = Kommando & ",true)"
        Else
            Kommando = Kommando & ",false)"
        End If
    End If

    If MaximaDecOutType = 3 Then    ' scientific notation and maybe units
        MaximaInputStreng = MaximaInputStreng & Kommando & ";"
        MaximaInputStrengSec = "ev(" & inkom & ",numer);scinotall(%);ev(is(equal(ev(scanmap(strtonum,%th(1)),nouns),%th(2))),numer);"
        MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec
    ElseIf MaximaExact = 2 Or MaximaUnits Then    ' units overrides auto,exact,num
        If MaximaBigFloat Then
            MaximaInputStreng = MaximaInputStreng & "fpprec:50$convertbfloat(" & Kommando & ");"
        Else
            MaximaInputStreng = MaximaInputStreng & Kommando & ",numer;"
        End If
        MaximaInputStrengSec = "simplifynum(" & inkom & ");"    ' dectal is needed for units and correct number of bet cif. . try with ,numer
#If Mac Then
        If MaximaUnits Then
            '            MaximaInputStrengSec = "applyb1(%,unitrule),numer;" 'does not improve
            If InStr(ConvertUnits(OutUnits), "%") > 0 Then ' if % in outunits it may give unit placed before number
                MaximaInputStrengSec = MaximaInputStrengSec & "%,numer;" ' dectalall places units with %m prefix before numbers, so rather wrong number of sig fig?
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
        MaximaInputStrengSec = MaximaInputStrengSec & "ev(is(equal(ev(scanmap(strtonum,%th(1)),nouns),%th(2))),numer);"     ' dectal is needed for units and correct number of bet cif. . try with ,numer
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
        MaximaInputStrengSec = inkom & ";simplifynum(%),numer;dectalall(" & inkom & "),numer;is(equal(%th(2),%th(3)))"    ' different ending then or can come back again
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
            If MsgBox(TT.A(839), vbRetryCancel, "Stop") = vbRetry Then
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

    Dim nr As Integer, i As Integer, DefNo As Integer
    Dim eql As String
    nr = UBound(Minput)
    '    DefNo = UBound(ldefname) + 1
    DefNo = defindex

    If DefNo > 0 And nr <= 1 + DefNo Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf '
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(DefNo - nr) Is Nothing Then MathOArr(DefNo - nr).Range.Select
        ElseIf DefNo - nr - 1 >= 0 Then
            If Not MathOArr(DefNo - nr - 1) Is Nothing Then MathOArr(DefNo - nr - 1).Range.Select
        End If
        DefFejl = True
        '        ActiveWindow.ScrollIntoView MathOArr(defindex - nr - 1).Range, True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    ElseIf MaximaExact = 0 And MaximaDecOutType < 3 And Not MaximaUnits Then
        Dim exactres As String
        Dim numres As String
        If nr > 1 Then moutput(nr - 1) = Replace(moutput(nr - 1), vbLf, "")
        If nr > 2 Then moutput(nr - 2) = Replace(moutput(nr - 2), vbLf, "")
        numres = ConvertToWordSymbols(moutput(nr - 1))
        If UBound(moutput) > 3 Then exactres = ConvertToWordSymbols(moutput(nr - 3))
        If exactres = "" And numres = "" Then exactres = MaximaOutput    ' if there have been any questions
        eql = MaximaOutput
        If eql = "true" Or StartKommando3 = numres Then
            eql = "="
        Else
            eql = ChrW$(&H2248)
        End If
        MaximaOutput = ""
        pkommando = ConvertToWordSymbols(Replace(pkommando, DecSeparator, ".")) ' pcommand cannot be compared 100% with output until it has also been run through
        If exactres <> "" Then
            '    MsgBox "pkommando:" & pkommando & vbCrLf & " exactres:" & ConvertToWordSymbols(exactres)
            If exactres = numres Or pkommando = exactres Or StartKommando3 = exactres Then ' if numeric and exactly equal or exactly equal input
                MaximaOutput = MaximaOutput & eql & numres
            Else
                MaximaOutput = MaximaOutput & "=" & exactres & eql & numres
            End If
            
            KommentarOutput = ""
            For i = 0 To nr - 1
                KommentarOutput = KommentarOutput & Minput(nr - i)
            Next
            KommentarOutput = Replace(KommentarOutput, vbCrLf, "")
        Else
            MaximaOutput = ""
        End If
    Else    ' vid not, num, exact
        Dim eqsg As String
        If MaximaExact <> 1 Then
            eql = MaximaOutput
            nr = nr - 1 ' if auto or num, then last output is just true/false
        End If
        If MaximaExact = 1 Then ' exact
            eqsg = "="
        ElseIf eql = "true" Then
            eqsg = "="
        Else
            eqsg = ChrW$(&H2248)
        End If
        If moutput(nr) <> vbNullString Then
            MaximaOutput = eqsg & ConvertToWordSymbols(moutput(nr))
            KommentarOutput = ""
            If MaximaOutput = "=fejl" Then
                KommentarOutput = moutput(nr - 2) & moutput(nr - 1)
            End If
            For i = 0 To nr - 1
                KommentarOutput = KommentarOutput & Minput(nr - i)
            Next
            KommentarOutput = Replace(KommentarOutput, vbCrLf, "")
        Else
            MaximaOutput = ""
        End If
        'GoTo slut
    End If
    
    If ListSeparator = "," Then MaximaOutput = Replace(MaximaOutput, ",", ", ") ' primarily necessary for lists
        
    If gemMaxExct >= 0 Then MaximaExact = gemMaxExct
slut:
End Sub
Sub ConvertOutputToSolved()
' converts output from [1,3] to x=1 v x=3 or L={1,3}
' intended for numerical solution only
    Dim Variable As String
    Dim arr As Variant
    Dim i As Integer
    Dim Sep As String
    If LmSet Then
        ' notation med L={}
        MaximaOutput = Trim$(MaximaOutput)
        If MaximaOutput <> "" Then
            MaximaOutput = Mid$(MaximaOutput, 2, Len(MaximaOutput) - 2)
            MaximaOutput = Replace(MaximaOutput, "[", "(")
            MaximaOutput = Replace(MaximaOutput, "]", ")")
            MaximaOutput = Replace(MaximaOutput, "=", "")
            arr = Split(Mid$(Variable, 2, Len(Variable) - 2), ",")
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
        MaximaOutput = Replace(MaximaOutput, Sep, "    " & ChrW$(8743) & "    ")
        o = ea.GetNextBracketContent
        If o = "" Then Exit Sub    ' only 1 solution

        MaximaOutput = "(" & MaximaOutput & ")"
        o = "(" & Replace(o, Sep, "    " & ChrW$(8743) & "    ") & ")"
        MaximaOutput = MaximaOutput & "   " & ChrW$(8744) & "   " & o

        Do
            o = ea.GetNextBracketContent
            If o = "" Then Exit Do
            o = "(" & Replace(o, Sep, "    " & ChrW$(8743) & "    ") & ")"
            MaximaOutput = MaximaOutput & "   " & ChrW$(8744) & "   " & o
        Loop
    End If

End Sub
Sub ConvertOutputToSolvedGGB()
' for geogebra output for systems of equations
    Dim p As Integer, s As String, i As Integer, j As Integer
    Dim ea As New ExpressionAnalyser
    Dim arr() As String, Arr2() As String
    
    If MaximaOutput = "?" Or MaximaOutput = "" Then
        Exit Sub
    End If
    ea.SetNormalBrackets
    ea.text = MaximaOutput
'    p = InStr(ea.text, "(_")
    p = 1
    s = ea.GetNextBracketContent(p + 1)
    arr = Split(s, "@")
    MaximaOutput = ""
    
    For i = 0 To UBound(arr)
        Arr2 = Split(arr(i), "&")
        If UBound(Arr2) > 0 Then
            If UBound(arr) > 0 Then MaximaOutput = MaximaOutput & "("
            For j = 0 To UBound(Arr2)
                MaximaOutput = MaximaOutput & Arr2(j) & "    " & ChrW$(8743) & "    "
            Next
            MaximaOutput = Left$(MaximaOutput, Len(MaximaOutput) - 9)
            If UBound(arr) > 0 Then MaximaOutput = MaximaOutput & ")"
        End If
        MaximaOutput = MaximaOutput & "    " & ChrW$(8744) & "    "
    Next
    If Len(MaximaOutput) > 9 Then MaximaOutput = Left$(MaximaOutput, Len(MaximaOutput) - 9)
    
End Sub
Sub ConvertOutputToSolved2(Optional Variable As String)
' converts output from [[x=1,y=3],[x=3,y=5]] to (x=1 and x=3) v ... or L={(1,3),(3,5)}
' intended for symbolic solution and mnewton and SolveIneq only
    Dim arr As Variant
    Dim i As Integer

    If MaximaOutput = "[]" Then
        If LmSet Then
            MaximaOutput = "L=" & ChrW$(8709)
        Else
            MaximaOutput = Variable & ChrW$(8712) & ChrW$(8709)
        End If
    ElseIf MaximaOutput = "all" Or MaximaOutput = "[[all]]" Or MaximaOutput = "[" & Variable & "=k_1]" Then
        If LmSet Then
            MaximaOutput = "L=" & ChrW$(8477)
        Else
            MaximaOutput = Variable & ChrW$(8712) & ChrW$(8477)
        End If
    ElseIf MaximaOutput = "false" Then
        MaximaOutput = ""
    Else

        If LmSet Then
            ' notation L={}
            MaximaOutput = Trim$(MaximaOutput)
            If MaximaOutput <> "" Then
                MaximaOutput = Mid$(MaximaOutput, 2, Len(MaximaOutput) - 2)
                MaximaOutput = Replace(MaximaOutput, "[", "(")
                MaximaOutput = Replace(MaximaOutput, "]", ")")
                MaximaOutput = Replace(MaximaOutput, "=", "")
                arr = Split(Mid$(Variable, 2, Len(Variable) - 2), ",")
                For i = 0 To UBound(arr)
                    MaximaOutput = Replace(MaximaOutput, arr(i), "")
                Next
                MaximaOutput = "L={" & MaximaOutput & "}"
            End If

        Else    ' notation x=...

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
            MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & ChrW$(8743) & "    ")

            o = ea.GetNextBracketContent
            If o = "" Then Exit Sub    ' only 1 solution

            If InStr(MaximaOutput, ChrW$(8743)) > 0 Then    ' only parentheses implication sign
                MaximaOutput = "(" & MaximaOutput & ")"
            End If

            o = Replace(o, ListSeparator, "    " & ChrW$(8743) & "    ")
            If InStr(o, ChrW$(8743)) > 0 Then    ' only parentheses implication sign
                o = "(" & o & ")"
            End If

            MaximaOutput = MaximaOutput & "   " & ChrW$(8744) & "   " & o

            Do
                o = ea.GetNextBracketContent
                If o = "" Then Exit Do
                o = Replace(o, ListSeparator, "    " & ChrW$(8743) & "    ")
                If InStr(o, ChrW$(8743)) > 0 Then    ' only parentheses implication sign
                    o = "(" & o & ")"
                End If
                MaximaOutput = MaximaOutput & "   " & ChrW$(8744) & "   " & o
            Loop

        End If
    End If

End Sub
Sub Omskriv(vidnot As Boolean, autosimp As Boolean, faktoriser As Boolean, udvid As Boolean, rationaliser As Boolean, trigsimp As Boolean, CompleteSquare As Boolean)
    Dim pkommando As String
    pkommando = Kommando

    Kommando = CodeForMaxima(Kommando)

    If TempDefs <> "" Then
        MaximaInputStreng = AddDefinition("definer:" & TempDefs) & MaximaInputStreng
    End If

    If MaximaUnits Then 'First reduce the units of the expression before the command. Can sometimes give a more complicated expression.
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
    If CompleteSquare Then
        Kommando = "CompleteSquare(" & Kommando & ")"
    End If

    If Not Radians Then
        '        kommando = "simplify(ConvertToDegr(" & kommando & "))"
        Kommando = "ConvertToDegr(" & Kommando & ")"
    End If

    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
        If (InStr(pkommando, "log" & ChrW$(8289)) > 0 Or InStr(pkommando, "10" & ChrW$(12311) & "^") > 0) Then
            Kommando = "ConvertLog(" & Kommando & ")"
        End If
    ElseIf MaximaLogOutput = 2 And Not MaximaExact = 2 Then
        Kommando = "ConvertLog(" & Kommando & ")"
    End If

    If MaximaExact = 2 Then
        Kommando = Kommando & ",numer"
    End If
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"


    If MaximaDecOutType = 3 Then   ' scientific notation
        MaximaInputStrengSec = "scinotall(%),numer;"
    ElseIf MaximaExact = 2 Then
        MaximaInputStrengSec = "dectalall(%),numer;"
    Else
        MaximaInputStrengSec = ""
    End If

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec

    Call RunMaxima
    If StopNow Then GoTo slut

    Dim nr As Integer, i As Integer
    nr = UBound(Minput)

    If defindex > 0 And nr <= 1 + defindex Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf '
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(defindex - nr) Is Nothing Then MathOArr(defindex - nr).Range.Select
        ElseIf defindex - nr - 1 >= 0 Then
            If Not MathOArr(defindex - nr - 1) Is Nothing Then MathOArr(defindex - nr - 1).Range.Select
        End If
        DefFejl = True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    End If
slut:
End Sub
Sub Reduce()
    Dim pkommando As String
    pkommando = Kommando

    Kommando = CodeForMaxima(Kommando)
    If MaximaUnits Then 'First reduce the units of the expression before the command. Can sometimes give a more complicated expression.
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
        If (InStr(pkommando, "log" & ChrW$(8289)) > 0 Or InStr(pkommando, "10" & ChrW$(12311) & "^") > 0) Then
            Kommando = "ConvertLog(" & Kommando & ")"
        End If
    ElseIf MaximaLogOutput = 2 And Not MaximaExact = 2 Then
        Kommando = "ConvertLog(" & Kommando & ")"
    End If

    If MaximaExact = 2 Then
        Kommando = Kommando & ",numer"
    End If
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"

    If MaximaDecOutType = 3 Then     ' scientific notation
        MaximaInputStrengSec = "scinotall(%),numer;"
    ElseIf MaximaExact = 2 Then
        MaximaInputStrengSec = "dectalall(%),numer;"
    Else
        MaximaInputStrengSec = ""
    End If

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec

    Call RunMaxima
    If StopNow Then GoTo slut

    Dim nr As Integer, i As Integer
    nr = UBound(Minput)

    If defindex > 0 And nr <= 1 + defindex Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf '
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(defindex - nr) Is Nothing Then MathOArr(defindex - nr).Range.Select
        ElseIf defindex - nr - 1 >= 0 Then
            If Not MathOArr(defindex - nr - 1) Is Nothing Then MathOArr(defindex - nr - 1).Range.Select
        End If
        DefFejl = True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    End If
slut:
End Sub
Sub CompareTest()
    Dim pkommando As String
    pkommando = Kommando
    Kommando = CodeForMaxima(Kommando)
    Kommando = "TestTF(" & Kommando & ")"
    If Not Radians Then
        Kommando = "ConvertToDegr(" & Kommando & ")"    ' It might be problematic that ratsimp cannot be called when expanding again. converttodegree then doesn't work?
    End If
    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
        If (InStr(pkommando, "log" & ChrW$(8289)) > 0 Or InStr(pkommando, "10" & ChrW$(12311) & "^") > 0) Then
            Kommando = "ConvertLog(" & Kommando & ")"
        End If
    ElseIf MaximaLogOutput = 2 And Not MaximaExact = 2 Then
        Kommando = "ConvertLog(" & Kommando & ")"
    End If
    If MaximaExact = 2 Then
        Kommando = Kommando & ",numer"
    End If
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"

    If MaximaDecOutType = 3 Then     ' scientific notation
        MaximaInputStrengSec = "scinotall(%),numer;"
    ElseIf MaximaExact = 2 Then
        MaximaInputStrengSec = "dectalall(%),numer;"
    Else
        MaximaInputStrengSec = ""
    End If

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec

    Call RunMaxima
    If StopNow Then GoTo slut

    Dim nr As Integer, i As Integer
    nr = UBound(Minput)

    If defindex > 0 And nr <= 1 + defindex Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf '
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(defindex - nr) Is Nothing Then MathOArr(defindex - nr).Range.Select
        ElseIf defindex - nr - 1 >= 0 Then
            If Not MathOArr(defindex - nr - 1) Is Nothing Then MathOArr(defindex - nr - 1).Range.Select
        End If
        DefFejl = True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    End If

slut:
End Sub
Sub Factor()
    Dim pkommando As String
    pkommando = Kommando
    Kommando = CodeForMaxima(Kommando)
    Kommando = "factor(" & Kommando & ")"
    If Not Radians Then
        Kommando = "ConvertToDegr(" & Kommando & ")"    ' It might be problematic that ratsimp cannot be called when expanding again. converttodegree then doesn't work?
    End If
    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
        If (InStr(pkommando, "log" & ChrW$(8289)) > 0 Or InStr(pkommando, "10" & ChrW$(12311) & "^") > 0) Then
            Kommando = "ConvertLog(" & Kommando & ")"
        End If
    ElseIf MaximaLogOutput = 2 And Not MaximaExact = 2 Then
        Kommando = "ConvertLog(" & Kommando & ")"
    End If
    If MaximaExact = 2 Then
        Kommando = Kommando & ",numer"
    End If
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"

    If MaximaDecOutType = 3 Then     ' scientific notation
        MaximaInputStrengSec = "scinotall(%),numer;"
    ElseIf MaximaExact = 2 Then
        MaximaInputStrengSec = "dectalall(%),numer;"
    Else
        MaximaInputStrengSec = ""
    End If

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec

    Call RunMaxima
    If StopNow Then GoTo slut

    Dim nr As Integer, i As Integer
    nr = UBound(Minput)

    If defindex > 0 And nr <= 1 + defindex Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf '
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(defindex - nr) Is Nothing Then MathOArr(defindex - nr).Range.Select
        ElseIf defindex - nr - 1 >= 0 Then
            If Not MathOArr(defindex - nr - 1) Is Nothing Then MathOArr(defindex - nr - 1).Range.Select
        End If
        DefFejl = True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    End If

slut:
End Sub
Sub Expand()
    Dim pkommando As String
    pkommando = Kommando

    Kommando = CodeForMaxima(Kommando)
    Kommando = "expand(" & Kommando & ")"
    If Not Radians And Not MaximaExact = 2 Then
        Kommando = "ConvertToDegr(" & Kommando & ")"
    End If
    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
        If (InStr(pkommando, "log" & ChrW$(8289)) > 0 Or InStr(pkommando, "10" & ChrW$(12311) & "^") > 0) Then
            Kommando = "ConvertLog(" & Kommando & ")"
        End If
    ElseIf MaximaLogOutput = 2 And Not MaximaExact = 2 Then
        Kommando = "ConvertLog(" & Kommando & ")"
    End If

    If MaximaExact = 2 Then
        Kommando = Kommando & ",numer"
    End If
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"

    If MaximaDecOutType = 3 Then     ' scientific notation
        MaximaInputStrengSec = "scinotall(%),numer;"
    ElseIf MaximaExact = 2 Then
        MaximaInputStrengSec = "dectalall(%),numer;"
    Else
        MaximaInputStrengSec = ""
    End If

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec

    Call RunMaxima
    If StopNow Then GoTo slut

    Dim nr As Integer, i As Integer
    nr = UBound(Minput)

    If defindex > 0 And nr <= 1 + defindex Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf '
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(defindex - nr) Is Nothing Then MathOArr(defindex - nr).Range.Select
        ElseIf defindex - nr - 1 >= 0 Then
            If Not MathOArr(defindex - nr - 1) Is Nothing Then MathOArr(defindex - nr - 1).Range.Select
        End If
        DefFejl = True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    End If
slut:
End Sub
Sub Differentier(variabel As String)
    Dim pkommando As String
    pkommando = Kommando

    Kommando = CodeForMaxima(Kommando)

    Kommando = "diff(" & Kommando & "," & variabel & ",1)"
    If Not Radians Then
        Kommando = "ratsimp(ConvertToDegr(" & Kommando & "))"
    End If
    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
        If (InStr(pkommando, "log" & ChrW$(8289)) > 0 Or InStr(pkommando, "10" & ChrW$(12311) & "^") > 0) Then
            Kommando = "ConvertLog(" & Kommando & ")"
        End If
    ElseIf MaximaLogOutput = 2 And Not MaximaExact = 2 Then
        Kommando = "ConvertLog(" & Kommando & ")"
    End If

    Kommando = "dectalallNum(" & Kommando & ")"
    MaximaInputStreng = MaximaInputStreng & Kommando & ";"

    AddStandardSecondary

    Call RunMaxima
    
    Dim nr As Integer, i As Integer
    nr = UBound(Minput)

    If defindex > 0 And nr <= 1 + defindex Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf '
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(defindex - nr) Is Nothing Then MathOArr(defindex - nr).Range.Select
        ElseIf defindex - nr - 1 >= 0 Then
            If Not MathOArr(defindex - nr - 1) Is Nothing Then MathOArr(defindex - nr - 1).Range.Select
        End If
        DefFejl = True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    End If
    
slut:
End Sub
Sub UserRegression(fkt As String, Var As String, xlist As String, ylist As String, prec As Integer, VarGuess As String)
    Dim pkommando As String
    pkommando = fkt

    Kommando = "GeneralRegr(" & CodeForMaxima(fkt) & "," & Var & "," & xlist & "," & ylist & ",1e-" & prec & "," & VarGuess & ")"
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

    If MaximaDecOutType = 3 Then     ' scientific notation
        MaximaInputStrengSec = "scinotall(%),numer;"
    ElseIf MaximaExact = 2 Then
        MaximaInputStrengSec = "dectalall(%),numer;"
    Else
        MaximaInputStrengSec = ""
    End If

    MaximaInputStreng = MaximaInputStreng & MaximaInputStrengSec

    Call RunMaxima
    If StopNow Then GoTo slut

    Dim nr As Integer, i As Integer
    nr = UBound(Minput)

    If defindex > 0 And nr <= 1 + defindex Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf '
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(defindex - nr) Is Nothing Then MathOArr(defindex - nr).Range.Select
        ElseIf defindex - nr - 1 >= 0 Then
            If Not MathOArr(defindex - nr - 1) Is Nothing Then MathOArr(defindex - nr - 1).Range.Select
        End If
        DefFejl = True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    Else
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    End If
slut:
End Sub
Sub Integrer(variabel As String)
    Dim pkommando As String
    pkommando = Kommando

    Kommando = CodeForMaxima(Kommando)
    Kommando = "integrate(" & Kommando & "," & variabel & ")"
    If Not Radians Then
        Kommando = "ratsimp(ConvertToDegr(" & Kommando & "))"
    End If
    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
        If (InStr(pkommando, "log" & ChrW$(8289)) > 0 Or InStr(pkommando, "10" & ChrW$(12311) & "^") > 0) Then
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

    Dim nr As Integer, i As Integer
    nr = UBound(Minput)

    If defindex > 0 And nr <= 1 + defindex Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf '
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(defindex - nr) Is Nothing Then MathOArr(defindex - nr).Range.Select
        ElseIf defindex - nr - 1 >= 0 Then
            If Not MathOArr(defindex - nr - 1) Is Nothing Then MathOArr(defindex - nr - 1).Range.Select
        End If
        DefFejl = True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    End If
slut:
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
        MaximaInputStreng = AddDefinition("definer:" & TempDefs) & MaximaInputStreng
    End If
    
    MaximaInputStreng = MaximaInputStreng & "qDElist:rk(" & DElist & "," & yvarlist & "," & initguesslist & ",[" & variabel & "," & xmin & "," & xmax & "," & xstep & "])$qDElist"

    MaximaInputStreng = MaximaInputStreng & ";"
    
    AddStandardSecondary
    
    If MaximaUnits Then
        MaximaUnits = False
        Call RunMaxima
        MaximaUnits = True
    Else
        Call RunMaxima
    End If
    If StopNow Then GoTo slut
    
    Dim nr As Integer, i As Integer
    nr = UBound(Minput)

    If defindex > 0 And nr <= 1 + defindex Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf '
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(defindex - nr) Is Nothing Then MathOArr(defindex - nr).Range.Select
        ElseIf defindex - nr - 1 >= 0 Then
            If Not MathOArr(defindex - nr - 1) Is Nothing Then MathOArr(defindex - nr - 1).Range.Select
        End If
        DefFejl = True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    End If
slut:
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
        MaximaInputStreng = AddDefinition("definer:" & TempDefs) & MaximaInputStreng
    End If

    If InStr(funktion, "(") > 0 Then    ' if function
        Kommando = Replace(Kommando, funktion, "qvar")
        ea.text = Kommando
        ea.ReplaceVar Split(funktion, "(")(0), "qvar"
        Kommando = ea.text
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

    If Not Radians Then
        Kommando = "simplify(ConvertToDegr(" & Kommando & "))"
    End If
    If MaximaLogOutput = 0 And Not MaximaExact = 2 Then
        If (InStr(pkommando, "log" & ChrW$(8289)) > 0 Or InStr(pkommando, "10" & ChrW$(12311) & "^") > 0) Then
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

    If MaximaUnits Then
        MaximaUnits = False
        Call RunMaxima
        MaximaUnits = True
    Else
        Call RunMaxima
    End If
    
    If StopNow Then GoTo slut

    Dim nr As Integer, i As Integer
    nr = UBound(Minput)

    If defindex > 0 And nr <= 1 + defindex Then ' error in definition
        MaximaOutput = vbNullString
        KommentarOutput = vbNullString
        For i = 1 To nr
            KommentarOutput = KommentarOutput & Minput(i) & vbLf & moutput(i) & vbLf '
        Next
        If nr = 1 And Len(moutput(1)) < 2 Then
            If Not MathOArr(defindex - nr) Is Nothing Then MathOArr(defindex - nr).Range.Select
        ElseIf defindex - nr - 1 >= 0 Then
            If Not MathOArr(defindex - nr - 1) Is Nothing Then MathOArr(defindex - nr - 1).Range.Select
        End If
        DefFejl = True
    ElseIf nr <= 1 Then
        MaximaOutput = ""
        KommentarOutput = Minput(1) & vbLf & moutput(1)
    Else
        If funktion = "qvar" Then
            MaximaOutput = Replace(MaximaOutput, "qvar", gemfkt)
        End If
    
        If Len(MaximaOutput) > 1 Then
            MaximaOutput = RemoveBrackets(MaximaOutput)
            MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & ChrW$(8744) & "    ")    ' Comma replaced by or
        End If
    End If
slut:
End Sub
Sub Plot2D(forskrifter As String, punkter As String, Var As String, xmin As String, xmax As String, ymin As String, ymax As String, xaksetitel As String, yaksetitel As String, Optional SaveFile As Boolean = False)
' functions separated by commas
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
        text = text & ",key=""" & ConvertToAscii(arr(i)) & """,color=" & GetNextColor() & ",explicit(" & fkt & "," & Var & "," & xmin & "," & xmax & ")"
    Next
    If punkter <> "" Then
        text = text & ",key=""punkter"",point_type=filled_circle,point_size=1,points(" & punkter & ")"
    End If

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
    Dim WaitTime As Integer
    If df = "" Then
#If Mac Then
'        text = "gnuplot_command:""/Applications/WordMat/gnuplot/gnuplot""$set_plot_option([gnuplot_term, aqua])$draw2d("
    text = "gnuplot_command:""/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/MaximaWM/maxima/bin/gnuplot""$set_plot_option([gnuplot_term, aqua])$draw2d("
'    text = "gnuplot_command:""/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/gnuplot/gnuplot""$set_plot_option([gnuplot_term, aqua])$draw2d("
'        waittime = -20
#Else
        text = "draw2d("
#End If
        WaitTime = -400
    Else
#If Mac Then
    text = "gnuplot_command:""/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/MaximaWM/maxima/bin/gnuplot""$set_plot_option([gnuplot_term, aqua])$load(drawdf)$drawdf(" & df & ","
'text = "gnuplot_command:""/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/gnuplot/gnuplot""$set_plot_option([gnuplot_term, aqua])$load(drawdf)$drawdf(" & df & ","
'        waittime = -25
#Else
        text = "load(drawdf)$drawdf(" & df & ","
#End If
        WaitTime = -600
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
            '    text = text & ",user_preamble = ""set arrow from graph 0.95, first 0 to graph 1, first 0""" ' arrow on x-axis not nice
            '            text = text & ",user_preamble = ""set grid lt -1 """ ' black lines solid
            text = text & ",user_preamble = ""set grid lt 0 lc -1 lw 1"""    ' black dotted lines
            '           text = text & ",user_preamble = ""set grid lt 0 lw " & resolution & """" ' linewidth>1 does not display correctly on horizontal lines
            ' lw 2 is linewidth 2, lt is linetype, lc is linecolor
        End If
    End If
    If xaksetitel <> "" Then
        text = text & ",xlabel=""" & xaksetitel & """"
    End If
    If yaksetitel <> "" Then
        text = text & ",ylabel=""" & yaksetitel & """"
    End If

    text = text & "),numer;" ' numer introduced because there was an expression with numerical integral, where it just waited without. There is no need for exact expression when plotting.
    MaximaInputStreng = MaximaInputStreng & text
    
#If Mac Then
    If MaximaUnits Then
        MaximaUnits = False
        Call RunMaxima(WaitTime)
        MaximaUnits = True
    Else
        Call RunMaxima(WaitTime)
    End If
#Else
        Call RunMaxima(WaitTime)
#End If

End Sub
Sub PlotDF(lign As String, IndepVar As String, DepVar As String, xmin As String, xmax As String, ymin As String, ymax As String, xtrajec As String, ytrajec As String, parametre As String)
'plot of curve integrals
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

Sub Draw3D(graphobj As String, Optional antalobj As Integer = 1)
'    text = "load(draw)$draw3d("
    Dim text As String
    Dim tid As Integer

    tid = 20 + antalobj * 10

#If Mac Then
'        Text = "gnuplot_command:""/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/gnuplot/gnuplot""$set_plot_option([gnuplot_term, aqua])$draw3d("
'        Text = Text & "terminal=pdf"
'        Text = Text & ",file_name=""" & Replace(GetTempDir(), "\", "/") & "WordMatGraf""," & MacDrawDims(500, 350) & ","
#Else
    text = "draw3d("
#End If
    text = text & graphobj

    text = text & ");"
    MaximaInputStreng = MaximaInputStreng & text

#If Mac Then
'    If MaximaUnits Then
'        MaximaUnits = False
'        Call RunMaxima(-1 * tid)
'        MaximaUnits = True
'    Else
'        Call RunMaxima(-1 * tid)
'    End If
#Else
        Call RunMaxima(-1 * tid)
        If InStr(MaxProc.LastMaximaOutput, "no surface within these ranges") Then
            ' if even one equation does not exist within the selected window there is an error
            MsgBox TT.A(420), vbOKOnly, TT.Error
        End If
#End If

End Sub

Function CodeForMaxima(ByVal text As String, Optional CAStype As Integer = 0) As String
' Castype=0 is maxima, CasType=1/2 is GeoGebra (geogebra friendly. still needs convertToGeogebraSyntax)
    Dim ea As New ExpressionAnalyser
    Dim ea2 As New ExpressionAnalyser
    Dim pos As Long
    Dim posb As Long
    Dim rod As Integer
    Dim pos2 As Long
    Dim pos3 As Long
    Dim pos4 As Long
    Dim pos5 As Long
    Dim posog As Long
    Dim arr As Variant, Arr2() As String, Arr3() As String
    Dim Var As String
    Dim startn As String
    Dim slutn As String
    Dim Expr As String
    Dim sluts As String
    Dim app As String
    Dim F As String
    Dim Udtryk As String
    Dim i As Integer, j As Integer
    Dim c As String
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
    Dim subs As Integer, sups As Integer, ends As Integer
    Dim antalm As Integer, varl As Integer, di As Integer, antalmS As String

    On Error GoTo fejl
    ea.SetNormalBrackets
    ea2.StartBracket = ChrW$(12310)
    ea2.EndBracket = ChrW$(12311)

    text = Trim$(text)

    '    For j = 1 To Len(text)
    '    i = AscW(mid$(text, j, 1))
    '    s = s & mid$(text, j, 1) & " - " & i & vbCrLf
    '    Next
    '    MsgBox s

    If InStr(text, ChrW$(9633)) <> 0 Then text = Replace(text, ChrW$(9633), vbNullString) ' characters before parentheses that make them invisible. Possibly problematic as the parentheses are not removed. also makes fractions small
    If InStr(text, ChrW$(11)) <> 0 Then text = Replace(text, ChrW$(11), vbNullString)    ' shift-enter
    If InStr(text, vbLf) <> 0 Then text = Replace(text, vbLf, vbNullString)    ' shift-enter and enter
    If InStr(text, vbCrLf) <> 0 Then text = Replace(text, vbCrLf, vbNullString)
    If InStr(text, vbCr) <> 0 Then text = Replace(text, vbCr, vbNullString)
    If InStr(text, """") <> 0 Then text = Replace(text, """", vbNullString)    ' apostrophe is removed
    If InStr(text, "+-") <> 0 Then text = Replace(text, "+-", "-")    ' plus-minus is minus
    If InStr(text, ChrW$(8212)) <> 0 Then text = Replace(text, ChrW$(8212), "+")    'double minus is plus
    If InStr(text, ChrW$(8211)) <> 0 Then text = Replace(text, ChrW$(8211), "-")    'A special minus sign that is slightly shorter. May appear when copying/pasting from a book.
    If InStr(text, ChrW$(8711) & ChrW$(183)) <> 0 Then text = Replace(text, ChrW$(8711) & ChrW$(183), "divq ")    ' nabla mulltiply
    If InStr(text, ChrW$(8711) & ChrW$(215)) <> 0 Then text = Replace(text, ChrW$(8711) & ChrW$(215), "curlq ")    ' nabla cross - curl
    If InStr(text, ChrW$(8711) & "^2") <> 0 Then text = Replace(text, ChrW$(8711) & "^2", "laplaceq ")    ' nabla
    If InStr(text, ChrW$(8711)) <> 0 Then text = Replace(text, ChrW$(8711), "gradq ")    ' nabla
    If InStr(text, ChrW$(183)) <> 0 Then text = Replace(text, ChrW$(183), "*")    ' dot is replaced with times
    If InStr(text, ChrW$(8901)) <> 0 Then text = Replace(text, ChrW$(8901), "*")    ' \cdot , don't know why there seem to be two
    If InStr(text, ChrW$(8729)) <> 0 Then text = Replace(text, ChrW$(8729), "*")    ' \cdot is replaced with times
    If InStr(text, ChrW$(8226)) <> 0 Then text = Replace(text, ChrW$(8226), "*")    ' thick dot replaced with times
    If InStr(text, ChrW$(247)) <> 0 Then text = Replace(text, ChrW$(247), "/")    '
    If InStr(text, ChrW$(180) & ChrW$(180)) <> 0 Then text = Replace(text, ChrW$(180) & ChrW$(180), "^''")    ' there are three different apostrophes they are all understood as diff
    If InStr(text, ChrW$(96) & ChrW$(96)) <> 0 Then text = Replace(text, ChrW$(96) & ChrW$(96), "^''")    ' apostrophe
    If InStr(text, ChrW$(180)) <> 0 Then text = Replace(text, ChrW$(180), "^'")    ' apostrophe
    If InStr(text, ChrW$(96)) <> 0 Then text = Replace(text, ChrW$(96), "^'")    ' apostrophe
    If InStr(text, ChrW$(&H393) & "_incomplete ") <> 0 Then text = Replace(text, ChrW$(&H393) & "_incomplete ", "gamma_incomplete")
    If InStr(text, ChrW$(915)) <> 0 Then text = Replace(text, ChrW$(915), "gamma")    ' big gamma
    If InStr(text, ChrW$(8800)) <> 0 Then text = Replace(text, ChrW$(8800), "#")    ' not equal
    If InStr(text, ChrW$(8804)) <> 0 Then text = Replace(text, ChrW$(8804), "<=")    '
    If InStr(text, ChrW$(8805)) <> 0 Then text = Replace(text, ChrW$(8805), ">=")    '
    If InStr(text, ChrW$(8710)) <> 0 Then text = Replace(text, ChrW$(8710), "increment")    ' speciel delta increment
    If InStr(text, "^" & ChrW$(176) & " C") <> 0 Then text = Replace(text, "^" & ChrW$(176) & " C", "degC")    ' speciel degree(\circ) og C
    If InStr(text, "^" & ChrW$(8728) & " C") <> 0 Then text = Replace(text, "^" & ChrW$(8728) & " C", "degC")    ' speciel degree(\circ) og C
    If InStr(text, ChrW$(176) & "C") <> 0 Then text = Replace(text, ChrW$(176) & "C", "degC")    ' speciel degree(\circ) og C
    If InStr(text, ChrW$(8451)) <> 0 Then text = Replace(text, ChrW$(8451), "degC")    ' specielt oC
    If InStr(text, "^" & ChrW$(8728)) <> 0 Then text = Replace(text, "^" & ChrW$(8728), ChrW$(176))    ' another way of typing degrees
    If InStr(text, "%") <> 0 Then text = Replace(text, "%", "/100")    ' percentage symbol
    If InStr(text, ChrW$(160)) <> 0 Then text = Replace(text, ChrW$(160), "")    ' before accent e.g. vector arrow 160 8407
    If Not MaximaComplex Then
        If InStr(text, ChrW$(8736)) <> 0 Then text = Replace(text, ChrW$(8736), "Symangle")    '
    End If
    If CAStype = 0 Then
        If Radians Then
            If InStr(text, ChrW$(176)) <> 0 Then text = Replace(text, ChrW$(176), "*%pi/180")    ' degree symbol
        Else
            If InStr(text, ChrW$(176)) <> 0 Then text = Replace(text, ChrW$(176), vbNullString)    ' degree symbol
        End If
    End If
    
'If InStr(text, ChrW$()) <> 0 Then
    ' not changed back
    If InStr(text, ChrW$(230)) <> 0 Then text = Replace(text, ChrW$(230), "ae")
    If InStr(text, ChrW$(248)) <> 0 Then text = Replace(text, ChrW$(248), "oe")
    If InStr(text, ChrW$(229)) <> 0 Then text = Replace(text, ChrW$(229), "aa")
    If InStr(text, ChrW$(198)) <> 0 Then text = Replace(text, ChrW$(198), "AE")
    If InStr(text, ChrW$(216)) <> 0 Then text = Replace(text, ChrW$(216), "OE")
    If InStr(text, ChrW$(197)) <> 0 Then text = Replace(text, ChrW$(197), "AA")

    'greek letters
    If InStr(text, ChrW$(916)) <> 0 Then text = Replace(text, ChrW$(916), "Delta")
    If InStr(text, ChrW$(948)) <> 0 Then text = Replace(text, ChrW$(948), "delta")
    If InStr(text, ChrW$(945)) <> 0 Then text = Replace(text, ChrW$(945), "alpha")
    If InStr(text, ChrW$(946)) <> 0 Then text = Replace(text, ChrW$(946), "beta")
    If InStr(text, ChrW$(947)) <> 0 Then text = Replace(text, ChrW$(947), "gammaLB")
    If InStr(text, ChrW$(952)) <> 0 Then text = Replace(text, ChrW$(952), "theta")
    If InStr(text, ChrW$(920)) <> 0 Then text = Replace(text, ChrW$(920), "Theta")
    If InStr(text, ChrW$(955)) <> 0 Then text = Replace(text, ChrW$(955), "lambda")
    If InStr(text, ChrW$(923)) <> 0 Then text = Replace(text, ChrW$(923), "Lambda")
    If InStr(text, ChrW$(956)) <> 0 Then text = Replace(text, ChrW$(956), "mu")
    If InStr(text, ChrW$(961)) <> 0 Then text = Replace(text, ChrW$(961), "rho")
    If InStr(text, ChrW$(963)) <> 0 Then text = Replace(text, ChrW$(963), "sigma")
    If InStr(text, ChrW$(931)) <> 0 Then text = Replace(text, ChrW$(931), "Sigma")
    If InStr(text, ChrW$(966)) <> 0 Then text = Replace(text, ChrW$(966), "varphi")
    If InStr(text, ChrW$(981)) <> 0 Then text = Replace(text, ChrW$(981), "phi")
    If InStr(text, ChrW$(934)) <> 0 Then text = Replace(text, ChrW$(934), "Phi")
    If InStr(text, ChrW$(949)) <> 0 Then text = Replace(text, ChrW$(949), "varepsilon")
    If InStr(text, ChrW$(1013)) <> 0 Then text = Replace(text, ChrW$(1013), "epsilon")
    If InStr(text, ChrW$(968)) <> 0 Then text = Replace(text, ChrW$(968), "psi")
    If InStr(text, ChrW$(936)) <> 0 Then text = Replace(text, ChrW$(936), "Psi")
    If InStr(text, ChrW$(926)) <> 0 Then text = Replace(text, ChrW$(926), "Xi")
    If InStr(text, ChrW$(958)) <> 0 Then text = Replace(text, ChrW$(958), "xi")
    If InStr(text, ChrW$(935)) <> 0 Then text = Replace(text, ChrW$(935), "Chi")
    If InStr(text, ChrW$(967)) <> 0 Then text = Replace(text, ChrW$(967), "chi")
    If InStr(text, ChrW$(928)) <> 0 Then text = Replace(text, ChrW$(928), "Pi")
    If InStr(text, ChrW$(964)) <> 0 Then text = Replace(text, ChrW$(964), "tau")
    If InStr(text, ChrW$(957)) <> 0 Then text = Replace(text, ChrW$(957), "greek-nu")
    If InStr(text, ChrW$(954)) <> 0 Then text = Replace(text, ChrW$(954), "kappa")
    If InStr(text, ChrW$(951)) <> 0 Then text = Replace(text, ChrW$(951), "eta")
    If InStr(text, ChrW$(950)) <> 0 Then text = Replace(text, ChrW$(950), "zeta")
    If InStr(text, ChrW$(969)) <> 0 Then text = Replace(text, ChrW$(969), "omega")    ' small omega
    
    ea.text = text
    ' remove double spaces ' introduced here 1.24. Is also just before replacespacebymult. Introduced to perhaps catch some errors with many spaces
    Do While InStr(ea.text, "  ") <> 0
        ea.text = Replace(ea.text, "  ", " ")
    Loop

    ' half symbol becomes 1/2 or text Symhalf depending on context. Symhalf if letter in front
    pos = 0
    Do
        pos = InStr(pos + 1, ea.text, ChrW$(189))
        If pos <> 0 Then
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
        If pos <> 0 Then
            ea.text = text
            text = Left$(text, pos - 1) & "generalized_lambert_w(0, " & Right$(text, Len(text) - pos - 4)
        End If
    Loop While pos <> 0
    'lambertW-1
    Do
        pos = InStr(text, "W_(-1)")
        If pos <> 0 Then
            ea.text = text
            text = Left$(text, pos - 1) & "generalized_lambert_w(-1, " & Right$(text, Len(text) - pos - 7)
        End If
    Loop While pos <> 0

    ' autodetection of matrix brackets
    If InStr(text, "[") <> 0 Then
        matrixstartbracket = "["
        matrixendbracket = "]"
    End If

    ' vector symbol for SymVecta
    ea.text = text
    pos = 1
    Do
        pos = InStr(pos, ea.text, ChrW$(8407))
        If pos <> 0 Then
            matvekt = True
            If ea.ChrByIndex(pos - 1) = ")" Then
                s = ea.GetPrevBracketContent(pos)
                s = Replace(s, " ", "SPC")
                MaximaInputStreng = MaximaInputStreng & "declare(" & s & "SymVecta,nonscalar)$"
                KillDef = KillDef & s & "SymVecta,"
                'AddDefinition ("definer: declare(" & s & "SymVecta,nonscalar)")
                ea.text = Left$(ea.text, ea.pos - 2) & s & "SymVecta" & Right$(ea.text, Len(ea.text) - pos)
            Else
                '            AddDefinition ("definer: declare(" & ea.ChrByIndex(pos - 1) & "SymVecta,nonscalar)")
                s = ea.ChrByIndex(pos - 1)
                MaximaInputStreng = MaximaInputStreng & "declare(" & s & "SymVecta,nonscalar)$"
                KillDef = KillDef & s & "SymVecta,"
                ea.text = Left$(ea.text, pos - 1) & "SymVecta" & Right$(ea.text, Len(ea.text) - pos)
            End If
            pos = pos + 7
        End If
    Loop While pos <> 0

    ' vector character 2 (semi arrow) to SymVectb
    pos = 1
    Do
        pos = InStr(pos, ea.text, ChrW$(8401))
        If pos <> 0 Then
            matvekt = True
            If ea.ChrByIndex(pos - 1) = ")" Then
                s = ea.GetPrevBracketContent(pos)
                ea.text = Left$(ea.text, ea.pos - 2) & s & "SymVectb" & Right$(ea.text, Len(ea.text) - pos)
            Else
                ea.text = Left$(ea.text, pos - 1) & "SymVectb" & Right$(ea.text, Len(ea.text) - pos)
            End If
            pos = pos + 7
        End If
    Loop While pos <> 0

    ' matrix symbol for SymVectc
    pos = 1
    Do
        pos = InStr(pos, ea.text, ChrW$(773))
        If pos <> 0 Then
            matvekt = True
            If ea.ChrByIndex(pos - 1) = ")" Then
                s = ea.GetPrevBracketContent(pos)
                ea.text = Left$(ea.text, ea.pos - 2) & s & "SymVectc" & Right$(ea.text, Len(ea.text) - pos)
            Else
                ea.text = Left$(ea.text, pos - 1) & "SymVectc" & Right$(ea.text, Len(ea.text) - pos)
            End If
            pos = pos + 7
        End If
    Loop While pos <> 0


    ' matrix symbol \overbar for SymVectc
    
    pos = 1
    Do
        pos = InStr(pos, ea.text, ChrW$(175))
        If pos <> 0 Then
            matvekt = True
            s = ea.GetNextVar(pos)
            If Len(ea.text) = ea.pos Then
                ea.text = Left$(ea.text, pos - 1) & " " & s & "SymVectc"
            Else
                ea.text = Left$(ea.text, pos - 1) & " " & s & "SymVectc" & Right$(ea.text, Len(ea.text) - ea.pos)
            End If
            pos = pos + 7
        End If
    Loop While pos <> 0

    text = ea.text

    If InStr(text, ChrW$(9608)) > 0 And InStr(text, ChrW$(9508)) <= 0 Then matvekt = True    '9608=matrix symbol, however, it is also used for piecewise defined functions, where 9508 is used as the terminating character
    If InStr(text, ChrW$(9632)) > 0 Then matvekt = True    '9632= an other vector symbol
    If InStr(text, ChrW$(166)) > 0 Then matvekt = True    ' used for vectors entered from the equation menu
    
    If InStr(text, ChrW$(9508)) <> 0 Then text = Replace(text, ChrW$(9508), "")    ' empty space character at all times [ ] used for piecewise def. functions moved here 1.26 due to problems with vekt

    If matvekt Then  '9632?
        If InStr(text, ChrW$(215)) <> 0 Then text = Replace(text, ChrW$(215), "~")    ' vectorproduct
    Else
        If InStr(text, ChrW$(215)) <> 0 Then text = Replace(text, ChrW$(215), "*")    ' cross replaced with multiply
    End If

    'cross vector \hat
    pos = 1
    Do
        pos = InStr(pos, text, ChrW$(770))
        If pos <> 0 Then
            ea.text = text
            If ea.ChrByIndex(pos - 1) = ")" Then
                s = ea.GetPrevBracketContent(pos)
                text = Left$(text, ea.pos - 2) & "tvvec(" & s & ")" & Right$(text, Len(text) - pos)
                pos = pos + 4
            Else
                s = ea.GetPrevVar(pos)
                If pos - Len(s) - 2 < 0 Then
                  MsgBox "An error orrcured with hat operator. "
                  GoTo fejl
                Else
                   text = Left$(text, pos - Len(s) - 2) & "tvvec(" & Mid$(text, pos - Len(s) - 1, Len(s)) & ")" & Right$(text, Len(text) - pos)
                End If
                pos = pos + 5
            End If
        End If
    Loop While pos <> 0

    ' trig with large letters
    ea.text = text
    ea.ReplaceVar "Sin", "sin"
    ea.ReplaceVar "Cos", "cos"
    ea.ReplaceVar "Tan", "tan"
    ea.ReplaceVar "Log", "log"
    ea.ReplaceVar "Ln", "ln"
    text = ea.text

    ' remove +- symbol
    pos = InStr(text, ChrW$(177))
    If pos <> 0 Then
        pm = InputBox(TT.A(421), TT.A(422), "-")
        If pm = "+" Or LCase$(pm) = "plus" Then
            text = Replace(text, ChrW$(177), "+")
        ElseIf pm = "-" Or LCase$(pm) = "minus" Then
            text = Replace(text, ChrW$(177), "-")
        Else
            MsgBox2 TT.A(423)
            text = Replace(text, ChrW$(177), "+")
        End If
    End If

    ' commas and list separator
    If Not MaximaSeparator Then
        ea.text = text
        ea.ConvertDecSeparator
        text = ea.text
        '        text = Replace(text, ",", ".")
    End If
    If InStr(text, ";") <> 0 Then text = Replace(text, ";", ",")

'  piecewise defined functions defined with a curly bracket. fork functions
   Dim ass As String, k As Long
    ea.text = text
    ea.SetNormalBrackets
    i = 0
    Do
        pos = InStr(ea.text, "{" & ChrW$(9608))
        If pos <> 0 Then

'            pos2 = InStr(pos, ea.Text, chrw$(9508))
'            expr = mid$(ea.Text, pos + 2, pos2 - pos - 1)
            Expr = ea.GetNextBracketContent(pos)
            Arr2 = Split(Expr, "@")
            s = "(if&&"
            For j = 0 To UBound(Arr2)
               If InStr(Arr2(j), ";") > 0 Then
                  c = ";"
               ElseIf InStr(Arr2(j), ", ") > 0 Then
                  c = ", "
               ElseIf InStr(Arr2(j), " ,") > 0 Then
                  c = " ,"
               ElseIf InStr(Arr2(j), ",") > 0 Then
                  c = ","
               ElseIf InStr(Arr2(j), "|") > 0 Then
                  c = "|"
               Else
                  c = ""
               End If
                Arr3 = Split(Arr2(j), c)
                  If UBound(Arr3) > 1 Then
                     k = 1
                     ass = ""
                     Do While k <= UBound(Arr3) ' if divided incorrectly then repaired
                        If Len(ass) > 0 Or InStr(Arr3(k), "<") > 0 Or InStr(Arr3(k), ">") > 0 Or InStr(Arr3(k), ChrW$(8712)) > 0 Then
                           If ass = "" Then
                              ass = Arr3(k)
                           Else
                              ass = ass & c & Arr3(k)
                           End If
                        Else
                           Arr3(0) = Arr3(0) & c & Arr3(k)
                        End If
                        k = k + 1
                     Loop
                  ElseIf UBound(Arr3) = 1 Then
                     ass = Arr3(1)
                  End If
                If UBound(Arr3) > 0 Then
                  ass = Replace(ass, "&", "") ' If using a template from Word, there will be an & after the comma.
                  ass = Replace(ass, "" & ChrW$(8743) & " ", "&&and&&")
                  ass = Replace(ass, " " & ChrW$(8743) & " ", "&&and&&")
                  ass = Replace(ass, "" & ChrW$(8744) & " ", "&&or&&")
                  ass = Replace(ass, " " & ChrW$(8744) & " ", "&&or&&")
                  ass = Replace(ass, " AND ", "&&and&&")
                  ass = Replace(ass, " and ", "&&and&&")
                  ass = Replace(ass, " OR ", "&&or&&")
                  ass = Replace(ass, " or ", "&&or&&")
                  If CAStype = 0 Then
                    s = s & sqbrAssump(SplitAssump(Trim$(ass))) & "&&then&&(" & Trim$(Arr3(0)) & ")"
                  Else
                    s = s & ass & "&&then&&(" & Trim$(Arr3(0)) & ")"
                  End If
               Else
                    s = s & "0=0&&then&&(" & Trim$(Arr3(0)) & ")"
                End If
                If j < UBound(Arr2) Then s = s & "&&elseif&&"
            Next
            ea.text = Left$(ea.text, pos - 1) & s & ")" & Right$(ea.text, Len(ea.text) - pos - Len(Expr) - 3)
            i = i + 1
        End If
    Loop While pos > 0 And i < 1000
    text = ea.text
    

    ' |x| -> abs(x) '
    If CASengine = 0 Then
    ea.text = text
    i = 0
    Do
        pos = InStr(ea.text, ChrW$(124))
        If pos <= 0 Then pos = InStr(ea.text, ChrW$(8214))
        If pos > 0 Then
            posb = pos
            Do
                i = i + 1    ' incorrect input can otherwise create an infinite loop
                posb = posb + 1
                c = ea.ChrByIndex(posb)
                If c = "|" Or c = ChrW$(8214) Then
                    If Not (ea.IsOperator(ea.ChrByIndex(posb - 1), 4)) Then
                        If matvekt Or InStr(ea.text, ChrW$(9632)) > 0 Then
                            If CAStype = 0 Then
                                ea.text = Left$(ea.text, pos - 1) & "vecmag(" & Mid$(ea.text, pos + 1, posb - pos - 1) & ")" & Right$(ea.text, Len(ea.text) - posb)
                            Else
                                ea.text = Left$(ea.text, pos - 1) & "length(" & Mid$(ea.text, pos + 1, posb - pos - 1) & ")" & Right$(ea.text, Len(ea.text) - posb)
                            End If
                        Else
                            ea.text = Left$(ea.text, pos - 1) & "abs(" & Mid$(ea.text, pos + 1, posb - pos - 1) & ")" & Right$(ea.text, Len(ea.text) - posb)
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

            posb = InStr(ea.text, ChrW$(124))
        End If
    Loop While pos > 0 And i < 1000
    If i = 1000 Then
        MsgBox TT.A(424), vbOKOnly, TT.A(105)
    End If
    i = 0
    text = ea.text
   End If

    ' find implied parentheses after ^ and / ' (must be after diff and other functions with comma)
    ea.text = text
    ' insert parentheses at e.g. log 2 -> log(2)
    ea.InsertFunctionBrackets ' v. 1.29 moved from just after d/dx notation so that log a/b does not become log(a)/b. Tested with RunTestseq, and log a/b+2 added test
    ea.InsertBracketAfter ("/")
    ea.InsertBracketBefore ("/")
        
    ' d/dx notation
    ea.text = ConvertDifferentials(ea.text, ChrW$(8518)) ' \dd
    If Left$(ea.text, 5) = "Fejl!" Then
        text = ea.text
        ConvertErrorText = text
        GoTo slut
    End If
    ea.text = ConvertDifferentials(ea.text, ChrW$(8706)) ' \partial
    If Left$(ea.text, 5) = "Fejl!" Then
        text = ea.text
        ConvertErrorText = text
        GoTo slut
    End If
    If dAsDiffChr Then
        ea.text = ConvertDifferentials(ea.text, "d") ' normal d
        If Left$(ea.text, 5) = "Fejl!" Then
            text = ea.text
            ConvertErrorText = text
            GoTo slut
        End If
    End If
        
    If InStr(ea.text, ChrW$(12310)) <> 0 Then ea.text = Replace(ea.text, ChrW$(12310), "")    ' special invisible brackets are removed   moved from before trig with uppercase letters to make insertfunctionbrackets work
    If InStr(ea.text, ChrW$(12311)) <> 0 Then ea.text = Replace(ea.text, ChrW$(12311), "")    ' special invisible brackets are removed

    If InStr(ea.text, "cosh^(-1)") <> 0 Then ea.text = Replace(ea.text, "cosh^(-1)", "acosh")
    If InStr(ea.text, "sinh^(-1)") <> 0 Then ea.text = Replace(ea.text, "sinh^(-1)", "asinh")
    If InStr(ea.text, "tanh^(-1)") <> 0 Then ea.text = Replace(ea.text, "tanh^(-1)", "atanh")
    If InStr(ea.text, "cos^(-1) ") <> 0 Then ea.text = Replace(ea.text, "cos^(-1) ", "acos")    ' in rare circumstances there may be a space instead of a 8289 character after
    If InStr(ea.text, "sin^(-1) ") <> 0 Then ea.text = Replace(ea.text, "sin^(-1) ", "asin")    ' it will be understood as a multiplication sign if not removed
    If InStr(ea.text, "tan^(-1) ") <> 0 Then ea.text = Replace(ea.text, "tan^(-1) ", "atan")
    If InStr(ea.text, "cos^(-1)") <> 0 Then ea.text = Replace(ea.text, "cos^(-1)", "acos")
    If InStr(ea.text, "sin^(-1)") <> 0 Then ea.text = Replace(ea.text, "sin^(-1)", "asin")
    If InStr(ea.text, "tan^(-1)") <> 0 Then ea.text = Replace(ea.text, "tan^(-1)", "atan")

    ea.ConvertTrigPow ("sin")    ' must be after inverse trig
    ea.ConvertTrigPow ("cos")
    ea.ConvertTrigPow ("tan")
    text = ea.text

    ' logarithms with bases other than 10 and e
    Dim base As String
    Dim spos As Integer
    ea.text = text
    spos = 1
    Do
        '    ea.text = text
        pos = ea.FindChr("log_", spos)
        If pos <> 0 Then
            pos2 = InStr(pos + 2, ea.text, ChrW$(8289))
            If Not pos2 > 0 Then
                base = ea.GetNextVar(pos)
                base = Right$(base, Len(base) - 4)
            Else
                base = Mid$(ea.text, pos + 4, pos2 - pos - 4)
            End If
            ea.text = Left$(ea.text, pos + 2) & "base" & Right$(ea.text, Len(ea.text) - pos - Len(base) - 4)
            ea.pos = pos + 7
            '        ea.InsertBeforePos ("(")
            '        ea.pos = ea.pos + 3
            ea.GetNextBracketContent
            ea.InsertBeforePos "," & base, ea.pos - 1
            '        text = ea.text
            '        spos = ea.pos + 10
        End If
    Loop While pos <> 0
    text = ea.text


    If InStr(ea.text, ChrW$(8289)) <> 0 Then text = Replace(text, ChrW$(8289), vbNullString)    ' symbol defining function removed from word syntax

    'logarithms
    If ConvertLnLog Then
        ea.text = text
        ea.ReplaceVar "log", "log10"
        text = ea.text
    End If

    ' 3 og 4 root
    For rod = 3 To 4
        Do
            pos = InStr(text, ChrW$(8728 + rod))
            If pos > 0 Or pos4 > 0 Or pos5 > 0 Then
                ea.text = text
                ea.pos = pos + 1
                If Mid$(text, pos + 1, 1) <> "(" Then
                    ea.InsertUnderstoodBracketPair
                End If
                ea.pos = pos
                Call ea.GetNextBracketContent    ' just to find the end parenthesis
                Call ea.InsertBeforePos("^(1/" & rod & ")") ' the method here can actually cause problems with e.g. \cbrt(4)^2 as this should be interpreted as cbrt(4^2) and not cbrt(4)^2
                ea.text = Replace(ea.text, ChrW$(8728 + rod), "", 1, 1)
                text = ea.text
            End If
        Loop While pos <> 0
    Next

    'square root
    Do
        pos = InStr(text, ChrW$(8730))
        If pos <> 0 Then
            If Mid$(text, pos + 1, 1) <> "(" Then
                text = Replace(text, ChrW$(8730), "sqrt", 1, 1)
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
                    text = Replace(text, ChrW$(8730), "sqrt", 1, 1)
                ElseIf UBound(arr) = 1 Then
                    Call ea.InsertBeforePos("^(1/(" & arr(0) & ")))")
                    ea.text = Replace(ea.text, ChrW$(8730), "(", 1, 1)
                    posog = ea.FindChr("&", 1)
                    text = Left$(ea.text, pos + 1) & Right$(ea.text, Len(ea.text) - posog)
                ElseIf UBound(arr) = -1 Then
                    MsgBox2 TT.A(169), vbOKOnly, TT.Error
                    GoTo fejl
                End If
            End If
        End If
    Loop While pos <> 0

    ' sum symbol
    Do
        pos = InStr(text, ChrW$(8721))
        If pos <> 0 Then
            ea.text = text
            ea.pos = pos
            subs = ea.FindChr("_")
            sups = ea.FindChr("^")
            ends = ea.FindChr(ChrW$(9618))
            arr = Split(Mid$(text, subs + 2, sups - subs - 3), "=")
            If UBound(arr) < 1 Then    ' if definition is =, then translate to :
                arr = Split(Mid$(text, subs + 2, sups - subs - 3), ":")
            End If
            If UBound(arr) < 1 Then
                MsgBox2 TT.A(425), vbOKOnly, TT.Error
                GoTo videre
            End If
            Var = arr(0)
            startn = arr(1)
            slutn = Mid$(text, sups + 1, ends - sups - 1)
            sluts = ea.FindHierakiBreak(1)
            If ea.ChrByIndex(ea.pos) = "+" Or ea.ChrByIndex(ea.pos) = "-" Or ea.ChrByIndex(ea.pos) = "*" Or ea.ChrByIndex(ea.pos) = "/" Then sluts = sluts - 1
            Expr = Mid$(text, ends + 1, sluts - ends)
            text = Left$(text, pos - 1) & "sum(" & Expr & "," & Var & "," & startn & "," & slutn & ")" & Right$(text, Len(text) - sluts)
        End If
    Loop While pos <> 0

    ' product symbol
    Do
        pos = InStr(text, ChrW$(8719))
        If pos <> 0 Then
            ea.text = text
            ea.pos = pos
            subs = ea.FindChr("_")
            sups = ea.FindChr("^")
            ends = ea.FindChr(ChrW$(9618))
            arr = Split(Mid$(text, subs + 2, sups - subs - 3), "=")
            If UBound(arr) < 1 Then    ' if definition is =, then translate to :
                arr = Split(Mid$(text, subs + 2, sups - subs - 3), ":")
            End If
            If UBound(arr) < 1 Then
                MsgBox2 TT.A(426), vbOKOnly, TT.Error
                GoTo videre
            End If
            Var = arr(0)
            startn = arr(1)
            slutn = Mid$(text, sups + 1, ends - sups - 1)
            sluts = ea.FindHierakiBreak(1)
            Expr = Mid$(text, ends + 1, sluts - ends)
            text = Left$(text, pos - 1) & "product(" & Expr & "," & Var & "," & startn & "," & slutn & ")" & Right$(text, Len(text) - sluts)
        End If
    Loop While pos <> 0

videre:

    'limits
    pos = 1
    Do
        pos = InStr(pos, text, "lim")
        If pos <> 0 Then
            ea.text = text
            If ea.ChrByIndex(pos + 3) = ChrW$(9516) Or ea.ChrByIndex(pos + 3) = "_" Then
                ea.pos = pos
                Var = ea.GetNextBracketContent
                ends = ea.pos
                sluts = ea.FindHierakiBreak(2)
                Expr = Mid$(text, ends, sluts - ends + 1)

                arr = Split(Var, ChrW$(8594))
                Var = Trim$(arr(0))
                DiffEqVar = Var
                app = Trim$(arr(1))
                If Right$(app, 2) = "^+" Then
                    app = Left$(app, Len(app) - 2) & ",plus"
                ElseIf Right$(app, 2) = "^-" Then
                    app = Left$(app, Len(app) - 2) & ",minus"
                End If
                text = Left$(text, pos - 1) & "limit(" & Expr & "," & Var & "," & app & ")" & Right$(text, Len(text) - sluts)
            Else
                pos = pos + 3
            End If
        End If
    Loop While pos <> 0


    'derivatives
    Dim subt As String, vararr() As String, s2 As String, f2 As String
    di = 0    ' if there are more f'(2) more variables must be used
#If Mac Then
    If InStr(text, "^('')") <> 0 Then text = Replace(text, "^('')", "^''")
    If InStr(text, "^(''')") <> 0 Then text = Replace(text, "^(''')", "^'''")
    If InStr(text, "^('''')") <> 0 Then text = Replace(text, "^('''')", "^''''")
#End If
    Do
        pos = InStr(text, "^'")
        If pos <> 0 Then
            ea.text = text
            ea.pos = pos
            '        f = ea.ChrByIndex(pos - 1)
            F = ea.GetPrevVar(pos - 1)
            antalm = 0
            Do
                antalm = antalm + 1
            Loop Until ea.ChrByIndex(pos + 1 + antalm) <> "'"
            If CAStype = 0 Then
                antalmS = "round(" & antalm & ")"
            Else
                antalmS = CStr(antalm)
            End If
            If ea.ChrByIndex(pos + antalm + 1) = " " Then    ' remove any spaces after '
                text = Left$(text, pos + 1) & Right$(text, Len(text) - pos - 2)
                ea.text = text
            End If

            If ea.ChrByIndex(pos + antalm + 1) = "(" Then
                '            var = ea.ChrByIndex(pos + 3 + antalm)
                Var = ea.GetNextBracketContent(pos + 1 + antalm)
                varl = Len(Var) + 2
            Else
                If DiffEqVar = "" Then
                    Var = "x"    ' if no variable is stated then assume x
                Else
                    Var = DiffEqVar    ' if diffeq then diffeqvar is assumed
                End If
                varl = 0
            End If
            
            pos2 = InStr(F, "_")
            subt = ""
            If pos2 > 0 Then
               subt = Right$(F, Len(F) - pos2)
               For i = 1 To Len(subt)
                  If Not ea.IsLetter(Mid$(subt, i, 1)) Then
                     subt = ""
                     Exit For
                  End If
               Next
            End If

            If ea.ChrByIndex(pos - 1) = ")" Then    ' if (...)'(x)
                ' find startparanthesis?
                Udtryk = ea.GetPrevBracketContent(pos)
                If ea.ChrByIndex(ea.pos - 2) = "_" Then    '  f_(...)'(x)
                    F = ea.GetPrevVar(ea.pos - 1) & "(" & Udtryk & ")"
                    GoTo msat
                ElseIf ea.IsLetterPos(ea.pos - 2) Then    ' sin(x)'
                    F = ea.GetPrevVar(ea.pos - 1)
                    text = Left$(text, ea.pos - Len(F)) & "subst(" & Var & ",qkqg" & di & ",diff(" & F & "(qkqg" & di & "),qkqg" & di & "," & antalmS & "))" & Right$(text, Len(text) - pos - varl - antalm)
                    di = di + 1
                Else    ' (x^2)'
                    text = Left$(ea.text, ea.pos - 2) & "diff(" & Udtryk & "," & Var & "," & antalmS & ")" & Right$(ea.text, Len(ea.text) - pos - varl - antalm)
                End If
            ElseIf ea.ChrByIndex(pos - 2 - Len(subt)) = ")" And Len(subt) > 0 And Len(subt) = antalm Then ' partiel (...)_xy''
                  s = ea.GetPrevBracketContent(pos - 2)
                  For i = 1 To antalm
                     s = "diff(" & s & "," & Mid$(subt, i, 1) & ",1)"
                  Next
                 text = Left$(ea.text, ea.pos - 2) & s & Right$(ea.text, Len(ea.text) - pos - varl - antalm)
            
            ElseIf Not ea.ChrByIndex(pos + antalm + 1) = "(" Then    ' is y'  variable not function
                If Len(subt) > 0 And Len(subt) = antalm Then ' if partiel y'
                  s = Left$(F, pos2 - 1)
                  For i = 1 To antalm
                     s = "diff(" & s & "," & Mid$(subt, i, 1) & ",1)"
                  Next
                 text = Left$(text, pos - 1 - Len(F)) & s & Right$(text, Len(text) - pos - varl - antalm)
               Else
                text = Left$(text, pos - 1 - Len(F)) & "diff(" & F & "," & Var & "," & antalmS & ")" & Right$(text, Len(text) - pos - varl - antalm)
               End If
            ElseIf ea.ChrByIndex(pos + antalm + 1) = "(" Then    ' if f'(x)
msat:
                If DiffEqVar = Var Or CAStype = 1 Or CAStype = 2 Then  'if diffligning then no subst
                  If Len(subt) > 0 And Len(subt) = antalm Then ' if partiel
                        text = Left$(text, pos - 1 - Len(F)) & "diff(" & F & "(" & Var & ")," & Var & "," & antalmS & ")" & Right$(text, Len(text) - pos - varl - antalm)
                    Else
                        text = Left$(text, pos - 1 - Len(F)) & "substitute(diff(" & F & "(qkqg" & di & "),qkqg" & di & "," & antalmS & "),qkqg" & di & "," & Var & ")" & Right$(text, Len(text) - pos - varl - antalm)
                        di = di + 1
                    End If
                Else    ' ikke diff ligning
                  If Len(subt) > 0 And Len(subt) = antalm Then ' if partiel
'                     s = left$(f, pos2 - 1) & "(" & var & ")"
                     f2 = Left$(F, pos2 - 1)
                     s = "last(fundef(" & f2 & "))"
                     vararr = Split(Var, ",") ' ;?
                     For i = 1 To antalm
'                        s = "diff(" & s & "," & mid$(subt, i, 1) & ",1)"
                        s = "diff(" & s & "," & Mid$(subt, i, 1) & ",1)"
                     Next
                     s2 = "subst(["
                     For i = 0 To UBound(vararr)
                        s2 = s2 & "part(first(fundef(" & f2 & "))," & i + 1 & ")=" & vararr(i) & ","
                     Next
                     s = Left$(s2, Len(s2) - 1) & "]," & s & ")"
                    text = Left$(text, pos - 1 - Len(F)) & s & Right$(text, Len(text) - pos - varl - antalm)
                  Else ' normal not partiel f'(x)
                    text = Left$(text, pos - 1 - Len(F)) & "subst(" & Var & ",qkqg" & di & ",diff(" & F & "(qkqg" & di & "),qkqg" & di & "," & antalmS & "))" & Right$(text, Len(text) - pos - varl - antalm)
                    di = di + 1
                  End If
                End If
            Else
                GoTo fejl
            End If
        End If
    Loop While pos <> 0

    'integrals
    Dim integrand As String
    pos = InStr(text, ChrW$(8747))
    If pos <> 0 Then
        ea.text = text
        ea.text = Replace(ea.text, "diff", "qiff")    ' the d will cause problems as it will be interpreted as the end of the integral
        ea.text = Replace(ea.text, "round", "rounq")
        i = 0
        Do
            ea.pos = 0
            Var = vbNullString
            ea.StartBracket = ChrW$(8747)
            ea.EndBracket = ChrW$(8518)    'differential d
            integrand = ea.GetNextBracketContent
            pos = ea.pos - Len(integrand) - 2  'pos set to start-integral symbol
            pos2 = ea.pos
            If Len(integrand) = 0 Then ' if no differential d is used
                ea.EndBracket = "d"
                integrand = ea.GetNextBracketContent(pos - 2)
                pos = ea.pos - Len(integrand) - 2    'pos is set to the starting integral symbol
                pos2 = ea.pos
                If Not (ea.IsLetterPos(ea.pos, False) And Not (ea.IsLetterPos(ea.pos - 2, False))) Then ' if equal to d or as of variable/function that does not start with d, then it is not the end of the integral. Therefore, a simpler method is tried to find d
                    j = 0
                    Var = vbNullString
                    Do
                        j = j + 1
                        pos3 = InStr(ea.pos, ea.text, "d")
                        ea.pos = pos3 + 1
                    Loop While j < 100 And Not (ea.IsLetterPos(ea.pos, False) And Not (ea.IsLetterPos(ea.pos - 2, False))) And Not pos3 = 0
                    If pos > 0 Then
                        integrand = Mid$(ea.text, pos + 1, pos3 - pos - 1)
                        pos2 = ea.pos
                    Else ' inden d fundet
                        ConvertErrorText = TT.A(895)
                        GoTo slut
                    End If
                End If
                If Var = vbNullString Then Var = ea.GetNextVar(ea.pos - 1)
            Else
                Var = "d" & ea.GetNextVar(ea.pos - 1) ' differential d replaced by d
                ea.pos = ea.pos - 1
            End If
            If Len(integrand) > 0 Then
                If Len(integrand) < 3 And Mid$(ea.text, pos2 - Len(Var), 2 + Len(Var)) = "(" & Var & ")" Then    ' if speciel notation with \int (dx)/(...)
                    ea.SetNormalBrackets
                    integrand = ea.GetNextBracketContent(pos2 + 1)
                    pos2 = ea.pos - 1
                End If
                If Len(Var) < 2 Then
                    ConvertErrorText = TT.A(427)
                    GoTo slut
                End If
                Var = Right$(Var, Len(Var) - 1)
                integrand = Trim$(integrand)
                pos2 = pos2 + Len(Var)
                If Len(integrand) > 0 And Len(Var) > 0 Then
                    IntegrationVars = IntegrationVars & Var & ";"
                    poss = InStr(integrand, ChrW$(9618))
                    posn = InStr(integrand, "_")
                    If Mid$(integrand, posn + 1, 1) = "(" Then
                        ea2.text = integrand
                        ea2.SetNormalBrackets
                        ea2.GetNextBracketContent (posn)
                        poso = InStr(ea2.pos, integrand, "^")
                    Else
                        poso = InStr(integrand, "^")
                    End If
                    If posn > 0 And poso > 0 And posn < poss And poso < poss Then ' definite integral
                        nedre = Mid$(integrand, posn + 1, poso - posn - 1)
                        upper = Mid$(integrand, poso + 1, poss - poso - 1)
                        '        text = left$(text, posn - 1) & mid$(text, poss + 1, pos2 - poss - 1) & "," & mid$(text, pos2 + 1, 1) & "," & nedre & "," & upper & ")" & mid$(text, pos2 + 2, Len(text) - pos2 - 1)
                        If CAStype = 0 Then
                           integrand = Right$(integrand, Len(integrand) - poss) & "," & Var & "," & nedre & "," & upper & ")"
                        Else ' geogebra only supports variable x
                           ea2.text = Right$(integrand, Len(integrand) - poss)
                           If Var <> "x" Then ea2.ReplaceVar Var, "x", 1
                           integrand = ea2.text & "," & nedre & "," & upper & ")"
                        End If
                        If Left$(integrand, 1) = ChrW$(9618) Then    'remove function symbol
                            integrand = Right$(integrand, Len(integrand) - 1)
                        End If
                        If MaximaExact = 2 Then
                            '                        ea.text = Replace(ea.text, chrw$(8747), "NIntegrate(", 1, 1)
                            ea.text = Left$(ea.text, pos - 1) & "NIntegrate(" & integrand & Right$(ea.text, ea.Length - pos2 - Len(Var) + 2)
                        Else
                            ea.text = Left$(ea.text, pos - 1) & "Integrate(" & integrand & Right$(ea.text, ea.Length - pos2 - Len(Var) + 2)
                            '                        ea.text = Replace(ea.text, chrw$(8747), "Integrate(", 1, 1)
                        End If
                    Else    ' indefinite integral
                        integrand = integrand & "," & Var & ")"
                        If Left$(integrand, 1) = ChrW$(9618) Then    'remove function symbol
                            integrand = Right$(integrand, Len(integrand) - 1)
                        End If
                        ea.text = Left$(ea.text, pos - 1) & "integrate(" & integrand & Right$(ea.text, ea.Length - pos2 - Len(Var) + 2)
                    End If
                    i = i + 1
                End If
            Else
'                MsgBox2 TT.A(895), vbOKOnly, TT.Error
                ConvertErrorText = TT.A(895)
                GoTo slut
            End If
        Loop While i < 100 And InStr(ea.text, ChrW$(8747)) > 0 'Len(integrand) > 0
        ea.text = Replace(ea.text, "qiff", "diff")
        ea.text = Replace(ea.text, "rounq", "round")
        ea.SetNormalBrackets
        text = ea.text
    End If

    'double integrals
    Dim var2 As String
    pos = InStr(text, ChrW$(8748))
    If pos <> 0 Then
        ea.text = text
        ea.text = Replace(ea.text, "diff", "qiff")    ' the d will cause problems as it will be interpreted as the end of the integral
        i = 0
        Do
            ea.pos = 0
            ea.StartBracket = ChrW$(8748)
            ea.EndBracket = ChrW$(8518)    'differential d
            integrand = ea.GetNextBracketContent
            pos = ea.pos - Len(integrand)    'pos is set to the starting integral sign
            pos2 = ea.pos
            If Len(integrand) = 0 Then
                ea.EndBracket = "d"
                integrand = ea.GetNextBracketContent
                pos = ea.pos - Len(integrand) - 2    'pos is set to the starting integral sign
                pos2 = ea.pos
            End If
            If Len(integrand) > 0 Then
                Var = ea.GetNextVar(ea.pos - 1)
                If Len(Var) < 2 Then
                    ConvertErrorText = TT.A(427)
                    GoTo slut
                End If
                Var = Right$(Var, Len(Var) - 1)
                var2 = ea.GetNextVar(ea.pos + 1)
                If Len(var2) < 2 Or (Left$(var2, 1) <> "d" And Left$(var2, 1) <> ChrW$(8518)) Then
                    ConvertErrorText = TT.A(428)
                    GoTo slut
                End If
                var2 = Right$(var2, Len(var2) - 1)
                integrand = Trim$(integrand)
                pos2 = pos2 + Len(Var)
                If Len(integrand) > 0 And Len(Var) > 0 And Len(var2) > 0 Then
                    IntegrationVars = IntegrationVars & Var & ";"
                    poss = InStr(integrand, ChrW$(9618))
                    posn = InStr(integrand, "_")
                    If Mid$(integrand, posn + 1, 1) = "(" Then
                        ea2.text = integrand
                        ea2.SetNormalBrackets
                        ea2.GetNextBracketContent (posn)
                        poso = InStr(ea2.pos, integrand, "^")
                    Else
                        poso = InStr(integrand, "^")
                    End If
                    If posn > 0 And poso > 0 And posn < poss And poso < poss Then
                        ConvertErrorText = TT.A(429)
                        GoTo slut
                    Else    ' ubestemt integrale
                        integrand = integrand & "," & Var & ")," & var2 & ")"
                        If Left$(integrand, 1) = ChrW$(9618) Then    'remove function symbol
                            integrand = Right$(integrand, Len(integrand) - 1)
                        End If
                        ea.text = Left$(ea.text, pos - 1) & "integrate(integrate(" & integrand & Right$(ea.text, ea.Length - pos2 - Len(Var) - Len(var2))
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
    pos = InStr(text, ChrW$(8749))
    If pos <> 0 Then
        ea.text = text
        ea.text = Replace(ea.text, "diff", "qiff")    ' the d will cause problems as it will be interpreted as the end of the integral
        i = 0
        Do
            ea.pos = 0
            ea.StartBracket = ChrW$(8749)
            ea.EndBracket = ChrW$(8518)    'differential d
            integrand = ea.GetNextBracketContent()
            pos = ea.pos - Len(integrand)    'pos is set to the starting integral sign
            pos2 = ea.pos
            If Len(integrand) = 0 Then
                ea.EndBracket = "d"
                integrand = ea.GetNextBracketContent
                pos = ea.pos - Len(integrand) - 2    'pos is set to the starting integral sign
                pos2 = ea.pos
            End If
            If Len(integrand) > 0 Then
                Var = ea.GetNextVar(ea.pos - 1)
                If Len(Var) < 2 Then
                    ConvertErrorText = TT.A(427)
                    GoTo slut
                End If
                Var = Right$(Var, Len(Var) - 1)
                var2 = ea.GetNextVar(ea.pos + 1)
                If Len(var2) < 2 Or (Left$(var2, 1) <> "d" And Left$(var2, 1) <> ChrW$(8518)) Then
                    ConvertErrorText = TT.A(428)
                    GoTo slut
                End If
                var2 = Right$(var2, Len(var2) - 1)
                var3 = ea.GetNextVar(ea.pos + 1)
                If Len(var3) < 2 Or (Left$(var3, 1) <> "d" And Left$(var3, 1) <> ChrW$(8518)) Then
                    ConvertErrorText = TT.A(428)
                    GoTo slut
                End If
                var3 = Right$(var3, Len(var3) - 1)
                integrand = Trim$(integrand)
                pos2 = pos2 + Len(Var)
                If Len(integrand) > 0 And Len(Var) > 0 And Len(var2) > 0 Then
                    IntegrationVars = IntegrationVars & Var & ";"
                    poss = InStr(integrand, ChrW$(9618))
                    posn = InStr(integrand, "_")
                    If Mid$(integrand, posn + 1, 1) = "(" Then
                        ea2.text = integrand
                        ea2.SetNormalBrackets
                        ea2.GetNextBracketContent (posn)
                        poso = InStr(ea2.pos, integrand, "^")
                    Else
                        poso = InStr(integrand, "^")
                    End If
                    If posn > 0 And poso > 0 And posn < poss And poso < poss Then
                        ConvertErrorText = TT.A(429)
                        GoTo slut
                    Else    ' ubestemt integrale
                        integrand = integrand & "," & Var & ")," & var2 & ")," & var3 & ")"
                        If Left$(integrand, 1) = ChrW$(9618) Then    'remove function symbol
                            integrand = Right$(integrand, Len(integrand) - 1)
                        End If
                        ea.text = Left$(ea.text, pos - 1) & "integrate(integrate(integrate(" & integrand & Right$(ea.text, ea.Length - pos2 - Len(Var) - Len(var2) - Len(var3) - 2)
                    End If
                    i = i + 1
                End If
            End If
        Loop While i < 100 And Len(integrand) > 0
        ea.text = Replace(ea.text, "qiff", "diff")
        ea.SetNormalBrackets
        text = ea.text
    End If

' find implied parentheses after ^ and / ' (must be after diff and other functions with commas). Must be before ea.InsertMultSigns
' moved up to account for hidden parentheses
    ea.text = text
    ea.SetNormalBrackets
    ea.InsertBracketAfter ("^")
    text = ea.text

    ' konverter a_i - > a[i] a_i,j -> a[i,j]  if set in setting
    If MaximaIndex Then
        pos = 1
        Do
            pos = InStr(pos, ea.text, "_")
            If pos <> 0 Then
                If ea.IsLetterPos(pos - 1, False) Or ea.ChrByIndex(pos - 1) = ")" Then
                    ea.RemoveChar pos
                    ea.pos = pos
                    If ea.ChrByIndex(pos) = "(" Then
                        Expr = ea.GetNextBracketContent()
                        ea.text = Left$(ea.text, pos - 1) & "[" & Expr & "]" & Right$(ea.text, Len(ea.text) - pos - Len(Expr) - 1)
                    Else
                        posb = ea.FindHierakiBreak(5)    ' changed 5 because f_1^'(x)
                        Expr = ea.GetTextInterval(pos, posb)
                        ea.text = Left$(ea.text, pos - 1) & "[" & Expr & "]" & Right$(ea.text, Len(ea.text) - pos - Len(Expr) + 1)
                    End If
                Else
                    pos = pos + 1
                End If
            End If
        Loop While pos <> 0
    End If
spring:
    ea.RemoveSpaceAfterSubSuperscript ("_")
    '    ea.RemoveSpaceAfterSubSuperscript ("^")
    ' remove double spaces ' this was also introduced at the beginning of codeformaxima, so probably redundant here
    Do While InStr(ea.text, "  ") > 0
        ea.text = Replace(ea.text, "  ", " ")
    Loop
    ea.ReplaceSpaceByMult    ' replaces spaces with * but not if illegal

    If InStr(ea.text, " ") <> 0 Then ea.text = Replace(ea.text, " ", vbNullString)     ' any remaining spaces are removed
    '    ea.text = Replace(ea.text, " ", "*")     ' must be after diff and integral ' problematic at (x^2 )' -> diff(x^2*,x,1)

    If InStr(ea.text, "det(") <> 0 Then text = Replace(ea.text, "det(", "determinant2(")    ' a little uncertain as there may be functions that end in "det"
    
    text = ea.text
    
    
    ' Matrices
    Do
        pos = InStr(text, ChrW$(9632))
        If pos <> 0 Then
            matvekt = True
            stpos = pos - 1
            ea.text = text
            ea.pos = pos
            matrix = ea.GetNextBracketContent
            slpos = ea.pos
            If ea.ChrByIndex(slpos + 1) = "*" Then
                text = Left$(text, slpos) & " ## " & Right$(text, Len(text) - slpos - 1)
            End If
            arr = Split(matrix, "@")
            If CAStype = 1 Or CAStype = 2 Then
                mtext = "{"
                For i = 0 To UBound(arr)
                    mtext = mtext & "{" & Replace(arr(i), "&", ",") & "},"
                Next
                mtext = Left$(mtext, Len(mtext) - 1) & "}"
            Else
                mtext = "matrix("
                For i = 0 To UBound(arr)
                    mtext = mtext & "[" & Replace(arr(i), "&", ",") & "],"
                Next
                mtext = Left$(mtext, Len(mtext) - 1) & ")"
            End If
            text = Left$(text, stpos - 1) & mtext & Right$(text, Len(text) - slpos)
        End If
    Loop While pos <> 0

    'vectors
    Do
        pos = InStr(text, ChrW$(9608))
        If pos <> 0 Then
            matvekt = True
            stpos = pos - 1
            ea.text = text
            ea.pos = pos
            matrix = ea.GetNextBracketContent
            slpos = ea.pos
            arr = Split(matrix, "@")
            If CAStype = 1 Or CAStype = 2 Then
                mtext = "{"
                For i = 0 To UBound(arr)
                    mtext = mtext & "{" & arr(i) & "},"
                Next
                mtext = Left$(mtext, Len(mtext) - 1) & "}"
            Else
                mtext = "matrix("
                For i = 0 To UBound(arr)
                    mtext = mtext & "[" & arr(i) & "],"
                Next
                mtext = Left$(mtext, Len(mtext) - 1) & ")"
            End If
            
            If Len(text) - slpos >= 0 Then
                If ea.ChrByIndex(stpos) = "[" Then
                    text = Left$(text, stpos - 1) & mtext & Right$(text, Len(text) - slpos) ' remove wrapping characters if [
                Else
                    text = Left$(text, stpos) & mtext & Right$(text, Len(text) - slpos + 1) ' preserve wraparound if ( or |
                End If
'                text = left$(text, stpos - 1) & mtext & right$(text, Len(text) - slpos) ' changed 1.31 due to vector with || where brackets disappear
            Else
                text = Left$(text, stpos - 1) & mtext
            End If
        End If
    Loop While pos <> 0

    'vectors with "(1¦2)" notation. When entering vectors via equation menu | Parentheses | Examples and stacks | ( _ _ )
    Do
        pos = InStr(text, ChrW$(166))
        If pos <> 0 Then
            matvekt = True
            ea.text = text
            ea.pos = pos
'            ' identify bracket type ( [ or |
'            i = Pos - 1
'            Do While i > 0
'                If ea.ChrByIndex(i) = "(" Then
'                    ea.StartBracket = "("
'                    ea.EndBracket = ")"
'                    Exit Do
'                ElseIf ea.ChrByIndex(i) = "[" Then
'                    ea.StartBracket = "["
'                    ea.EndBracket = "]"
'                    Exit Do
'                ElseIf ea.ChrByIndex(i) = "|" Then
'                    ea.StartBracket = "|"
'                    ea.EndBracket = "|"
'                    Exit Do
'                End If
'                i = i - 1
'            Loop
            
hopBracket:
            ea.pos = ea.FindPrevBracket(pos - 1)
            stpos = ea.pos
            pos2 = stpos
            
            ea.StartBracket = "|" ' If you type || about a vector, Word can remove the () in the vector, leaving only []
            ea.EndBracket = "|"
            ea.pos = ea.FindPrevBracket(pos - 1)
            If ea.pos > stpos Then
                stpos = ea.pos
            Else
                ea.SetNormalBrackets
            End If
            
            ea.StartBracket = "["
            ea.EndBracket = "]"
            ea.pos = ea.FindPrevBracket(pos - 1)
            If ea.pos > stpos Then
                stpos = ea.pos
            Else
                ea.SetNormalBrackets
            End If
            
            matrix = ea.GetNextBracketContent(stpos)
            If Not InStr(matrix, ChrW$(166)) > 0 And pos > 0 Then
                pos = pos - 1
                GoTo hopBracket
            End If
            slpos = ea.pos - 1
            ea.pos = pos
            arr = Split(matrix, ChrW$(166))
            If CAStype = 1 Or CAStype = 2 Then
                mtext = "{"
                For i = 0 To UBound(arr)
                    mtext = mtext & "{" & arr(i) & "},"
                Next
                mtext = Left$(mtext, Len(mtext) - 1) & "}"
            Else
                mtext = "matrix("
                For i = 0 To UBound(arr)
                    mtext = mtext & "[" & arr(i) & "],"
                Next
                mtext = Left$(mtext, Len(mtext) - 1) & ")"
            End If
            
            If Len(text) - slpos >= 0 Then
'                text = left$(text, stpos - 1) & mtext & right$(text, Len(text) - slpos)
                text = Left$(text, stpos) & mtext & Right$(text, Len(text) - slpos + 1)
            Else
                text = Left$(text, stpos - 1) & mtext
            End If
        End If
    Loop While pos <> 0
    ea.SetNormalBrackets

' insert implied multiplication signs between letter space and letter
' works but then you can't do functions with subscript because there are spaces
    '    ea.text = text
    '    pos = 1
    '    Do
    '    pos = InStr(pos, ea.text, " ")
    '    If pos > 0 Then
    '        If ea.IsLetterPos(pos - 1) And ea.IsLetterPos(pos + 1) Then
    '            ea.text = left$(ea.text, pos - 1) & "*" & right$(ea.text, Len(ea.text) - pos)
    '        End If
    '        pos = pos + 2
    '    End If
    '    Loop While pos > 0
    '    text = ea.text


'insert implied multiplication signs' must be after remove space
    ea.text = text
    ea.pos = 1
    ea.InsertMultSigns
    text = ea.text
    If Radians Then
        text = Replace(text, ChrW$(8736), " CVinkel ")    ' if complex then angle is polar notation
    Else
        text = Replace(text, ChrW$(8736), " CVinkelgrad ")    ' if complex then angle is polar notation
    End If

    If InStr(text, "gradq*") <> 0 Then text = Replace(text, "gradq*", "gradq ")
    If InStr(text, "divq*") <> 0 Then text = Replace(text, "divq*", "divq ")
    If InStr(text, "curlq*") <> 0 Then text = Replace(text, "curlq*", "curlq ")
    If InStr(text, "laplaceq*") <> 0 Then text = Replace(text, "laplaceq*", "laplaceq ")


    If matvekt Then
        If CASengine = 0 Then
            If InStr(text, "^") <> 0 Then text = Replace(text, "^", "^^")  ' defined matrices were not handled correctly with the above. Now trying to replace all and hope it doesn't cause other problems
            If InStr(text, "*") <> 0 Then text = Replace(text, "*", " ## ")    ' really only important for matrix multiplication and problematic when * is not matrix multiplication, as there are then problems with hierarchy 2/3 . matrix([2],[3]) here 3 is multiplied in first. However solved with own _ operator instead of .
        Else
            If InStr(text, ChrW$(215)) <> 0 Then text = Replace(text, ChrW$(215), ChrW$(8855))
            If InStr(text, "~") <> 0 Then text = Replace(text, "~", ChrW$(8855)) ' vectorproduct
        End If
    End If

    'replace e with %e and i with %i has been moved to after implied multiplication signs as % is not considered a letter
    text = CodeConstantPercentSign(text, "e")
    If MaximaComplex Then
        If CASengine = 0 Then
            text = CodeConstantPercentSign(text, "i")
        Else
            ea.text = text
            ea.ReplaceVar "i", ChrW$(943)
            text = ea.text
        End If
    End If

    ea.text = text
    If ConvertLnLog Then
        ea.ReplaceVar "ln", "log"
    End If
    ea.ReplaceVar "arg", "carg"
    
    'Variables replaced to stop maxima output in reverse order
    If CAStype = 0 Then 'Maxima
        ea.text = ReplaceFakeVarNames(ea.text)
    End If
    text = ea.text


    ' trig functions converts to radians
    If Not Radians Then
        text = ConvertDegreeToRad(text, "sin")
        text = ConvertDegreeToRad(text, "cos")
        text = ConvertDegreeToRad(text, "tan")
        text = ConvertDegreeToRad(text, "sec")
        text = ConvertDegreeToRad(text, "cot")
        text = ConvertDegreeToRad(text, "csc")
    End If


    If CAStype = 0 Then
        If InStr(text, ChrW$(960)) <> 0 Then text = Replace(text, ChrW$(960), "%pi")
        If InStr(text, "-" & ChrW$(8734)) <> 0 Then text = Replace(text, "-" & ChrW$(8734), "minf")
        If InStr(text, ChrW$(8734)) <> 0 Then text = Replace(text, ChrW$(8734), "inf")
    Else
'        text = Replace(text, "-" & chrw$(8734), "minf")
        If InStr(text, ChrW$(8734)) <> 0 Then text = Replace(text, ChrW$(8734), "infinity")
    End If


    'units
    If MaximaUnits Then
        text = ConvertUnits(text)
    Else
        If InStr(text, ChrW$(937)) <> 0 Then text = Replace(text, ChrW$(937), "Omega")    ' omega
    End If

    If InStr(text, "&&") <> 0 Then text = Replace(text, "&&", " ") ' method to bypass automatic multiplication sign. Used with piecewise defined functions

    '    text = Replace(text, chrw$(8744), ",")  and
    '    text = Replace(text, chrw$(8743), ",") or

    GoTo slut:
fejl:
    MsgBox TT.A(99), vbOKOnly, TT.Error
slut:
    If Len(ConvertErrorText) > 0 Then
        MsgBox2 ConvertErrorText, vbOKOnly, TT.Error
        CodeForMaxima = ""
    Else
      CodeForMaxima = text
    End If
End Function
Function ConvertDifferentials(text As String, DiffChr As String) As String
' d/dx notation
' converts d/dx to maxima diff(...,x) notation
' DiffChr is the symbol for d. It could be, for example, chrw$(8518) (\dd) or chrw$(...)(\partial) or perhaps a regular d
    
    Dim num As String, dem As String, vararr(9) As String, varind As Integer, antalm2 As Integer
    Dim udtryk1 As String, udtryk2 As String, pos As Integer, pos3 As Integer, antalm As Integer
    Dim pos2 As Integer, Var As String, i As Integer, Udtryk As String, posstart As Integer, posslut As Integer
    Dim ea As New ExpressionAnalyser
    ea.SetNormalBrackets
    Do
        pos = InStr(text, "/(" & DiffChr)
        If pos <> 0 Then
            ea.text = text
            dem = ea.GetNextBracketContent(pos + 1)
            pos3 = ea.pos
            num = ea.GetPrevBracketContent(pos)
            If Left$(num, 1) <> DiffChr Then
                GoTo slut
            End If
            pos = ea.pos
            varind = 0
            If Left$(num, 2) = DiffChr & "^" Then
                antalm = Mid$(num, 3, 1)
                pos2 = InStr(dem, DiffChr)
                Do
                    If Mid$(dem, pos2 + 2, 1) = "^" Then
                        Var = Mid$(dem, pos2 + 1, 1)
                        antalm2 = Mid$(dem, pos2 + 3, 1)
                        For i = varind To varind + antalm2 - 1
                            vararr(varind) = Var
                            varind = varind + 1
                        Next
                    Else
                        vararr(varind) = Mid$(dem, pos2 + 1, 1)
                        varind = varind + 1
                    End If
                    pos2 = InStr(pos2 + 1, dem, DiffChr)
                Loop While pos2 > 0
                If antalm <> varind Then
                    text = TT.A(430)
                    GoTo slut
                End If
                ElseIf Left$(num, 1) = DiffChr Then
                    antalm = 1
                    vararr(0) = Mid$(dem, 2, 1)
                Else
                    text = TT.A(430)
                    GoTo slut
            End If

            For i = 0 To antalm - 1
                udtryk1 = udtryk1 & "diff("
                udtryk2 = udtryk2 & "," & vararr(i) & ",1)"
            Next
            If Len(num) > 1 And Mid$(num, 2, 1) <> "^" Then ' dy/dx or df(x)/dx
                Udtryk = Right$(num, Len(num) - 1)
                If Left$(Udtryk, 1) = "^" Then
                    Udtryk = Right$(Udtryk, Len(Udtryk) - 1)
                End If
                text = Left$(text, pos - 2) & udtryk1 & Udtryk & udtryk2 & Right$(text, Len(text) - pos3 + 1)
            ElseIf Mid$(text, pos3 + 1, 1) = "(" Then ' parentheses right after d/dx operator
                Udtryk = ea.GetNextBracketContent(pos3)
                text = Left$(text, pos - 2) & udtryk1 & Udtryk & udtryk2 & Right$(text, Len(text) - ea.pos + 1)
                ' Notation for df(x)/dx is missing here
            Else ' no parentheses after d/dx operator so the expression must be found via hierarchy
                posstart = pos3
                ea.pos = pos3 + 1
                posslut = ea.FindHierakiBreak(2)
                Udtryk = Trim$(Mid$(text, pos3, posslut - posstart + 1))
                text = Left$(text, pos - 2) & udtryk1 & Udtryk & udtryk2 & Right$(text, Len(text) - posslut)
            End If
        End If
    Loop While pos <> 0
slut:
    ConvertDifferentials = text
End Function
Function ConvertUnits(text As String) As String
    Dim ea As New ExpressionAnalyser
    ea.text = text
    If InStr(ea.text, ChrW$(937)) <> 0 Then ea.text = Replace(ea.text, ChrW$(937), "Ohm") ' Omega is ohm
    ea.ReplaceVar ChrW$(176) & "C", "degC"  ' degree symbol and C
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
    ea.ReplaceVar "mL", "%%mL"
    ea.ReplaceVar "muL", "%muL"
    ea.ReplaceVar "ml", "%%mL"
    ea.ReplaceVar "mul", "%muL"
    
    ConvertUnits = ea.text

End Function
Function ConvertDegreeToRad(text As String, trigfunc As String) As String
    Dim pos As Integer, spos As Integer, s As String
    Dim ea As New ExpressionAnalyser ', ea2 As New ExpressionAnalyser
    ea.SetNormalBrackets
'    ea2.SetNormalBrackets
    ea.text = text
    spos = 1

    Do
        pos = ea.FindChr("a" & trigfunc, spos)
        If pos <> 0 Then
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
    Loop While pos <> 0

    spos = 1
    Do
        pos = ea.FindChr(trigfunc, spos)
        If pos <> 0 Then
            If (pos = 1 Or Not (ea.IsLetterPos(pos - 1, False))) And Not (ea.IsLetterPos(pos + Len(trigfunc), False)) Then
                'If Not (ea.ChrByIndex(pos - 1) = "a") Then
                s = ea.GetNextBracketContent()
                If CASengine = 0 Then
                    ea.InsertBeforePos (")")
                    ea.pos = pos + Len(trigfunc)
                    ea.InsertAfterPos ("%pi/180*(")
                Else ' geogebra.
                    If InStr(s, ChrW$(176)) <= 0 Then
                     ea.InsertBeforePos (")")
                     ea.pos = pos + Len(trigfunc)
                     ea.InsertAfterPos ("pi/180*(")
                    End If
                End If
                spos = pos + 13
            Else
                spos = pos + 3
            End If
        End If
    Loop While pos <> 0

    ConvertDegreeToRad = ea.text
End Function

Function CodeConstantPercentSign(text As String, tegn As String) As String
' inserts % in front of constant in expression if there is no letter before or after
    Dim pos As Integer
    Dim ea As New ExpressionAnalyser

    ea.text = text
    pos = 1
    Do
        pos = InStr(pos, ea.text, tegn)
        If pos > 0 Then
            If Not (ea.IsLetterPos(pos + 1, False) Or ea.ChrByIndex(pos + 1) = "[" Or ea.IsLetterPos(pos - 1, False)) Then
                ea.text = Left$(ea.text, pos - 1) & "%" & Right$(ea.text, Len(ea.text) - pos + 1)
                pos = pos + 2
            Else
                pos = pos + 1
            End If
        End If
    Loop While pos > 0

    CodeConstantPercentSign = ea.text
End Function
Function ConvertToAscii(ByVal text As String) As String
' converts from word input to ascii

    text = Replace(text, ChrW$(12310), "")    ' special invisible brackets are removed
    text = Replace(text, ChrW$(12311), "")    ' special invisible brackets are removed
    text = Replace(text, ChrW$(9633), "")    ' characters before parentheses that make them invisible. Possibly problematic as the parentheses are not removed. also makes fractions small
    text = Replace(text, ChrW$(9508), "")    ' blank space character. certainly at [ ]
    text = Replace(text, ChrW$(11), "")    ' shift-enter
    text = Replace(text, vbLf, "")    ' shift-enter and enter
    text = Replace(text, vbCrLf, "")
    text = Replace(text, vbCr, "")
    text = Replace(text, """", "")    ' apostrophe is removed
    text = Replace(text, "^'", ChrW$(180))    ' The regular apostrophe cannot be used, you must use ´ (accent facing upwards)
    text = Replace(text, "'", ChrW$(180))    ' The regular apostrophe cannot be used
    text = Replace(text, "+-", "-")    ' plus-minus is minus
    text = Replace(text, ChrW$(8289), "")    ' symbol defining function removed from word syntax
    text = Replace(text, ChrW$(8212), "+")    'double minus is plus

    text = Replace(text, ChrW$(183), "*")  ' SBCL cant show cdot
    text = Replace(text, ChrW$(215), "*")  '
    text = Replace(text, "  ", " ")
    text = Replace(text, """", "")    ' apostrophe is removed
    text = Replace(text, "+-", "-")    ' plus-minus is minus
    text = Replace(text, "%e", "e")    '

    ConvertToAscii = text
End Function
Function ConvertToWordSymbols(ByVal text As String) As String
    On Error GoTo fejl
    Dim gtegn As String
    Dim ea As New ExpressionAnalyser
    Dim pos As Long, pos2 As Long
    Dim Udtryk As String
    Dim possep As Long
    Dim arr As Variant
    Dim mtext As String
    Dim slutp As Integer
    Dim s As String, s2 As String

    If text = "true" Then GoTo slut ' happens often

    ea.SetNormalBrackets
    gtegn = MaximaGangeTegn

    If CASengine = 0 Then
        If InStr(text, " ") <> 0 Then text = Replace(text, " ", vbNullString)   ' space removed
    Else
        If InStr(text, " + ") <> 0 Then text = Replace(text, " + ", "+")
        If InStr(text, " - ") <> 0 Then text = Replace(text, " - ", "-")
        If InStr(text, " * ") <> 0 Then text = Replace(text, " * ", "*")
        If InStr(text, " / ") <> 0 Then text = Replace(text, " / ", "/")
    End If
    If InStr(text, " . ") <> 0 Then text = Replace(text, " . ", "*")    ' dot product
    If InStr(text, """") <> 0 Then text = Replace(text, """", vbNullString)    ' apostrophe is removed
    If InStr(text, "+-") <> 0 Then text = Replace(text, "+-", "-")    ' plus-minus is minus
    If InStr(text, ChrW$(11)) <> 0 Then text = Replace(text, ChrW$(11), vbNullString)    ' shift-enter
    If InStr(text, vbCrLf) <> 0 Then text = Replace(text, vbCrLf, vbNullString)    ' enter
    If InStr(text, ChrW$(10)) <> 0 Then text = Replace(text, ChrW$(10), vbNullString)    ' shift-enter
    If InStr(text, ChrW$(13)) <> 0 Then text = Replace(text, ChrW$(13), vbNullString)    ' shift-enter
    If InStr(text, "'lgog") <> 0 Then text = Replace(text, "'lgog", "lgog")    '

    ea.text = text

    ea.RemoveFaultySeparator

    ' fx e^x^-1 rettes
    pos = InStr(ea.text, "^-")
    Do While pos <> 0
        ea.InsertTextAt "(", pos + 1
        ea.pos = pos + 3
        ea.FindHierakiBreak 3
        ea.InsertTextAt ")", ea.pos + 1
        pos = InStr(pos + 1, ea.text, "^-")
    Loop

    ea.InsertBracketAfter ("^")

    'exponential notation positive' before insertbracketbefore otherwise 4E-13/x -> 4E-(13)/x but c^5.0E-1 -> c^0.5*10^-1 does not work
    pos = 0
    Do
        pos = InStr(pos + 1, ea.text, "e+")
        If pos <> 1 Then
            If ea.IsNumber(pos - 1) Then
                ea.pos = pos + 1
                Do
                    ea.pos = ea.pos + 1
                Loop While ea.IsNumber(ea.pos)
                ' space removed to make insert bracket work
                ea.text = Left$(ea.text, pos - 1) & gtegn & "10^" & Mid$(ea.text, pos + 2, ea.pos - pos - 2) & " " & Right$(ea.text, Len(ea.text) - ea.pos + 1)
                'ea.text = left$(ea.text, pos - 1) & gtegn & "10^" & mid$(ea.text, pos + 2, ea.pos - pos - 2) & right$(text, Len(ea.text) - ea.pos + 1)
            End If
        End If
    Loop While pos <> 0

    ' exponential notation negative
    pos = 0
    Do
        pos = InStr(pos + 1, ea.text, "e-")
        If pos > 1 Then
            If ea.IsNumber(pos - 1) Then
                ea.pos = pos + 1
                Do
                    ea.pos = ea.pos + 1
                Loop While ea.IsNumber(ea.pos)
                ea.text = Left$(ea.text, pos - 1) & gtegn & "10^" & Mid$(ea.text, pos + 1, ea.pos - pos - 1) & " " & Right$(ea.text, Len(ea.text) - ea.pos + 1)
                '    ea.text = left$(ea.text, pos - 1) & gtegn & "10^" & mid$(ea.text, pos + 1, ea.pos - pos - 1) & right$(ea.text, Len(ea.text) - ea.pos + 1)
            End If
        End If
    Loop While pos > 1

    ' unsigned exponential notation
    pos = 0
    Do
        pos = InStr(pos + 1, ea.text, "e")
        If pos > 1 Then
            If ea.IsNumber(pos - 1) And ea.IsNumber(pos + 1) Then
                ea.pos = pos + 1
                Do
                    ea.pos = ea.pos + 1
                Loop While ea.IsNumber(ea.pos)
                ea.text = Left$(ea.text, pos - 1) & gtegn & "10^" & Mid$(ea.text, pos + 1, ea.pos - pos - 1) & " " & Right$(ea.text, Len(ea.text) - ea.pos + 1)
                '    ea.text = left$(ea.text, pos - 1) & gtegn & "10^" & mid$(ea.text, pos + 1, ea.pos - pos - 1) & right$(ea.text, Len(ea.text) - ea.pos + 1)
            End If
        End If
    Loop While pos > 1

    ' must be after exponential as 1E-1*x goes wrong
    ' 1* removed
    pos = InStr(ea.text, "1*")
    Do While pos > 0
        If pos = 1 Then
            ea.text = Left$(ea.text, pos - 1) & Right$(ea.text, Len(ea.text) - pos - 1)    ' remove 1*
        ElseIf Mid$(ea.text, pos - 1, 1) = "*" Then  ' times in front so always remove
            ea.text = Left$(ea.text, pos - 1) & Right$(ea.text, Len(ea.text) - pos - 1)    ' remove 1*
        ElseIf Mid$(ea.text, pos - 1, 1) = "(" Then  ' parentheses in front so always remove
            ea.text = Left$(ea.text, pos - 1) & Right$(ea.text, Len(ea.text) - pos - 1)    ' remove 1*
        ElseIf Mid$(ea.text, pos - 1, 1) = ")" Then  ' parentheses in front so always remove
            ea.text = Left$(ea.text, pos - 1) & Right$(ea.text, Len(ea.text) - pos - 1)    ' remove 1*
        ElseIf Mid$(ea.text, pos - 1, 1) = "=" Then  ' equal in front so always remove
            ea.text = Left$(ea.text, pos - 1) & Right$(ea.text, Len(ea.text) - pos - 1)    ' remove 1*
        ElseIf Mid$(ea.text, pos - 1, 1) = ChrW$(8776) Then  ' cirka equal in front so always remove
            ea.text = Left$(ea.text, pos - 1) & Right$(ea.text, Len(ea.text) - pos - 1)    ' remove 1*
        ElseIf ea.IsNumber(pos - 1) Then    ' number in front so do nothing

        ElseIf ea.IsLetterPos(pos - 1, True) Then    ' letter in front so do nothing

        ElseIf Mid$(ea.text, pos - 1, 1) = "." Then    ' .1* in front so do nothing

        ElseIf pos > 2 Then
            If Mid$(ea.text, pos - 2, 2) = "^-" Then    ' ^-1 dont remove
            Else
                ea.text = Left$(ea.text, pos - 1) & Right$(ea.text, Len(ea.text) - pos - 1)    ' remove 1*
            End If
        Else
            ea.text = Left$(ea.text, pos - 1) & Right$(ea.text, Len(ea.text) - pos - 1)    ' remove 1*
        End If
        pos = InStr(pos + 1, ea.text, "1*")
    Loop

    ' 1,0* removed and ,0* ->*
    pos = InStr(ea.text, ".0*")
    Do While pos <> 0
        If ea.ChrByIndex(pos - 1) = "1" Then
            If Not (ea.IsNumber(pos - 2)) Or pos = 2 Then
                ea.text = Left$(ea.text, pos - 2) & Right$(ea.text, Len(ea.text) - pos - 2)
            Else
                ea.text = Left$(ea.text, pos - 1) & Right$(ea.text, Len(ea.text) - pos - 1)    ',0* ->*
            End If
        Else
            ea.text = Left$(ea.text, pos - 1) & Right$(ea.text, Len(ea.text) - pos - 1)    ',0* ->*
        End If
        pos = InStr(pos + 1, ea.text, ".0*")
    Loop
    ' number,non-number is removed e.g. 240, -> 240 and x*,45 ->x*0,45
    
    '#If Mac Then
    '** MAC   .0 on all numbers because float is used instead of dectalallnum
    pos = InStr(ea.text, ".0")
    Do While pos > 0
        If Not (ea.IsNumber(pos + 2)) Then
            ea.text = Left$(ea.text, pos - 1) & Right$(ea.text, Len(ea.text) - pos - 1)
        End If
        pos = InStr(pos + 1, ea.text, ".0")
    Loop
    '#End If

    ' maxima 5.47.1 vil write a minus bracket in some cirsumstances  -2*x  -> -(2*x)  but only for the first term. A litte uncertain if there can be situations where the bracket should not be removed
    If Left$(ea.text, 2) = "-(" Then
        ea.GetNextBracketContent (1)
        ea.text = "-" & Mid$(ea.text, 3, ea.pos - 4) & Right$(ea.text, Len(ea.text) - ea.pos + 1)
    End If

    ' converts a[i,j] -> a_(i,j)
    If MaximaIndex Then
        pos = 1
        Do
            pos = InStr(pos + 1, ea.text, "[")
            If pos > 1 And ea.IsLetterPos(pos - 1, False) Then
                ea.pos = pos
                ea.SetSquareBrackets
                Udtryk = ea.GetNextBracketContent(pos)
                '       if ea.ChrByIndex(ea.pos)=
                ea.text = Left$(ea.text, pos - 1) & "_(" & Udtryk & ")" & Right$(ea.text, Len(ea.text) - ea.pos + 1)

            End If
        Loop While pos > 1
        ea.SetNormalBrackets
    End If

    ' remove implied multiplication signs
    ' unfortunately causes problems, among other things, in front of inverse trig 180sin^-1(12)/\pi and (1+x)=2 sin^(-1)(1/2)/\pi is converted incorrectly to prof
    ' does not work if * is replaced with a space or removed
    ' the order of these two may need to be swapped as it causes problems when identifying identical expressions for autocalc
    
    ' ea.InsertBracketBefore ("/") ' must be used for removemultsign to work, but problems occur with e.g. matrix([5/6];[13/10]) -> matrix([5/6];([13)/10])
    If MaximaUnits Then ' version 1.07 is attempted when devices are turned on
        ea.RemoveMultSign ' still problematic when e.g. e^-x^2*(e^(x^2)+2ac)/2
    End If
    If CASengine = 0 Then
        ea.SetNormalBrackets
    Else
        ea.SetTuborgBrackets
    End If
    ea.AddSpaceAfterSubSuperscript ("_")
    ea.SetNormalBrackets

    If Not MaximaSeparator Then
        If InStr(ea.text, ",") <> 0 Then ea.text = Replace(ea.text, ",", ";")
        If InStr(ea.text, ".") <> 0 Then ea.text = Replace(ea.text, ".", ",")
    End If
    If InStr(ea.text, "%pi") <> 0 Then ea.text = Replace(ea.text, "%pi", ChrW$(960))
    If InStr(ea.text, "%e") <> 0 Then ea.text = Replace(ea.text, "%e", "e")
    If InStr(ea.text, "%c") <> 0 Then ea.text = Replace(ea.text, "%c", "c")
    If InStr(ea.text, "%k") <> 0 Then ea.text = Replace(ea.text, "%k", "c_")
    If InStr(ea.text, "%r") <> 0 Then ea.text = Replace(ea.text, "%r", "k_")
    If InStr(ea.text, "%z") <> 0 Then ea.text = Replace(ea.text, "%z", ChrW$(8484) & "_")
    If InStr(ea.text, "%Z") <> 0 Then ea.text = Replace(ea.text, "%Z", ChrW$(8484))
    If InStr(ea.text, "#") <> 0 Then ea.text = Replace(ea.text, "#", ChrW$(8800))    ' not equal

    If MaximaUnits Then
        If InStr(ea.text, ChrW$(937)) <> 0 Then ea.text = Replace(ea.text, ChrW$(937), "Ohm")    ' Omega is ohm
        If InStr(ea.text, "%%m") <> 0 Then ea.text = Replace(ea.text, "%%m", "m")
        If InStr(ea.text, "%mu") <> 0 Then ea.text = Replace(ea.text, "%mu", ChrW$(956))
        ea.ReplaceVar "%timer", "timer"
        ea.ReplaceVar "%min", "min"
        ea.ReplaceVar "aar", ChrW$(229) & "r"
    End If
    ea.ReplaceVar "*degC", ChrW$(8451)  ' degree symbol and C

    text = ea.text

    If InStr(text, "*") <> 0 Then text = Replace(text, "*", gtegn)    ' chrw$(183)

    If InStr(text, "exp(") <> 0 Then text = Replace(text, "exp(", "e^(")
    If InStr(text, "log") <> 0 Then text = Replace(text, "log", "ln" & ChrW$(8289))
    If InStr(text, "lgog(") <> 0 Then text = Replace(text, "lgog(", "log" & ChrW$(8289) & "(")
    If InStr(text, "sqrt") <> 0 Then text = Replace(text, "sqrt", ChrW$(8730))
    If InStr(text, "cbrt") <> 0 Then text = Replace(text, "cbrt", ChrW$(8731)) ' from geogebra

    '    text = Replace(text, "E+", gtegn & "10^")
    '    text = Replace(text, "E-", gtegn & "10^-")
        
    'root of higher degree
    If CASengine > 0 Then 'only for GeoGebra
        Do
            pos = InStr(text, "nroot(")
            If pos <> 0 Then
                ea.text = text
                Udtryk = ea.GetNextBracketContent(pos + 4)
                pos2 = ea.pos
                '            Arr = Split(udtryk, ";")
                ea.text = Udtryk
                ea.SplitNonBracket arr, ";"
                If UBound(arr) > 0 Then
                    text = Left$(text, pos - 1) & "" & ChrW$(8730) & "(" & arr(1) & "&" & arr(0) & ")" & Right$(text, Len(text) - pos2 + 1)
                Else
                    GoTo fejl
                End If
            End If
        Loop While pos <> 0
    End If

    ' notequal(x,y) -> x=y
    Do
        pos = InStr(text, "notequal(")
        If pos <> 0 Then
            ea.text = text
            Udtryk = ea.GetNextBracketContent(pos + 7)
            text = Left$(text, pos - 1) & " " & Replace(Udtryk, ListSeparator, ChrW$(8800)) & " " & Right$(text, Len(text) - ea.pos + 1)
        End If
    Loop While pos <> 0

    ' equal(x,y) -> x=y
    Do
        pos = InStr(text, "equal(")
        If pos <> 0 Then
            ea.text = text
            Udtryk = ea.GetNextBracketContent(pos + 4)
            text = Left$(text, pos - 1) & " " & Replace(Udtryk, ListSeparator, "=") & " " & Right$(text, Len(text) - ea.pos + 1)
        End If
    Loop While pos <> 0

    ' abs(x) -> |x|
    Do
        pos = InStr(text, "abs(")
        If pos <> 0 Then
            ea.text = text
            Udtryk = ea.GetNextBracketContent(pos)
            text = Left$(text, pos - 1) & ChrW$(124) & Udtryk & ChrW$(124) & Right$(text, Len(text) - ea.pos + 1)
        End If
    Loop While pos <> 0

    ' 'vecmag(x) -> |x|
    Do
        pos = InStr(text, "'vecmag(")
        If pos <> 0 Then
            ea.text = text
            Udtryk = ea.GetNextBracketContent(pos + 6)
            text = Left$(text, pos - 1) & ChrW$(124) & Udtryk & ChrW$(124) & Right$(text, Len(text) - ea.pos + 1)
        End If
    Loop While pos <> 0

    'atan2(x,y) to atan(x/y)
    pos = InStr(text, "atan2")
    Do While pos <> 0
        ea.text = text
        ea.GetNextBracketContent (pos)
        text = Left$(text, ea.pos - 1) & ")" & Right$(text, Len(text) - ea.pos + 1)    'closing parenthesis in denominator
        text = Left$(text, pos + 3) & "(" & Right$(text, Len(text) - pos - 4)    ' , atan2 to atan(
        possep = InStr(pos + 4, text, ListSeparator)
        text = Left$(text, possep - 1) & ")/(" & Right$(text, Len(text) - possep)    ' , to /

        pos = InStr(pos, text, "atan2")
    Loop

    ' integrate to integral sign
    pos = InStr(text, "integrate")
    Do While pos <> 0
        If Mid$(text, pos - 1, 1) = "'" Then
            text = Left$(text, pos - 2) & Right$(text, Len(text) - pos + 1)
            pos = pos - 1
        End If
        Dim udt As String
        ea.text = text
        ea.pos = pos + 8
        udt = ea.GetNextBracketContent()
        arr = Split(udt, ListSeparator)
        If UBound(arr) > 1 Then
            text = Left$(text, pos - 1) & ChrW$(8747) & "_" & arr(2) & "^" & arr(3) & ChrW$(9618) & arr(0) & "  d" & arr(1) & " " & Right$(text, Len(text) - pos - Len(udt) - 10)
        Else
            text = Left$(text, pos - 1) & ChrW$(8747) & ChrW$(9618) & arr(0) & "  d" & arr(1) & " " & Right$(text, Len(text) - pos - Len(udt) - 10)
        End If
        pos = InStr(pos, text, "integrate")
    Loop

    ' differentiate
    pos = InStr(text, "diff")
    Do While pos <> 0
        If Mid$(text, pos - 1, 1) = "'" Then
            text = Left$(text, pos - 2) & Right$(text, Len(text) - pos + 1)
            pos = pos - 1
        End If
        ea.text = text
        ea.pos = pos + 4
        udt = ea.GetNextBracketContent()
        arr = Split(udt, ListSeparator)
        If UBound(arr) > 1 Then
            If arr(2) = "1" Then
                text = Left$(text, pos - 1) & ChrW$(8518) & "/(" & ChrW$(8518) & arr(1) & ") (" & arr(0) & ")" & " " & Right$(text, Len(text) - Len(udt) - 5 - pos)
            Else
                text = Left$(text, pos - 1) & ChrW$(8518) & "/(" & ChrW$(8518) & "^" & arr(2) & " " & arr(1) & ") (" & arr(0) & ")" & " " & Right$(text, Len(text) - Len(udt) - 6)
            End If
        Else
            pos = pos + 4
        End If
        pos = InStr(pos, text, "diff")
    Loop

    ' Limit
    pos = InStr(text, "limit")
    Do While pos <> 0
        If Mid$(text, pos - 1, 1) = "'" Then
            text = Left$(text, pos - 2) & Right$(text, Len(text) - pos + 1)
            pos = pos - 1
        End If
        ea.text = text
        ea.pos = pos + 4
        udt = ea.GetNextBracketContent()
        arr = Split(udt, ListSeparator)
        'lim" & chrw$(9516) & "(x" & chrw$(8594) & "1)" & chrw$(8289) & "(x+1)
        text = Left$(text, pos - 1) & "lim" & ChrW$(9516) & "(" & arr(1) & ChrW$(8594) & arr(2) & ")" & ChrW$(8289) & "(" & arr(0) & ") " & Right$(text, Len(text) - pos - Len(udt) - 6)
        pos = InStr(pos, text, "limit")
    Loop

    If Not Radians Then
        If InStr(text, "sing(") <> 0 Then text = Replace(text, "sing(", "sin(")
        If InStr(text, "asing(") <> 0 Then text = Replace(text, "asing(", "asin(")
        If InStr(text, "cosg(") <> 0 Then text = Replace(text, "cosg(", "cos(")
        If InStr(text, "acosg(") <> 0 Then text = Replace(text, "acosg(", "acos(")
        If InStr(text, "tang(") <> 0 Then text = Replace(text, "tang(", "tan(")
        If InStr(text, "atang(") <> 0 Then text = Replace(text, "atang(", "atan(")

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
    ea.ReplaceVar "cos", "cos" & ChrW$(8289)
    ea.ReplaceVar "sin", "sin" & ChrW$(8289)
    ea.ReplaceVar "tan", "tan" & ChrW$(8289)
    ea.ReplaceVar "acos", "cos^(-1)" & ChrW$(8289)
    ea.ReplaceVar "asin", "sin^(-1)" & ChrW$(8289)
    ea.ReplaceVar "atan", "tan^(-1)" & ChrW$(8289)
    ea.ReplaceVar "acosh", "cosh^(-1)" & ChrW$(8289)
    ea.ReplaceVar "asinh", "sinh^(-1)" & ChrW$(8289)
    ea.ReplaceVar "atanh", "tanh^(-1)" & ChrW$(8289)
    ea.ReplaceVar "cosh", "cosh" & ChrW$(8289)
    ea.ReplaceVar "sinh", "sinh" & ChrW$(8289)
    ea.ReplaceVar "tanh", "tanh" & ChrW$(8289)
    
    text = ReplaceFakeVarNamesBack(ea.text)    'Variables replaced to stop maxima output in reverse order

    ' matrices
    pos = InStr(text, "matrix")
    Do While pos <> 0
        ea.text = text
        mtext = ea.GetNextBracketContent(pos + 5)
        mtext = Replace(mtext, "]" & ListSeparator & "[", "@")
        mtext = Replace(mtext, ListSeparator, "&")

        mtext = matrixstartbracket & ChrW$(9632) & "(" & Mid$(mtext, 2, Len(mtext) - 2) & ")" & matrixendbracket
        text = Left$(text, pos - 1) & mtext & Right$(text, Len(text) - ea.pos + 1)

        pos = InStr(pos + 1, text, "matrix")
    Loop

    ' 'and i %if
    pos = InStr(text, "'and(")
    If pos <> 0 Then
        ea.text = text
        ea.pos = pos + 3
        udt = ea.GetNextBracketContent()
        udt = Replace(udt, ListSeparator, "  " & ChrW$(8743) & "  ")
        text = Left$(text, pos - 1) & udt & Right$(text, Len(text) - ea.pos + 1)
    End If

    ' %if
    pos = InStr(text, "%if")
    Do While pos <> 0
        ea.text = text
        mtext = ea.GetNextBracketContent(pos + 2)
        slutp = ea.pos
        arr = Split(mtext, ListSeparator)
        text = Left$(text, pos - 1) & arr(1) & "   for   " & arr(0) & Right$(text, Len(text) - slutp + 1)
        pos = InStr(text, "%if")
    Loop

    'vecta_ -> vector arrow
    pos = InStr(text, "vecta_")
    Do While pos <> 0
        ea.text = text
        s = ea.GetNextVar(pos + 6)
        text = Left$(text, pos - 1) & "(" & s & ")" & ChrW$(160) & ChrW$(8407) & Right$(text, Len(text) - ea.pos)
        pos = InStr(text, "vecta_")
    Loop
    'vectb_ -> semi vector arrow
    pos = InStr(text, "vectb_")
    Do While pos <> 0
        ea.text = text
        s = ea.GetNextVar(pos + 6)
        text = Left$(text, pos - 1) & "(" & s & ")" & ChrW$(160) & ChrW$(8401) & Right$(text, Len(text) - ea.pos)
        pos = InStr(text, "vectb_")
    Loop
    'vectc_ -> matrix line
    pos = InStr(text, "vectc_")
    Do While pos <> 0
        ea.text = text
        s = ea.GetNextVar(pos + 6)
        text = Left$(text, pos - 1) & "(" & s & ")" & ChrW$(160) & ChrW$(773) & Right$(text, Len(text) - ea.pos)
        pos = InStr(text, "vectc_")
    Loop

    'greek letters
    If InStr(text, "Delta") <> 0 Then text = Replace(text, "Delta", ChrW$(916))
    If InStr(text, "delta") <> 0 Then text = Replace(text, "delta", ChrW$(948))
    If InStr(text, "alpha") <> 0 Then text = Replace(text, "alpha", ChrW$(945))
    If InStr(text, "beta") <> 0 Then text = Replace(text, "beta", ChrW$(946))
    If InStr(text, "gamma_incomplete") <> 0 Then text = Replace(text, "gamma_incomplete", "\Gamma_incomplete ")
    If InStr(text, "gamma") <> 0 Then text = Replace(text, "gamma", ChrW$(915))
    If InStr(text, "gammaLB") <> 0 Then text = Replace(text, "gammaLB", ChrW$(947))
    If InStr(text, "theta") <> 0 Then text = Replace(text, "theta", ChrW$(952))
    If InStr(text, "Theta") <> 0 Then text = Replace(text, "Theta", ChrW$(920))
    If InStr(text, "lambda") <> 0 Then text = Replace(text, "lambda", ChrW$(955))
    If InStr(text, "Lambda") <> 0 Then text = Replace(text, "Lambda", ChrW$(923))
    If InStr(text, "rho") <> 0 Then text = Replace(text, "rho", ChrW$(961))
    If InStr(text, "varphi") <> 0 Then text = Replace(text, "varphi", ChrW$(966))
    If InStr(text, "phi") <> 0 Then text = Replace(text, "phi", ChrW$(981))
    If InStr(text, "Phi") <> 0 Then text = Replace(text, "Phi", ChrW$(934))
    If InStr(text, "varepsilon") <> 0 Then text = Replace(text, "varepsilon", ChrW$(949))
    If InStr(text, "epsilon") <> 0 Then text = Replace(text, "epsilon", ChrW$(1013))
    If InStr(text, "psi") <> 0 Then text = Replace(text, "psi", ChrW$(968))
    If InStr(text, "Psi") <> 0 Then text = Replace(text, "Psi", ChrW$(936))
    If InStr(text, "sigma") <> 0 Then text = Replace(text, "sigma", ChrW$(963))
    If InStr(text, "Sigma") <> 0 Then text = Replace(text, "Sigma", ChrW$(931))
    If InStr(text, "mu") <> 0 Then text = Replace(text, "mu", ChrW$(956))
    If InStr(text, "Ohm") <> 0 Then text = Replace(text, "Ohm", ChrW$(937))
    If InStr(text, "Omega") <> 0 Then text = Replace(text, "Omega", ChrW$(937))
    If InStr(text, "omega") <> 0 Then text = Replace(text, "omega", ChrW$(969))
    If InStr(text, "Xi") <> 0 Then text = Replace(text, "Xi", ChrW$(926))
    If InStr(text, "xi") <> 0 Then text = Replace(text, "xi", ChrW$(958))
    If InStr(text, "Chi") <> 0 Then text = Replace(text, "Chi", ChrW$(935))
    If InStr(text, "chi") <> 0 Then text = Replace(text, "chi", ChrW$(967))
    If InStr(text, "tau") <> 0 Then text = Replace(text, "tau", ChrW$(964))
    If InStr(text, "Pi") <> 0 Then text = Replace(text, "Pi", ChrW$(928))
    If InStr(text, "greek-nu") <> 0 Then text = Replace(text, "greek-nu", ChrW$(957))
    If InStr(text, "kappa") <> 0 Then text = Replace(text, "kappa", ChrW$(954))
    If InStr(text, "zeta") <> 0 Then text = Replace(text, "zeta", ChrW$(950))
    If InStr(text, "eta") <> 0 Then text = Replace(text, "eta", ChrW$(951))    ' must be last as eta is included in others
    If InStr(text, "increment") <> 0 Then text = Replace(text, "increment", ChrW$(8710))  ' specielt delta increment
    If InStr(text, "Symhalf") <> 0 Then text = Replace(text, "Symhalf", ChrW$(189))    ' _
                                                                                         If InStr(text, "degC") <> 0 Then text = Replace(text, "degC", chrw$(8451))   ' speciel oC symbol
    If InStr(text, "<=") <> 0 Then text = Replace(text, "<=", ChrW$(8804))  '
    If InStr(text, ">=") <> 0 Then text = Replace(text, ">=", ChrW$(8805))    '
    If InStr(text, "CVinkelO") <> 0 Then text = Replace(text, "CVinkelO", ChrW$(8736))    '
    If InStr(text, "CVinkel") <> 0 Then text = Replace(text, "CVinkel", ChrW$(8736))    '
    If InStr(text, "Symangle") <> 0 Then text = Replace(text, "Symangle", ChrW$(8736))  '

    ea.text = text
    ' SymVecta to vector arrow
    pos = 1
    Do
        pos = InStr(pos, ea.text, "SymVecta")
        If pos <> 0 Then
            ea.pos = pos
            s = ea.GetPrevVar(pos)
            s = Left$(s, Len(s) - (ea.pos - pos) - 1)
            If pos + 7 = Len(ea.text) Then
                ea.text = Left$(ea.text, pos - 1 - Len(s)) & "(" & s & ")" & ChrW$(8407)
            Else
                ea.text = Left$(ea.text, pos - 1 - Len(s)) & "(" & s & ")" & ChrW$(8407) & Right$(ea.text, Len(ea.text) - pos - 7)
            End If
            pos = pos + 1
        End If
    Loop While pos <> 0

    ' SymVectb to semi vector arrow
    pos = 1
    Do
        pos = InStr(pos, ea.text, "SymVectb")
        If pos <> 0 Then
            ea.pos = pos
            s = ea.GetPrevVar(pos)
            s = Left$(s, Len(s) - (ea.pos - pos) - 1)
            If pos + 7 = Len(ea.text) Then
                ea.text = Left$(ea.text, pos - 1 - Len(s)) & "(" & s & ")" & ChrW$(8401)
            Else
                ea.text = Left$(ea.text, pos - 1 - Len(s)) & "(" & s & ")" & ChrW$(8401) & Right$(ea.text, Len(ea.text) - pos - 7)
            End If
            pos = pos + 1
        End If
    Loop While pos <> 0

    ' SymVectc to semi vector arrow
    pos = 1
    Do
        pos = InStr(pos, ea.text, "SymVectc")
        If pos <> 0 Then
            ea.pos = pos
            s = ea.GetPrevVar(pos)
            s = Left$(s, Len(s) - (ea.pos - pos) - 1)
            If pos + 7 = Len(ea.text) Then
                ea.text = Left$(ea.text, pos - 1 - Len(s)) & "(" & s & ")" & ChrW$(773)
            Else
                ea.text = Left$(ea.text, pos - 1 - Len(s)) & "(" & s & ")" & ChrW$(773) & Right$(ea.text, Len(ea.text) - pos - 7)
            End If
            pos = pos + 1
        End If
    Loop While pos <> 0

    ' lambert w
    Do
        pos = InStr(ea.text, "generalized_lambert_w")
        If pos <> 0 Then
            s = ea.GetNextBracketContent(pos + 20)
            arr = Split(s, ListSeparator)
            If UBound(arr) = 0 Then
                arr = Split(s, ",")
            End If
            s2 = Trim$(arr(0))
            If (s2 = "0" Or s2 = "-1") And UBound(arr) = 1 Then
                ea.text = Left$(ea.text, pos - 1) & "W_(" & s2 & ") (" & arr(1) & ")" & Right$(ea.text, Len(ea.text) - ea.pos + 1)
            End If
        End If
    Loop While pos <> 0

    ' These may result from limit value
    ea.ReplaceVar "inf", ChrW$(8734)
    ea.ReplaceVar "minf", "-" & ChrW$(8734)
    ea.ReplaceVar "infinity", ChrW$(8734)
    ea.ReplaceVar "und", TT.A(431)
    ea.ReplaceVar "ind", TT.A(432)
    ea.ReplaceVar "tgrader", ChrW$(176)
    text = ea.text

    If InStr(text, "SPC") <> 0 Then text = Replace(text, "SPC", " ")

    'no complex expressions
    If Not MaximaComplex Then
        If InStr(text, "%i") > 0 Or InStr(text, ChrW$(943)) > 0 Then
            text = TT.A(433) & " " & ChrW$(8477)
            GoTo slut
        End If
    Else
        If InStr(text, "%i") <> 0 Then text = Replace(text, "%i", "i")
    End If

    GoTo slut
fejl:
    MsgBox TT.A(100) & vbCrLf & vbCrLf & text, vbOKOnly, TT.Error
slut:
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
            ea.InsertBeforePos (ChrW$(960) & "/180" & MaximaGangeTegn & "(")
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
            ea.InsertAfterPos ("180/" & ChrW$(960) & MaximaGangeTegn & "(")
            spos = pos + 11
        Else
            spos = spos + 3
        End If
    Loop While pos > 0

    ConvertRadToDegree = ea.text
End Function
Sub RemoveComplexSolutions()
'only real solutions' problem that even though there are in in it can perhaps be reduced to real
    Dim arr As Variant
    Dim i As Integer
    If Not MaximaComplex Then
        '        If InStr(MaximaOutput, "%i") Then
        arr = Split(MaximaOutput, ListSeparator)
        MaximaOutput = ""
        For i = 0 To UBound(arr)
            If InStr(arr(i), "%i") < 1 Then
                If MaximaOutput <> "" Then MaximaOutput = MaximaOutput & "    " & ChrW$(8744) & "    "
                MaximaOutput = MaximaOutput & arr(i)
            End If
        Next
        '        End If
    Else
        MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & ChrW$(8744) & "    ")    ' Comma replaced by or
    End If

End Sub
Function ConvertToLatex(text As String) As String
    Dim pos As Long, pos1 As Long, pos2 As Long, pos3 As Long
    Dim ea As New ExpressionAnalyser
    Dim ea2 As New ExpressionAnalyser
    Dim s As String, t As String
    Dim arr As Variant
    Dim rod As Integer
    Dim ends As Integer
    Dim subs As Integer, sups As Integer
    
    If ConvertTexWithMaxima Then
        ConvertToLatex = ConvertToLatexMaxima(text)
        Exit Function
    End If
    ea.SetNormalBrackets
    ea2.SetTuborgBrackets


    text = Replace(text, "-" & ChrW$(8734), "\minfty ")
    text = Replace(text, ChrW$(8734), "\infty ")

    'various symbols
'    text = Replace(text, chrw$(183), "* ")    ' dot is replaced with times
    text = Replace(text, ChrW$(8901), ChrW$(183))     ' \cdot , don't know why there seem to be two
    text = Replace(text, ChrW$(8729), ChrW$(183))    ' \cdot replaced by multiply (this is \bullet)
    text = Replace(text, ChrW$(8226), ChrW$(183))    ' thick dot replaced with times
    
    text = Replace(text, ChrW$(247), "\div ")    ' division sign with dots above/below


    text = Replace(text, vbCrLf, "")
    text = Replace(text, vbCr, "")
    text = Replace(text, vbLf, "")

    text = Replace(text, ChrW$(9633), "")    ' characters before parentheses that make them invisible. Possibly problematic as the parentheses are not removed. also makes fractions small
    text = Replace(text, ChrW$(9508), "")    ' blank space character certainly at [ ]
    text = Replace(text, ChrW$(11), "")    ' shift-enter
    text = Replace(text, """", "")    ' apostrophe is removed
    text = Replace(text, ChrW$(180) & ChrW$(180), "^''")    ' there are three different apostrophes that look similar. they are all understood as diff
    text = Replace(text, ChrW$(96) & ChrW$(96), "^''")    ' there are three different apostrophes that look similar. they are all understood as diff
    text = Replace(text, ChrW$(180), "^'")    ' there are three different apostrophes that look similar. they are all understood as diff
    text = Replace(text, ChrW$(96), "^'")    ' there are three different apostrophes that look similar. they are all understood as diff
    '    text = Replace(text, chrw$(176), "\circ ") ' degree symbol
    text = Replace(text, ChrW$(189), "\frac{1}{2} ")    ' _
    text = Replace(text, chrw$(8800), "\neq ")    ' not equal
    text = Replace(text, ChrW$(177), "\pm ")    ' plus/minus
    text = Replace(text, ChrW$(8723), "\mp ")    ' minus-plus
    text = Replace(text, ChrW$(8804), "\leq ")    ' <=
    text = Replace(text, ChrW$(8805), "\geq ")    ' >=
    text = Replace(text, ChrW$(8776), "\approx ")    ' cirka equal
    text = Replace(text, ChrW$(8773), "\cong ")    ' cirka equal
    text = Replace(text, ChrW$(8656), "\Leftarrow ")    '
    text = Replace(text, ChrW$(8658), "\Rightarrow ")    '
    text = Replace(text, ChrW$(8660), "\Leftrightarrow ")    ' bi implication
    text = Replace(text, ChrW$(8661), "\Updownarrow ")    ' bi implication up/down
    text = Replace(text, ChrW$(8657), "\Uparrow ")    ' doublearrow up
    text = Replace(text, ChrW$(8659), "\Downarrow ")    ' doublearrow down
    text = Replace(text, ChrW$(8593), "\uparrow ")    '
    text = Replace(text, ChrW$(8595), "\downarrow ")    '
    text = Replace(text, ChrW$(8596), "\leftrightarrow ")    '
    text = Replace(text, ChrW$(10187), "\partial ")    '
    text = Replace(text, ChrW$(8469), "\mathbb{N} ")    ' symbol for natural numbers N
    text = Replace(text, ChrW$(8484), "\mathbb{Z} ")    ' symbol for whole numbers Z
    text = Replace(text, ChrW$(8474), "\mathbb{Q} ")    ' symbol for rationel numbers Q
    text = Replace(text, ChrW$(8477), "\mathbb{R} ")    ' symbol for real numbers R
    text = Replace(text, ChrW$(8450), "\mathbb{C} ")    ' symbol for complex numbers C
    text = Replace(text, ChrW$(8463), "\hbar ")    ' planck constant/2pi
    text = Replace(text, ChrW$(8704), "\forall ")    ' math for all symbol
    text = Replace(text, ChrW$(8712), "\in ")    ' math belongs to symbol
    text = Replace(text, ChrW$(8733), "\propto ")    ' propotional symbol
    text = Replace(text, ChrW$(8810), "\ll ")    ' <<
    text = Replace(text, ChrW$(8811), "\gg ")    ' >>
    text = Replace(text, ChrW$(8704), "\exists ")    ' exist symobl
    text = Replace(text, ChrW$(8708), "\exists ")    ' exist symobl
    text = Replace(text, ChrW$(8713), "\notin ")    ' math does not belong to symbol
    text = Replace(text, ChrW$(8709), "\emptyset ")    ' empty set symbol
    text = Replace(text, ChrW$(8711), "\nabla ")    ' nabla operator
    text = Replace(text, ChrW$(8801), "\equiv ")    ' equal to with 3 dashes
    text = Replace(text, ChrW$(8710), "\inc ")    ' specielt delta increment
    text = Replace(text, ChrW$(8943), "\ldots ")    ' ...
    text = Replace(text, ChrW$(8942), "\vdots ")    ' 3 dots vertial
    text = Replace(text, ChrW$(8944), "\rddots ")    ' 3 dots diagonally up to the right
    text = Replace(text, ChrW$(8945), "\ddots ")    ' 3 dots diagonally down to the right
    text = Replace(text, ChrW$(8756), "\therefore ")    ' 3 dots in triangle
    text = Replace(text, ChrW$(8757), "\because ")    ' 3 prikker i trekant
    text = Replace(text, ChrW$(8834), "\subset ")    ' belongs to set (real subset)
    text = Replace(text, ChrW$(8838), "\subseteq ")    ' belongs to equal set
    text = Replace(text, ChrW$(8745), "\cap ")    ' common set
    text = Replace(text, ChrW$(8746), "\cup ")    ' unified set
    text = Replace(text, ChrW$(8764), "\sim ")    'ca equal (a wave)
    text = Replace(text, ChrW$(8771), "\simeq ")    ' ca equal (1 wave, 1 straight)
    text = Replace(text, ChrW$(8855), "\otimes ")    ' set product
    text = Replace(text, ChrW$(8614), "\mapsto ")    ' Arrow to function definition
    text = Replace(text, ChrW$(8640), "\rightharpoonup ")    ' right arrow with only half arrow
    text = Replace(text, ChrW$(8651), "\lrhar ")    ' harpoon right/left
    text = Replace(text, ChrW$(8598), "\nwarrow ")    ' arrow up left
    text = Replace(text, ChrW$(8599), "\nearrow ")    ' arrow up right
    text = Replace(text, ChrW$(8600), "\nwarrow ")    ' arrow down right
    text = Replace(text, ChrW$(8601), "\swarrow ")    ' down arrow left
    text = Replace(text, ChrW$(8728), "\circ ")    ' degree symbol
    text = Replace(text, ChrW$(8451), "^\circ\text{C} ")    ' speciel oC symbol
    text = Replace(text, ChrW$(8457), "^\circ\text{F} ")    ' speciel oF symbol
    text = Replace(text, ChrW$(176), "^\circ ")    '
    text = Replace(text, ChrW$(8736), "\angle ")    '
    text = Replace(text, ChrW$(160), "")    ' before accent e.g. vector arrow 160 8407
    text = Replace(text, ChrW$(8743), "\wedge ")    ' and
    text = Replace(text, ChrW$(8744), "\vee ")    ' or

'    text = Replace(text, chrw$(8721), "\sum ")    ' sum sign. Maybe it should be made more advanced
'    text = Replace(text, chrw$(8719), "\prod ")    ' product sign. Maybe needs to be made more advanced
    text = Replace(text, ChrW$(8748), "\iint ")    ' double integral sign. Maybe needs to be made more advanced
    text = Replace(text, ChrW$(8749), "\iiint ")    ' triple integral sign. Maybe needs to be made more advanced
    text = Replace(text, ChrW$(8750), "\oint ")    ' curve integral sign. Maybe needs to be made more advanced


    ' Remove problematic invisible parentheses after and before / maybe not necessary for latex? at least problematic for x/10^2 -> x/(10)^2
    '    ea2.StartBracket = chrw$(12310)
    '    ea2.EndBracket = chrw$(12311)
    '    pos = 1
    '    Do
    '    pos = InStr(pos, text, chrw$(12310))
    '    If pos > 0 Then
    '        ib = False
    '        ea2.text = text
    '        ea2.pos = pos
    '        ea2.FindPrevBreak
    '        pos2 = ea2.pos
    '        ea.text = ea2.GetNextBracketContent(pos)
    '        If ea.ValidateBrackets And InStr(ea.text, "=") < 1 Then
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

    ea.text = Replace(ea.text, ChrW$(12310), "")    ' special invisible brackets are removed moved from before trig with uppercase letters to make insertfunctionbrackets work
    ea.text = Replace(ea.text, ChrW$(12311), "")    ' special invisible brackets are removed moved

'lim
    pos = -1
    Do
        pos = InStr(pos + 2, ea.text, "lim")
        If pos > 0 And ea.ChrByIndex(pos - 1) <> "\" Then
            ea.pos = pos
            If ea.ChrByIndex(pos + 3) = ChrW$(9516) Then
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
            ends = ea.FindChr(ChrW$(8289))
            If ends > 0 Then ea.ReplaceStringAt " ", ends
            ea.text = Left$(ea.text, pos - 1) & "\lim " & Right$(ea.text, Len(ea.text) - pos - 2)
        End If
    Loop While pos > 0

'    ea.text = Replace(ea.text, "cosh^(-1)", "cosh^{-1}")
'    ea.text = Replace(ea.text, "sinh^(-1)", "sinh^{-1}")
'    ea.text = Replace(ea.text, "tanh^(-1)", "tanh^{-1}")
'    ea.text = Replace(ea.text, "cos^(-1) ", "cos^{-1}")    ' in rare circumstances there may be a space instead of a 8289 character after
'    ea.text = Replace(ea.text, "sin^(-1) ", "sin^{-1}")    ' it will be understood as a multiplication sign if not removed
'    ea.text = Replace(ea.text, "tan^(-1) ", "tan^{-1}")
'    ea.text = Replace(ea.text, "cos^(-1)", "cos^{-1}")
'    ea.text = Replace(ea.text, "sin^(-1)", "sin^{-1}")
'    ea.text = Replace(ea.text, "tan^(-1)", "tan^{-1}")

    ea.text = ConvertTrigPowLatex("sin", ea.text) ' must be before function sign is removed
    ea.text = ConvertTrigPowLatex("cos", ea.text) ' must be before function sign is removed
    ea.text = ConvertTrigPowLatex("tan", ea.text) ' must be before function sign is removed

    ea.text = Replace(ea.text, ChrW$(8289), "") ' hidden function symbol

    'fbox
    ea.text = ConvertPreAccent(ea.text, ChrW$(9645), "\fbox{", "}")

    'overbrace
    ea.text = ConvertPreAccent(ea.text, ChrW$(9182), "\overbrace{", "}")

    'underbrace
    ea.text = ConvertPreAccent(ea.text, ChrW$(9183), "\underbrace{", "}")


    'vector-notation with arrow over letter
    ea.text = ConvertAccent(ea.text, ChrW$(8407), "\overrightarrow{", "}")

    'cross vector \hat
    ea.text = ConvertAccent(ea.text, ChrW$(770), "\widehat{", "}")
    
    '\overline
    ea.text = ConvertAccent(ea.text, ChrW$(773), "\overline{", "}")
    'double overline
    ea.text = ConvertAccent(ea.text, ChrW$(831), "\overline{\overline{", "}}")
    
    'semi arrow
    ea.text = ConvertAccent(ea.text, ChrW$(8401), "\overset{\rightharpoonup}{", "}")

    'double arrow over
    ea.text = ConvertAccent(ea.text, ChrW$(8417), "\overleftrightarrow{", "}")
    
    'left arrow over
    ea.text = ConvertAccent(ea.text, ChrW$(8406), "\overleftarrow{", "}")

    '\widetilde
    ea.text = ConvertAccent(ea.text, ChrW$(771), "\widetilde{", "}")

    '\dot
    ea.text = ConvertAccent(ea.text, ChrW$(775), "\dot{", "}")
    '\ddot
    ea.text = ConvertAccent(ea.text, ChrW$(776), "\ddot{", "}")
    '\ddot
    ea.text = ConvertAccent(ea.text, ChrW$(8411), "\dddot{", "}")

    '\underset
    Do
        pos = InStr(ea.text, ChrW$(9516))
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
                    t = Mid$(ea.text, pos1, pos - pos1)
                    pos1 = pos1 - 1
                End If
            Else
                t = ea.GetPrevVar(pos)
                pos1 = ea.pos - Len(t)
            End If
            ea.text = Left$(ea.text, pos1) & "\underset{" & t & "}{" & s & "}" & Right$(ea.text, Len(ea.text) - pos2)
        End If
    Loop While pos > 0

    ' sum symbol
    Do
        pos = InStr(ea.text, ChrW$(8721))
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
            ends = ea.FindChr(ChrW$(9618))
            If ends > 0 Then ea.ReplaceStringAt " ", ends
            ea.text = Left$(ea.text, pos - 1) & "\sum " & Right$(ea.text, Len(ea.text) - pos)
        End If
    Loop While pos > 0

    ' product symbol
    Do
        pos = InStr(ea.text, ChrW$(8719))
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
            ends = ea.FindChr(ChrW$(9618))
            If ends > 0 Then ea.ReplaceStringAt " ", ends
            ea.text = Left$(ea.text, pos - 1) & "\prod " & Right$(ea.text, Len(ea.text) - pos)
        End If
    Loop While pos > 0

    ' integral
    pos = InStr(ea.text, ChrW$(8747))
    Do While pos > 0
        pos1 = InStr(pos, ea.text, "_")
        pos2 = InStr(pos, ea.text, "^")
        pos3 = InStr(pos, ea.text, ChrW$(9618))
        If pos3 = 0 Then pos3 = pos
        If pos1 > 0 Then
            Dim ul As String, ll As String
            ll = Mid$(ea.text, pos1 + 1, pos2 - pos1 - 1)
            ul = Mid$(ea.text, pos2 + 1, pos3 - pos2 - 1)
            If Left$(ll, 1) = "(" And Right$(ll, 1) = ")" Then ll = Mid$(ll, 2, Len(ll) - 2) ' Word usually inserts brackets that should not be seen. If there should be brackets, there is a double
            If Left$(ul, 1) = "(" And Right$(ul, 1) = ")" Then ul = Mid$(ul, 2, Len(ul) - 2)
            ea.text = Left$(ea.text, pos - 1) & "\int_{" & ll & "}^{" & ul & "} " & Right$(ea.text, Len(ea.text) - pos3)
        Else    ' indefinite
            ea.text = Left$(ea.text, pos - 1) & "\int " & Right$(ea.text, Len(ea.text) - pos3)
        End If
        pos = InStr(ea.text, ChrW$(8747))
    Loop

    ' ^
    pos = InStr(ea.text, "^")
    Do While pos > 0
        If ea.ChrByIndex(pos + 1) = "(" Then
            s = Trim$(ea.GetNextBracketContent(pos))
            If Not (Left$(s, 1) = "{" And Right$(s, 1) = "}") Then s = "{" & s & "}"
            ea.text = Left$(ea.text, pos) & s & Right$(ea.text, Len(ea.text) - ea.pos + 1)
        ElseIf ea.ChrByIndex(pos + 1) = "{" Then ' if already converted
        Else
            ea.pos = pos + 1
            ea.FindHierakiBreak 4, False
            s = ea.SubString(pos + 1, ea.pos)
            If Not (Left$(s, 1) = "{" And Right$(s, 1) = "}") Then s = "{" & s & "}"
            ea.text = Left$(ea.text, pos) & s & Right$(ea.text, NotZero(ea.Length - ea.pos))
        End If
        pos = InStr(pos + 1, ea.text, "^")
    Loop

    '_
    pos = InStr(ea.text, "_")
    Do While pos > 0
        If ea.ChrByIndex(pos + 1) = "(" Then
            s = ea.GetNextBracketContent
            If Not (Left$(s, 1) = "{" And Right$(s, 1) = "}") Then s = "{" & s & "}"
            ea.text = Left$(ea.text, pos) & s & Right$(ea.text, Len(ea.text) - ea.pos + 1)
        Else
            ea.pos = pos + 1
            ea.FindHierakiBreak 4, False
            s = ea.SubString(pos + 1, ea.pos)
            If Not (Left$(s, 1) = "{" And Right$(s, 1) = "}") Then s = "{" & s & "}"
            ea.text = Left$(ea.text, pos) & s & Right$(ea.text, NotZero(ea.Length - ea.pos))
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
                t = Mid$(ea.text, pos1, pos - pos1)
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
                t = Mid$(ea.text, pos1, pos - pos1)
            Else
                pos1 = ea.pos - 1
            End If
        Else
            ea.pos = pos - 1
            '            pos1 = ea.FindPrevBreak
            pos1 = ea.FindPrevHierakiBreak(4, False)
            t = Mid$(ea.text, pos1, pos - pos1)
        End If

        '        If ea.ChrByIndex(pos + 1) = "(" Then
        '            n = ea.GetNextBracketContent(pos)
        '            pos2 = ea.pos
        '        Else
        ea.pos = pos + 1
        ''            pos2 = ea.FindNextBreak
        pos2 = ea.FindHierakiBreak(3, False)
        
        n = Mid$(ea.text, pos + 1, pos2 - pos)
        If Left$(n, 1) = "(" And Right$(n, 1) = ")" Then
            n = Mid$(n, 2, Len(n) - 2)
        End If
        pos2 = pos2 + 1
        '        End If
        ea.text = Left$(ea.text, pos1 - 1) & "\frac{" & t & "}{" & n & "}" & Right$(ea.text, NotZero(Len(ea.text) - pos2 + 1))
        pos = InStr(ea.text, "/")
    Loop

    ' 3 og 4 root
    For rod = 3 To 4
        Do
            pos = InStr(ea.text, ChrW$(8728 + rod))
            If pos > 0 Then
                ea.pos = pos + 1
                If Mid$(ea.text, pos + 1, 1) <> "(" Then
                    ea.SetTuborgBrackets
                    ea.InsertUnderstoodBracketPair
                    ea.text = Replace(ea.text, ChrW$(8728 + rod), "\sqrt[" & rod & "]", 1, 1)
                Else
                    ea.SetNormalBrackets
                    s = ea.GetNextBracketContent()
                    ea.text = Left$(ea.text, pos - 1) & "\sqrt[" & rod & "]{" & s & "}" & Right$(ea.text, NotZero(ea.Length - ea.pos + 1))
                End If
                ea.pos = pos
            End If
        Loop While pos > 0
    Next

    'square root
    Do
        pos = InStr(ea.text, ChrW$(8730))
        If pos > 0 Then
            If Mid$(ea.text, pos + 1, 1) <> "(" Then
                ea.text = Replace(ea.text, ChrW$(8730), "\sqrt", 1, 1)
                pos = pos + 5
                ea.pos = pos
                ea.SetTuborgBrackets
                ea.InsertUnderstoodBracketPair
            Else
                ea.pos = pos
                arr = Split(ea.GetNextBracketContent, "&")
                pos2 = ea.pos
                If UBound(arr) = 0 Then
'                    ea.text = Replace(ea.text, chrw$(8730), "\sqrt ", 1, 1)
                    ea.text = Left$(ea.text, pos - 1) & "\sqrt{" & arr(0) & "}" & Right$(ea.text, NotZero(ea.Length - ea.pos + 1))
                ElseIf UBound(arr) = 1 Then
                    ea.text = Left$(ea.text, pos - 1) & "\sqrt[" & arr(0) & "]{" & arr(1) & "}" & Right$(ea.text, NotZero(ea.Length - ea.pos + 1))
                End If
            End If
        End If
    Loop While pos > 0

    ' matrices
    ea.text = Replace(ea.text, ChrW$(9608), ChrW$(9632))
    pos = InStr(ea.text, ChrW$(9632))
    Do While pos > 0
        ea.SetNormalBrackets
        s = ea.GetNextBracketContent(pos)
        ea.text = Left$(ea.text, pos - 1) & "\begin{matrix}" & Replace(s, "@", "\\") & "\end{matrix}" & Right$(ea.text, NotZero(ea.Length - pos - Len(s) - 2))
        pos = InStr(ea.text, ChrW$(9632))
    Loop

    ' functions
    ea.ReplaceVar "sin", "\sin "
    ea.ReplaceVar "cos", "\cos "
    ea.ReplaceVar "tan", "\tan "
    ea.ReplaceVar "log", "\log "
    ea.ReplaceVar "ln", "\ln "
    '    ea.text = Replace(ea.text, "sin" & chrw$(8289), "\sin ") '
    '    ea.text = Replace(ea.text, "cos" & chrw$(8289), "\cos ") '
    '    ea.text = Replace(ea.text, "tan" & chrw$(8289), "\tan ") '
    '    ea.text = Replace(ea.text, "log" & chrw$(8289), "\log ") '
    '    ea.text = Replace(ea.text, "ln" & chrw$(8289), "\ln ") '
    '    ea.text = Replace(ea.text, "sin^{-1}", "\sin^{-1} ") '

    ' units not in italics \,\text{cm}
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
        ea.text = Replace(ea.text, ChrW$(8289), "")    ' function symbol
        s = ea.GetNextVar()
    Loop While s <> ""
    End If

    'greek letters
    ea.text = Replace(ea.text, ChrW$(916), "\Delta ")
    ea.text = Replace(ea.text, ChrW$(948), "\delta ")
    ea.text = Replace(ea.text, ChrW$(945), "\alpha ")
    ea.text = Replace(ea.text, ChrW$(946), "\beta ")
    ea.text = Replace(ea.text, ChrW$(947), "\gammaLB ")
    ea.text = Replace(ea.text, ChrW$(952), "\theta ")
    ea.text = Replace(ea.text, ChrW$(920), "\Theta ")
    ea.text = Replace(ea.text, ChrW$(955), "\lambda ")
    ea.text = Replace(ea.text, ChrW$(923), "\Lambda ")
    ea.text = Replace(ea.text, ChrW$(956), "\mu ")
    ea.text = Replace(ea.text, ChrW$(961), "\rho ")
    ea.text = Replace(ea.text, ChrW$(963), "\sigma ")
    ea.text = Replace(ea.text, ChrW$(931), "\Sigma ")
    ea.text = Replace(ea.text, ChrW$(981), "\phi ")
    ea.text = Replace(ea.text, ChrW$(934), "\Phi ")
    ea.text = Replace(ea.text, ChrW$(949), "\varepsilon ")
    ea.text = Replace(ea.text, ChrW$(1013), "\epsilon ")
    ea.text = Replace(ea.text, ChrW$(968), "\psi ")
    ea.text = Replace(ea.text, ChrW$(936), "\Psi ")
    ea.text = Replace(ea.text, ChrW$(926), "\Xi ")
    ea.text = Replace(ea.text, ChrW$(958), "\xi ")
    ea.text = Replace(ea.text, ChrW$(935), "\Chi ")
    ea.text = Replace(ea.text, ChrW$(967), "\chi ")
    ea.text = Replace(ea.text, ChrW$(928), "\Pi ")
    ea.text = Replace(ea.text, ChrW$(964), "\tau ")
    ea.text = Replace(ea.text, ChrW$(957), "\greek-nu ")
    ea.text = Replace(ea.text, ChrW$(954), "\kappa ")
    ea.text = Replace(ea.text, ChrW$(951), "\eta ")
    ea.text = Replace(ea.text, ChrW$(950), "\zeta ")
    ea.text = Replace(ea.text, ChrW$(960), "\pi ")

    ea.text = Replace(ea.text, ChrW$(8592), "\leftarrow ")
    ea.text = Replace(ea.text, ChrW$(8594), "\rightarrow ")
    ea.text = Replace(ea.text, ChrW$(8652), "\rightleftharpoons ")    ' actually \leftrightharpoons but not supported by default Latex
    ea.text = Replace(ea.text, ChrW$(8651), "\rightleftharpoons ")
    ea.text = Replace(ea.text, ChrW$(8592), "\leftarrow ")
    ea.text = Replace(ea.text, ChrW$(10236), "\longmapsto ")
    ea.text = Replace(ea.text, ChrW$(8640), "\rightharpoonup ")
    ea.text = Replace(ea.text, ChrW$(8641), "\rightharpoondown ")
    ea.text = Replace(ea.text, ChrW$(183), "\cdot ")    ' dot is replaced with times
    ea.text = Replace(ea.text, ChrW$(215), "\times ")    ' x multiplication symbol
    ea.text = Replace(ea.text, " \", "\") ' In Moodle, spaces are created. Therefore, they should preferably be removed. Spaces before \ are not necessary.

    ' brackets
    ea.text = Replace(ea.text, "(", "\left$(")    '
    ea.text = Replace(ea.text, ")", "\right)")    '
    
    'square brackets ' not so simple because they are part of latex syntax
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
                ea.text = Left$(ea.text, pos2 - 1) & "\left" & Right$(ea.text, Len(ea.text) - pos2 + 1)
                ea.text = Left$(ea.text, pos3 + 3) & "\right" & Right$(ea.text, Len(ea.text) - pos3 - 3)
            End If
        Else
            pos1 = pos + 1
        End If
    Loop While pos > 0

    ConvertToLatex = ea.text
End Function
Function ConvertTrigPowLatex(Trig As String, text As String) As String
' converts  sin^2(x) -> sin^2(x)
Dim ea As ExpressionAnalyser
Dim pow As String
Dim pos2 As Integer, pos As Integer
pos = 1
Do
    pos = InStr(pos, text, Trig & "^")
    If pos > 0 Then
        pos = pos + Len(Trig) + 1
        pos2 = InStr(pos, text, ChrW$(8289))
        If pos2 = 0 Then
            Set ea = New ExpressionAnalyser
            ea.SetNormalBrackets
            ea.text = text
            ea.pos = pos
            pos2 = ea.FindEndOfBracket
        End If
        pow = Mid$(text, pos, pos2 - pos)
        If pow = "(-1)" Then pow = "-1"
        text = Left$(text, pos - 1) & "{" & pow & "}" & Right$(text, Len(text) - pos2)
'        InsertTextAt "^" & pow, FindEndOfBracket
        
    End If
Loop While pos > 0
ConvertTrigPowLatex = text
End Function
Function ConvertAccent(text As String, Accent As String, aString, bString) As String
' ConvertAccent(ea.text,chrw$(8407), "\overrightarrow{","}")
Dim ea As New ExpressionAnalyser, pos As Integer, pos1 As Integer, s As String
    ea.text = text
    pos = InStr(ea.text, Accent)
    Do While pos > 0
        If ea.ChrByIndex(pos + 1) = " " Then ea.RemoveChar (pos + 1) ' there may be spaces after the arrow which can be distracting
        If Mid$(ea.text, pos - 1, 1) = ")" Then
            ea.SetNormalBrackets
            s = ea.GetPrevBracketContent(pos)
            ea.text = Left$(ea.text, pos - Len(s) - 3) & aString & s & bString & Right$(ea.text, Len(ea.text) - pos)
        ElseIf Mid$(ea.text, pos - 1, 1) = "}" Then
            ea.SetTuborgBrackets
            s = ea.GetPrevBracketContent(pos)
            pos1 = ea.pos - 2
            Do While ea.IsLetterPos(pos1, False)
                pos1 = pos1 - 1
            Loop
            If ea.ChrByIndex(pos1) = "\" Then
                ea.text = Left$(ea.text, pos1 - 1) & aString & Mid$(ea.text, pos1, pos - pos1) & bString & Right$(ea.text, Len(ea.text) - pos - Len(Accent) + 1)
            End If
        Else
            ea.text = Left$(ea.text, pos - 2) & aString & Mid$(ea.text, pos - 1, 1) & bString & Right$(ea.text, Len(ea.text) - pos - Len(Accent) + 1)
        End If
        pos = InStr(pos + 1, ea.text, Accent)
    Loop
    ConvertAccent = ea.text
End Function
Function ConvertPreAccent(text As String, Accent As String, aString, bString) As String
    Dim ea As New ExpressionAnalyser
    Dim pos As Long, s As String
    ea.text = text
    Do
        pos = InStr(ea.text, Accent)
        If pos > 0 Then
            If ea.ChrByIndex(pos + 1) = "(" Then
                ea.SetNormalBrackets
                s = ea.GetNextBracketContent(pos)
                ea.text = Left$(ea.text, pos - 1) & aString & s & bString & Right$(ea.text, Len(ea.text) - ea.pos + 1)
            Else
                s = ea.GetNextVar(pos)
                ea.text = Left$(ea.text, pos - 1) & aString & s & bString & Right$(ea.text, Len(ea.text) - ea.pos)
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
    If StopNow Then GoTo slut

    '    MaximaOutput = Replace(MaximaOutput, "%i", "i")
    ConvertToLatexMaxima = MaximaOutput
slut:
End Function
Function ConvertLatexToWord(text As String) As String
    Dim ea As New ExpressionAnalyser
    Dim p As Integer
    Dim t As String, n As String

    text = Replace(text, ChrW$(11), "")    ' shift-enter
    text = Replace(text, vbCrLf, "")    ' enter
    text = Replace(text, ChrW$(10), "")    ' shift-enter
    text = Replace(text, ChrW$(13), "")    ' shift-enter
    text = Replace(text, "\infty", ChrW$(8734))    '
    text = Replace(text, "\cdot", ChrW$(183))    'dot multiplication

    'greek letters
    text = Replace(text, "\Delta", ChrW$(916))
    text = Replace(text, "\delta", ChrW$(948))
    text = Replace(text, "\alpha", ChrW$(945))
    text = Replace(text, "\beta", ChrW$(946))
    text = Replace(text, "\gamma", ChrW$(915))
    text = Replace(text, "\theta", ChrW$(952))
    text = Replace(text, "\Theta", ChrW$(920))
    text = Replace(text, "\lambda", ChrW$(955))
    text = Replace(text, "\Lambda", ChrW$(923))
    text = Replace(text, "\rho", ChrW$(961))
    text = Replace(text, "\phi", ChrW$(981))
    text = Replace(text, "\Phi", ChrW$(934))
    text = Replace(text, "\varepsilon", ChrW$(949))
    text = Replace(text, "\epsilon", ChrW$(1013))
    text = Replace(text, "\psi", ChrW$(968))
    text = Replace(text, "\Psi", ChrW$(936))
    text = Replace(text, "\sigma", ChrW$(963))
    text = Replace(text, "\Sigma", ChrW$(931))
    text = Replace(text, "\mu", ChrW$(956))
    text = Replace(text, "\Ohm", ChrW$(937))
    text = Replace(text, "\Omega", ChrW$(937))
    text = Replace(text, "omega", ChrW$(969))
    text = Replace(text, "\Xi", ChrW$(926))
    text = Replace(text, "\xi", ChrW$(958))
    text = Replace(text, "\Chi", ChrW$(935))
    text = Replace(text, "\chi", ChrW$(967))
    text = Replace(text, "\tau", ChrW$(964))
    text = Replace(text, "\Pi", ChrW$(928))
    '    text = Replace(text, "\greek-nu", chrw$(957))
    text = Replace(text, "\kappa", ChrW$(954))
    text = Replace(text, "\zeta", ChrW$(950))
    text = Replace(text, "\eta", ChrW$(951))    ' must be last as eta is included in others
    text = Replace(text, "\inc", ChrW$(8710))  ' speciel delta increment
    '    text = Replace(text, "degC", chrw$(8451))   ' speciel oC symbol
    '    text = Replace(text, "Symangle", chrw$(8736))  '

    ea.SetTuborgBrackets
    ea.text = text
    p = InStr(ea.text, "\frac")
    Do While p > 0
        t = ea.GetNextBracketContent
        n = ea.GetNextBracketContent
        ea.text = Left$(ea.text, p - 1) & "(" & t & ")/(" & n & ")" & Right$(ea.text, ea.Length - ea.pos + 1)
        p = InStr(ea.text, "\frac")
    Loop

    ConvertLatexToWord = ea.text

End Function
Sub ReadSelection()
' inserts the selected equation as a command
' this method does not use tempdoc, but changes the selected equation to text, reads it and changes it back to prof layout. This may cause changes in the way it is written.
' the method is probably a bit faster though

    Dim antal As Integer
    Dim sstart As Long
    Dim sslut As Long
    Dim j As Integer, antalgenstart As Integer
    On Error GoTo fejl

'if bold gives buildup problems' in v. 1.24.1 it is attempted to be handled without changing the original
'        For j = 1 To Selection.OMaths.Count
'            Selection.OMaths(j).Range.Font.Bold = False
'        Next
    sstart = Selection.start
    sslut = Selection.End
    antal = Selection.OMaths.Count
    Dim sr As Range
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
        Selection.OMaths(1).ParentOMath.Range.Select ' The above wasn't enough, but why was it there?
        antalgenstart = 0
genstart:
'        Kommando = ReadEquationCopy()
        Kommando = ReadEquation2()
        If Kommando = "" And antalgenstart = 0 Then    ' attempt to fix rare problem where everything is read as "" maybe it's due to tempdoc
            '            tempDoc.Close (False)
            '            Set tempDoc = Nothing
            antalgenstart = antalgenstart + 1
            GoTo genstart
        ElseIf Kommando = "" And antalgenstart = 1 Then
            RestartWordMat
            antalgenstart = antalgenstart + 1
            GoTo genstart
        ElseIf Kommando = "" Then
            MsgBox "WordMat cant read the command. This can be caused by error which can be corrected by restarting Word (Close all Word documents).", vbOKOnly, TT.Error
        End If

        If InStr(Kommando, ChrW$(8743)) Then
            Kommandoer = Split(Kommando, ChrW$(8743))
            '        For j = UBound(Kommandoer) To 0 Step -1
            '            Kommandoer(j + 1) = Kommandoer(j)
            '        Next
        End If

        Selection.Collapse (wdCollapseEnd)
    ElseIf antal > 1 Then
        ReDim Kommandoer(0 To 100)
        Dim mathcol As OMaths

        Set sr = Selection.Range
'        Selection.Copy
'        Set r = tempDoc.Range
'        r.Paste
'        With r
'            .OMaths.BuildUp
'            .OMaths.Linearize
'        End With

    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord
    
    Set mathcol = Selection.OMaths
    With Selection.OMaths
        .BuildUp
        .Linearize
    End With

        For j = 0 To antal - 1
            mathcol(j + 1).ConvertToNormalText
            Kommandoer(j) = Replace(mathcol(j + 1).Range.text, ChrW$(11), "")    ' remove shift-enter
        Next
        
    Oundo.EndCustomRecord
    ActiveDocument.Undo
        
hop:
        ReDim Preserve Kommandoer(0 To antal - 1)
        sr.Select
    End If
    GoTo slut
fejl:
    On Error GoTo slut
    ReadSelectionOld
slut:
'    tempDoc.Saved = True
    Selection.start = sstart
    Selection.End = sslut

End Sub

Function ReadEquation2(Optional ir As Range) As String
    Dim sr As Range
    Set sr = Selection.Range

    If ir Is Nothing Then
        Set ir = sr
    End If
    Dim OM As OMath
    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Set OM = ir.OMaths(1)
    OM.Range.Font.Bold = False
    Oundo.StartCustomRecord
    If OM.Range.End + 100 >= ActiveDocument.Range.End Then ' if the equation is at the very end of the document (or some enter characters after) and there are certain expressions, e.g. integrals, then a strange error occurs where the document appears empty. You can see it in draft view, however. The error does not occur if there is just one more character, so a new empty line is inserted
        OM.Range.Select
        Selection.Collapse wdCollapseEnd
        Selection.TypeParagraph
        Selection.TypeText "x"
    End If
    OM.BuildUp
    OM.Linearize
    OM.ConvertToNormalText
    ReadEquation2 = OM.Range.text
    Oundo.EndCustomRecord
    ActiveDocument.Undo
    sr.Select

Exit Function
'  below v1.30 and before
    
'    Dim r As Range
'    Set sr = Selection.Range
'    ir.Copy
'
'    Set r = tempDoc.Range
''    at.Insert r, True   ' removed 19/2 it has been replaced by paste under 'with' below?
'
'    With r
'        .Paste
'        .Font.Bold = False
'        .OMaths.BuildUp
'        .OMaths.Linearize
'        .OMaths(1).ConvertToNormalText
'    End With
'
'    ReadEquation2 = r.text
'    sr.Select
'
'    tempDoc.Range.Delete    ' this may fix Word being slow

End Function

Sub ReadSelectionOld()
' inserts the selected command
    Dim antal As Integer
    Dim sstart As Integer
    Dim sslut As Integer
    Dim j As Integer

    ' if bold causes buildup problems
    For j = 1 To Selection.OMaths.Count
        Selection.OMaths(j).Range.Font.Bold = False
    Next

    sstart = Selection.start
    sslut = Selection.End
    antal = Selection.OMaths.Count

    If antal = 0 Then
        If Len(Selection.text) < 2 Then
            Kommando = ""
            MsgBox "Mark what should be evaluated by Maxima.", vbOKOnly, "Choose"
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
        Selection.text = Replace(Selection.text, ChrW$(8727), MaximaGangeTegn)    'remove *
        Selection.text = Replace(Selection.text, "*", MaximaGangeTegn)    'remove *
        If DecSeparator = "," Then
            Selection.text = Replace(Selection.text, ".", ",")    'When comma is the decseparator, . can be changed to comma.
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
            Kommandoer(i) = Replace(mathcol(i).Range.text, ChrW$(11), "")
            mathcol(i).ConvertToMathText
            mathcol(i).BuildUp
        Next
        ReDim Preserve Kommandoer(1 To antal)
    End If
    Selection.start = sstart
    Selection.End = sslut
End Sub

Sub GoToEndOfSelectedMaths(Optional DontGoBack As Boolean = False, Optional ResultAfterNum As Boolean = True)
    Dim mc As OMaths
    Dim i As Integer
    
    If ResultAfterNum Then
        If Selection.Tables.Count > 0 Then 'nummerered equation
            If Selection.Tables(1).Columns.Count = 3 Then
                If Selection.Tables(1).Cell(1, 3).Range.Fields.Count > 0 Then
                    Selection.Tables(1).Select
                    Selection.Collapse wdCollapseEnd
                    '                Selection.MoveRight wdCharacter, 1
                    Exit Sub
                End If
            End If
        End If
    End If
    
    Selection.Collapse wdCollapseEnd
    Set mc = Selection.OMaths
    If mc.Count > 0 Then
        On Error Resume Next
        mc(mc.Count).ParentOMath.Range.Select
        On Error GoTo slut
        mc(mc.Count).Range.Select  ' works with word 2010, parentomath gives problems though. Hmm problem with selected part of expression and reducer
    ElseIf Not DontGoBack Then
        i = 0
        Do While Selection.OMaths.Count = 0 And i < 100
            Selection.MoveLeft wdCharacter, 1
            i = i + 1
        Loop
    End If
slut:
    On Error Resume Next
    Selection.Collapse wdCollapseEnd
    Dim r As Range
    Set r = Selection.Range
    r.MoveStart wdCharacter, -1
    If r.text = ChrW$(11) Then    ' if there is shift-enter at the end, replace with regular return
        r.text = ChrW$(13)
    End If
End Sub
Sub InsertMaximaOutput()
' inserts maxima output into word document
    Dim mo As Range, IsBold As Boolean, s As String

    Selection.Range.Font.Bold = False
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

    MaximaOutput = ReplaceFakeVarNamesBack(MaximaOutput)
    If Len(MaximaOutput) > 1000 Then
        If MsgBox2(TT.A(896) & Len(MaximaOutput) & VbCrLfMac & TT.A(897), vbYesNo, TT.A(846)) = vbNo Then
            MaximaOutput = Left$(MaximaOutput, 1000) & " " & TT.A(119)
        End If
    End If
    Set mo = Selection.OMaths.Add(Selection.Range)
    If mo.OMaths(1).Range.Bold Then ' There is an error in Word build that does not work correctly with bold text. For example, the solution of 2000=3000*0.85?x
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
    'Do While left$(comm, 1) = vbLf
    '    comm = right$(comm, Len(comm) - 1)
    'Loop
    ' trim linefeeds behind
    'Do While right$(comm, 1) = vbLf
    '    comm = left$(comm, Len(comm) - 1)
    'Loop

    If InStr(comm, "Solution unsure") > 0 Then
        outtext = outtext & TT.A(101)
    End If

    If InStr(comm, "solve: using arc-trig functions to get a solution.") > 0 Then
        outtext = outtext & TT.A(102)
    End If

    If InStr(comm, "Some solutions will be lost.") > 0 Then
        outtext = outtext & TT.A(103)
    End If

    If InStr(comm, "A number was found where a variable was expected -`solve'") > 0 Then
        outtext = outtext & ""
    End If

    If InStr(comm, "A number was found where a variable was expected -`solve'") > 0 Then
        outtext = outtext & TT.ErrorGeneral
    End If

    If InStr(comm, "Division by 0") > 0 Then
        outtext = outtext & TT.A(104)
    End If

    If InStr(comm, "syntax error") > 0 Then
        outtext = outtext & TT.A(105)
    End If

    If InStr(comm, "infix operator") > 0 Then
        outtext = outtext & TT.A(106)
    End If

    If InStr(comm, "solving system of equations") > 0 Then
        outtext = outtext & ""
    End If

    If InStr(comm, "Warning too few variables, only") > 0 Then
        outtext = outtext & TT.A(107)
    End If

    If InStr(comm, "More solutions may exist") > 0 Then
        outtext = outtext & TT.A(788)
    End If

    If InStr(comm, "Warning: More equations than variables.") > 0 Then
        outtext = outtext & TT.A(700)
    End If

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

    If Not (ShowAssum) Then GoTo slut

    If InStr(KommentarOutput, "Assumptions:") > 0 Then
        Dim pos As Integer
        Dim ea As New ExpressionAnalyser
        pos = InStr(KommentarOutput, "Assumptions:")
        ea.text = KommentarOutput
        ea.pos = pos
        assum = Trim$(omax.ConvertToWordSymbols(Replace(ea.GetNextBracketContent, ",", " , ")))
    End If
    If Len(assum) < 2 Then GoTo slut

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
    Selection.TypeText TT.A(900) & "  "
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

    GoTo slut
fejl:
slut:
End Sub

Function FindMaximaOutput(text As String) As String
' picks output2 from the entire maxima output
    Dim opos, ipos As Integer
    Dim text2 As String
    On Error GoTo fejl

    opos = InStr(text, "(%o2)")
    ipos = InStr(text, "(%i3)")
    'MsgBox text
    If opos < 1 Then GoTo fejl

    text2 = Trim$(Mid$(text, opos + 5, ipos - opos - 5))
    text2 = Replace(text2, "[", "")
    text2 = Replace(text2, "]", "")

    FindMaximaOutput = text2

    GoTo slut
fejl:
    opos = InStr(text, "(%i2)")
    text2 = Right$(text, Len(text) - opos - 4)
    text2 = Replace(text2, "(%i2)", "")
    text2 = TT.A(108) & vbCrLf & vbCrLf & text2
    MsgBox text2, vbOKOnly, TT.A(109)
    FindMaximaOutput = ""
slut:

End Function
Sub FindLastMaximaOutput(Optional ConvertToWord As Boolean = True, Optional Mcommand As String = "")
' picks last output from entire maxima output
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
    
    GoTo slut
fejl:
slut:
End Sub
Sub GenerateOutputArray(Optional text As String)
    Dim Output As String
    Dim nr As Integer
    Dim AntalCifre As Integer
    Dim pos As Long
    Dim pos2 As Long
    Dim posb As Long
    Dim ArrIndex As Integer
    ArrIndex = 1

    On Error GoTo fejl
#If Mac Then
#Else
    If text = "" Then text = MaxProc.LastMaximaOutput
#End If
    'text = right$(text, Len(text) - 270)
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
            GoTo slut
        End If
        If pos2 < 1 Then
            pos2 = Len(text)
            slutnu = True
        End If
        posb = InStr(pos, text, ")")

        nr = CInt(Mid$(text, pos + 4, posb - pos - 4))
        If nr < 10 Then
            AntalCifre = 1
        ElseIf nr < 100 Then
            AntalCifre = 2
        ElseIf nr < 1000 Then
            AntalCifre = 3
        Else
            AntalCifre = 4
        End If

        Output = Mid$(text, posb + 1, pos2 - pos - 4 - AntalCifre)
        If Mid$(text, pos + 3, 1) = "o" Then Output = RemovePrintetBool(Output)
        If Mid$(text, pos + 3, 1) = "i" Then
            ' next is input
            Minput(ArrIndex) = Minput(ArrIndex) & Output
        Else
'            Output = Replace(trim$(Output), VbCrLfMac, "") ' v. 1.29 this now only for output, as input can be errors, with multiple lines. Here it is important with arrow
            Output = Trim$(Output) ' replace vbcrlfmac removed 1.29 for the '23,,' the arrow showing the error was not placed correctly on the next line
            ' next is output
            moutput(ArrIndex) = moutput(ArrIndex) & Output
            ArrIndex = ArrIndex + 1
        End If
        pos = pos2
'        pos2 = InStr(pos + 3, text, vbLf & "(%")
      pos2 = FindNextio(pos + 3, text) - 1
    Loop Until slutnu
    nr = ArrIndex
    If moutput(nr) = "" And Len(Minput(nr)) < 6 Then
        nr = nr - 1
    End If

    '    ShowOutputArray ' for test

    GoTo slut
fejl:
slut:
   If nr < 1 Then nr = 1
    ReDim Preserve Minput(1 To nr) As String
    ReDim Preserve moutput(1 To nr) As String
End Sub

Function FindNextio(StartPos As Long, text As String) As Long
On Error GoTo fejl
Dim pos As Long, io As String
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
   
GoTo slut
fejl:
   FindNextio = 0
'    MsgBox "Fejl " & Err.Number & " (" & Err.Description & ") i procedure FindNextio, linje " & Erl & ".", vbOKOnly Or vbCritical Or vbSystemModal, TT.error
slut:
End Function
'#If Mac Then
Function RemovePrintetBool(s As String) As String
' removes the true/false which can be printed in output on Mac
' mac auto. A true or false may be printed at the end.
    Dim arr As Variant
    Dim t As String, ant As Integer, i As Integer
    On Error GoTo fejl
    s = TrimR(Trim$(s), vbLf)
    s = TrimR(Trim$(s), vbCr)
If DllConnType = 2 Then
    arr = Split(s, vbLf) ' wsh is a bit different
Else
    arr = Split(s, VbCrLfMac)
End If
    ant = UBound(arr)
    If ant > 0 Then
        t = Trim$(arr(ant))
        If t = "true" Or t = "false" Or t = "unknown" Or t = "expt: undefined: 0 to a negative exponent." Then
            s = vbNullString
            For i = 0 To ant - 1
                s = s & arr(i)
            Next
        End If
    End If
GoTo slut
fejl:
    MsgBox2 "removeprintetbool error", vbOKOnly, TT.Error
slut:
    RemovePrintetBool = s
End Function
'#End If
Function GetMaximaUnitOutput(Optional text As String) As String
    Dim Output As String
    Dim nr As Integer
    Dim AntalCifre As Integer
    Dim pos As Integer
    Dim pos2 As Integer
    Dim posb As Integer
    Dim ArrIndex As Integer
    ArrIndex = 1
    
    Dim MinputU() As String
    Dim moutputU() As String
    
    On Error GoTo fejl
    
#If Mac Then
'    Text = MaxProc.LastMaximaOutput
#Else
    If text = vbNullString Then
        text = MaxProcUnit.LastMaximaOutput
    End If
#End If
    ReDim MinputU(1 To 100) As String
    ReDim moutputU(1 To 100) As String
    Dim slutnu As Boolean
'    MsgBox text
   pos = FindNextio(1, text) - 1
   pos2 = FindNextio(pos + 2, text) - 1
    Do
        If pos < 1 Then
            Exit Function
        End If
        If pos2 < 1 Then
            pos2 = Len(text)
            slutnu = True
        End If
        posb = InStr(pos, text, ")")

        nr = CInt(Mid$(text, pos + 4, posb - pos - 4))
        If nr < 10 Then
            AntalCifre = 1
        ElseIf nr < 100 Then
            AntalCifre = 2
        ElseIf nr < 1000 Then
            AntalCifre = 3
        Else
            AntalCifre = 4
        End If
        Output = Replace(Trim$(Mid$(text, posb + 1, pos2 - pos - 4 - AntalCifre)), vbCrLf, "")

        If Mid$(text, pos + 3, 1) = "i" Then
            ' next is input
            MinputU(ArrIndex) = MinputU(ArrIndex) & Output
        Else
            ' next is output
            moutputU(ArrIndex) = moutputU(ArrIndex) & Output
            ArrIndex = ArrIndex + 1
        End If
        pos = pos2
        pos2 = FindNextio(pos + 2, text) - 1
    Loop Until slutnu
    nr = ArrIndex
    If moutputU(nr) = "" And Len(MinputU(nr)) < 6 Then
        nr = nr - 1
    End If

'    ReDim Preserve MinputU(1 To nr) As String
'    ReDim Preserve moutputU(1 To nr) As String

    '    ShowOutputArray ' til test
    GetMaximaUnitOutput = moutputU(nr)

    GoTo slut
fejl:
    '    MsgBox text
slut:
End Function
Sub ShowOutputArray()
' for test
    Dim i As Integer
    For i = 1 To UBound(Minput)
        MsgBox "Minput:" & i & vbCrLf & (Minput(i)) & vbCrLf & vbCrLf & "Moutput:" & vbCrLf & moutput(i)
    Next

End Sub
Function RemoveBrackets(t As String) As String
' removes all [] from the string
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
    If i > 1 Then GoTo slut
    tasksave.Activate

slut:
End Sub
Function FindVariable(Optional text As String, Optional MaximaCode As Boolean = True, Optional CAStype As Integer = 0) As Boolean
' finds variables in the text and puts them in Vars-string
    Dim Var As String
    Dim i As Integer
    Dim j As Integer
    Dim arrint As Variant
    AntalVars = 0
    Vars = ""
    IntegrationVars = ""
    If text = vbNullString Then
        text = Kommando
        If text = vbNullString Then
            On Error GoTo fejl
            For j = 0 To UBound(Kommandoer)
                text = text & "+" & Kommandoer(j)    ' there must only be characters between that cannot be included in the variable
            Next
        End If
    End If
    Dim ea As New ExpressionAnalyser
    If MaximaCode Then
        ea.text = CodeForMaxima(text, CAStype)
    Else
        ea.text = text
    End If
    If ea.text = vbNullString Then ' fejl
        GoTo fejl
    End If
    ea.pos = 1
    Dim vararr(30) As String    ' stores the variables for comparison only
    Dim varind As Integer
    Dim varexists As Boolean

' first add integration variables
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
    Var = ea.GetNextVar()
    Var = Replace(Var, vbCrLf, vbNullString)
    Do While Var <> vbNullString
        If ea.ChrByIndex(ea.pos + 1) <> "(" And ea.ChrByIndex(ea.pos + 1) <> "'" And ea.ChrByIndex(ea.pos + 1) <> ChrW$(8289) And Not (InStr(DefString, Var & ":") = 1 Or InStr(DefString, ";" & Var & ":") > 0) And InStr(Var, "qkqg") < 1 Then
            varexists = False
            For i = 1 To varind - 1
                If vararr(i) = Var Then varexists = True
            Next
            If Not (varexists) And Var <> "e" And Var <> "i" And Var <> "pi" And Var <> "inf" And Var <> "minf" Then
                If Vars <> "" Then
                    Vars = Vars & ";" & Var
                Else
                    Vars = Var
                End If
                vararr(varind) = Var
                varind = varind + 1
                AntalVars = AntalVars + 1
            End If
        End If
        ea.pos = ea.pos + 1
        Var = ea.GetNextVar()
    Loop
    
    FindVariable = True
    GoTo slut
fejl:
    FindVariable = False
slut:
End Function
Function FindDefinitions() As String
    Dim s As String
    Dim t As String
    Dim i As Integer
    Dim start As Long, sslut As Long
    Dim ra As Range

    On Error GoTo fejl

    Application.ScreenUpdating = False
    Definitions = False
    defindex = 0
    defstringtext = ""
    ReDim ldefname(100)
    ReDim ldefvalue(100)
    ReDim MathOArr(100)
    Set DefColl = New Collection
    start = Selection.Range.start
    sslut = Selection.Range.End
    Set ra = ActiveDocument.Range
    ra.End = start
    For i = ra.OMaths.Count To 1 Step -1
        s = ActiveDocument.Range.OMaths(i).Range.text
        If InStr(s, ChrW$(8788)) > 0 Or InStr(s, ChrW$(8797)) > 0 Or InStr(s, ":") > 0 Or InStr(s, ChrW$(8801)) > 0 Then    ' only look at fields with : triple with def= and especially := characters
            s = ReadEquation2(ActiveDocument.Range.OMaths(i).Range)
            If Not ValidateInput(s, ActiveDocument.Range.OMaths(i)) Then
                FindDefinitions = "error"
                Exit Function
            End If
            t = AddDefinition(s, ActiveDocument.Range.OMaths(i)) 'ra.OMaths(i)
            If t = "STOP" Then Exit For
            FindDefinitions = t & FindDefinitions
        End If
    Next
    If defindex > 0 Then
        ReDim Preserve ldefname(defindex - 1)
        ReDim Preserve ldefvalue(defindex - 1)
        ReDim Preserve MathOArr(defindex - 1)
    Else
        ReDim Preserve ldefname(0)
        ReDim Preserve ldefvalue(0)
        ReDim Preserve MathOArr(0)
    End If

    GoTo slut
fejl:
    DefFejl = True
    '    DefString = ""
    FindDefinitions = ""
slut:
    If FindDefinitions <> "" Then
        '            FindDefinitions = "[" & left$(FindDefinitions, Len(FindDefinitions) - 1) & "]$"
        Definitions = True
    End If
    DefString = FindDefinitions
    Selection.start = start
    Selection.End = sslut
    '    Selection.Collapse (wdCollapseStart)
End Function
Function AddDefinition(s As String, Optional MathObj As OMath) As String
    Dim t As String, slc As String
    Dim ea As New ExpressionAnalyser
    Dim ea2 As New ExpressionAnalyser
    Dim defenk As String
    Dim arr As Variant
    Dim deflign As Boolean

    ea.SetNormalBrackets
    If InStr(s, ChrW$(12310)) Then s = Replace(s, ChrW$(12310), "")    ' special invisible brackets are removed
    If InStr(s, ChrW$(12311)) Then s = Replace(s, ChrW$(12311), "")    ' special invisible brackets are removed
    If InStr(s, vbCrLf) Then s = Replace(s, vbCrLf, "")    ' enter can accur
    If InStr(s, vbCr) Then s = Replace(s, vbCr, "")
    slc = LCase$(s)
 
    If InStr(s, ChrW$(8788)) > 0 Or InStr(s, ChrW$(8797)) > 0 Or InStr(s, ":=") > 0 Or InStr(slc, "definer:") > 0 Or InStr(slc, "define:") > 0 Or InStr(slc, "ligning:") > 0 Or InStr(slc, "equation:") > 0 Or InStr(slc, "slet def") > 0 Or InStr(slc, "delete def") > 0 Or InStr(s, ChrW$(8801)) > 0 Or InStr(slc, LCase$(TT.A(69))) > 0 Then
        If LCase$(Left$(s, 8)) = "definer:" Then
            s = Right$(s, Len(s) - 8)
        ElseIf LCase$(Left$(s, 7)) = "define:" Then
            s = Right$(s, Len(s) - 7)
        ElseIf LCase$(Left$(s, 16)) = "definer ligning:" Then
            s = Right$(s, Len(s) - 16)
            deflign = True
        ElseIf LCase$(Left$(s, 16)) = "define equation:" Then
            s = Right$(s, Len(s) - 16)
            deflign = True
        End If
        If InStr(s, ChrW$(8788)) Then s = Replace(s, ChrW$(8788), ":=")    ' \equiv and
        If InStr(s, ChrW$(8797)) Then s = Replace(s, ChrW$(8797), ":=")   ' def=
        If InStr(s, ChrW$(8801)) Then s = Replace(s, ChrW$(8801), ":=")    ' speciel := symbol

        s = Replace(s, ChrW$(11), "")
       slc = LCase$(s)
 
        '        If InStr(s, "block") < 1 Then
        '            s = Replace(s, chrw$(34), "") ' apostrophe
        '            s = Replace(s, "=", ":")
        '            s = Replace(s, "::", ":")
        '            s = Replace(s, "):", "):=")
        '        End If

        '        s = Replace(s, "):=:", "):=")
        If InStr(slc, "slet def") > 0 Or InStr(slc, "delete def") > 0 Or InStr(slc, LCase$(TT.A(69))) Then
            s = Replace(Trim$(Split(s, ":")(1)), vbCr, "")
            If s = "" Then
                AddDefinition = "STOP"
                Exit Function
            Else
                '                AddDefinition = "kill(" & Replace(Replace(s, ";", ","), vbCr, "") & ")$" & AddDefinition
                AddDefinition = "errcatch(kill(" & CodeForMaxima(Replace(s, ",", ";"), CASengine) & "))$" & AddDefinition
            End If
        Else    ' gem definitioner
            '            FindDefinitions = CodeForMaxima(s) & "," & FindDefinitions
            defstringtext = defstringtext & s & ";"
            ea.text = Replace(s, ";", " , ")
            ea.ConvertDecSeparator
            ea.pos = 1
            Do
                defenk = Trim$(ea.GetNextListItem(ea.pos, ","))
                If Len(defenk) = 0 Then GoTo hop
                If InStr(defenk, "block") > 0 Then
                    defenk = Replace(defenk, ";", ",")
                    defenk = Replace(defenk, ChrW$(9633), "")    ' characters before parentheses that make them invisible. Possibly problematic as the parentheses are not removed. also makes fractions small
                    defenk = Replace(defenk, ChrW$(9508), "")    ' blank space character at least [ ]
                    defenk = Replace(defenk, ChrW$(11), "")    ' shift-enter
                    defenk = Replace(defenk, vbLf, "")    ' shift-enter and enter
                    defenk = Replace(defenk, vbCrLf, "")
                    defenk = Replace(defenk, vbCr, "")
                    defenk = Replace(defenk, ChrW$(183), "*")    ' dot is replaced with times
                    defenk = Replace(defenk, ChrW$(8901), "*")    ' \cdot , don't know why there seem to be two
                    defenk = Replace(defenk, ChrW$(8729), "*")    ' \cdot replaced by multiply
                    defenk = Replace(defenk, ChrW$(8226), "*")    ' thick dot replaced with times
                    defenk = Replace(defenk, ChrW$(247), "/")
                    defenk = Replace(defenk, ChrW$(8800), "#")    ' not equal to
                    defenk = Replace(defenk, ChrW$(8804), "<=")
                    defenk = Replace(defenk, ChrW$(8805), ">=")
                    defenk = Replace(defenk, ChrW$(160), "")    ' before accent e.g. vector arrow 160 8407
                    defenk = Replace(defenk, ChrW$(8289), "") ' function symbol
                    defenk = Replace(defenk, ChrW$(9618), "") ' function symbol
                    defenk = Replace(defenk, ChrW$(12310), "(") ' hidden brackets
                    defenk = Replace(defenk, ChrW$(12311), ")") ' hidden brackets
                Else
                    defenk = Replace(defenk, ChrW$(34), vbNullString)    ' apostrof
                    defenk = CodeForMaxima(defenk, CASengine)
                    If InStr(defenk, "if") <= 0 And (InStr(defenk, "<") > 0 Or InStr(defenk, ">") > 0) Or (InStr(defenk, "#") > 0 And InStr(defenk, "##") <= 0) Then
                        AddDefinition = AddDefinition & GetAssumeText(defenk)
                        If UBound(ldefname) >= defindex Then
                            ldefname(defindex) = defenk
                            ldefvalue(defindex) = ""
                            Set MathOArr(defindex) = MathObj
                            defindex = defindex + 1
                        End If
                        GoTo hop
                    End If
                    If Not deflign Then
                        defenk = Replace(defenk, "=", ":", 1, 1)
                        defenk = Replace(defenk, "::", ":", 1, 1)
                        defenk = Replace(defenk, "):", "):=", 1, 1)
                    End If
                End If
                arr = Split(defenk, ":")
                If UBound(arr) = 0 Then
                    '                    If InStr(defenk, "setunits") > 0 Or InStr(defenk, "assume") > 0 Then
                    GoTo hop2
                    '                    End If
                End If
                If Not ValidateDef(arr(0), arr(1)) Then GoTo hop

                If Left$(arr(1), 1) = "=" Then arr(1) = Right$(arr(1), Len(arr(1)) - 1)
                If arr(1) = "" Then GoTo hop
            
                If UBound(ldefname) < defindex Then
                    ReDim Preserve ldefname(defindex)
                    ReDim Preserve ldefvalue(defindex)
                    ReDim Preserve MathOArr(defindex)
                End If
                ldefname(defindex) = arr(0)
                ldefvalue(defindex) = arr(1)
                Set MathOArr(defindex) = MathObj
                defindex = defindex + 1
hop2:                    ' if setunit or assume or declare
                ea2.text = defenk
                ea2.pos = 1
                t = ea2.GetNextVar

                If InStr(Split(defenk, ":")(0), "diff") > 0 Then
                    MsgBox2 TT.A(116) & s, vbOKOnly, TT.Error
                ElseIf (t = "e" And InStr(defenk, "[") <= 0) Or t = "integrate" Or t = "sin" Or t = "cos" Or t = "tan" Or t = "diff" Or t = "solve" Or t = "at" Or t = "for" Or t = "in" Or t = "do" Or t = "if" Or t = "and" Or t = "or" Or t = "product" Or t = "step" Or t = "next" Or t = "from" Or t = "limit" Or t = "sum" Or t = "then" Or t = "and" Or t = "or" Or t = "else" Or t = "unless" Or t = "while" Or t = "vecmag" Then
                    MsgBox2 TT.A(845) & " " & t & " " & TT.A(434), vbOKOnly, TT.Error
                ElseIf t = "setunits" Or t = "assume" Or t = "declare" Then
                    '                    AddDefinition =  AddDefinition & CodeForMaxima(defenk) & "$"
                    AddDefinition = AddDefinition & defenk & ";"
                ElseIf t <> "" Then
                    KillDef = KillDef & t & ","
                    AddDefinition = AddDefinition & defenk & ";"
                End If
hop:
            Loop Until ea.pos >= Len(ea.text)
        End If
    End If

End Function
Function SplitAssump(s As String) As String
' splits  0<x<5  to  x>0 and x<5
   Dim ea As New ExpressionAnalyser
   Dim v As String, p As Integer
   Dim pre As String, pst As String
   
   ea.SetNormalBrackets
   
    s = Replace(s, ChrW$(8804), "<=")    '
    s = Replace(s, ChrW$(8805), ">=")    '
    s = Replace(s, " ", vbNullString)
   ea.text = s
   
   v = ea.GetNextVar()
   If v = "" Then
      SplitAssump = s
      Exit Function
   End If
   
   p = InStr(s, "<=" & v & "<=")
   If p > 0 Then
      pre = Left$(s, p - 1)
      pst = Right$(s, Len(s) - p - 4)
      SplitAssump = v & ">=" & pre & "&&and&&" & v & "<=" & pst
'      SplitAssump = pre & "<=" & v & " and " & v & "<=" & pst
      Exit Function
   End If
   p = InStr(s, "<=" & v & "<")
   If p > 0 Then
      pre = Left$(s, p - 1)
      pst = Right$(s, Len(s) - p - 3)
      SplitAssump = pre & "<=" & v & "&&and&&" & v & "<" & pst
      Exit Function
   End If
   p = InStr(s, "<" & v & "<=")
   If p > 0 Then
      pre = Left$(s, p - 1)
      pst = Right$(s, Len(s) - p - 3)
      SplitAssump = pre & "<" & v & "&&and&&" & v & "<=" & pst
      Exit Function
   End If
   p = InStr(s, "<" & v & "<")
   If p > 0 Then
      pre = Left$(s, p - 1)
      pst = Right$(s, Len(s) - p - 2)
      SplitAssump = pre & "<" & v & "&&and&&" & v & "<" & pst
      Exit Function
   End If
   SplitAssump = s
End Function
Function sqbrAssump(s As String) As String
   Dim ea As ExpressionAnalyser, t As String
   Dim v As String, p As Integer, c As String
   Dim arr() As String, l1 As String, l2 As String
   t = s
   p = InStr(s, ChrW$(8712))
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
   s = Right$(s, Len(s) - p)
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
      sqbrAssump = v & l1 & arr(0) & "&&and&&" & arr(1) & l2 & v
   Else
      sqbrAssump = t
   End If

End Function
Sub ResetDefinitions()
    ReDim ldefname(100)
    ReDim ldefvalue(100)
    ReDim MathOArr(100)
    defindex = 0
End Sub
Function ValidateDef(ByVal LHS As String, ByVal RHS As String) As Boolean
' returnerer false hvis der er fejl
    Dim pos As Integer
    On Error GoTo fejl
    ValidateDef = False

    pos = InStr(LHS, "[")    ' You can write anything in subscript/index
    If pos > 0 Then
        LHS = Left$(LHS, pos - 1)
    End If

    ' If Not ea.IsAllText Then   not good enough then, and; not taken into account
    If InStr(LHS, "+") > 0 Or InStr(LHS, "-") > 0 Or InStr(LHS, "*") > 0 Or InStr(LHS, "/") > 0 Then
        MsgBox2 TT.A(117) & vbCrLf & vbTab & ConvertToAscii(LHS & "=" & RHS) & vbCrLf & TT.A(118), vbOKOnly, TT.Error
        Exit Function
    End If

    If Right$(RHS, 1) = "+" > 0 Then
        MsgBox2 TT.A(117) & vbCrLf & vbTab & ConvertToAscii(LHS & "=" & RHS) & vbCrLf & TT.A(118), vbOKOnly, TT.Error
        Exit Function
    End If


    ValidateDef = True
fejl:
End Function
Function GetAssumeText(defenk) As String
' takes text from word and converts to assume and inserts into forget
    Dim ea As ExpressionAnalyser
    Dim ea2 As ExpressionAnalyser
    Dim s As String
    Dim arr As Variant
    Dim p1 As Integer, p2 As Integer
    If Len(defenk) > 0 Then
        If Left$(defenk, 6) = "assume" Then
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
                s = ea.GetNextListItem(ea.pos)
                If InStr(s, "#") > 0 Then
                    arr = Split(s, "#")
                    s = "notequal(" & arr(0) & "," & arr(1) & ")"
                ElseIf Len(s) > 0 Then
                    ea2.text = s
                    p1 = ea2.FindNextGreaterOrLower(1)
                    p2 = ea2.FindNextGreaterOrLower(p1 + 1)
                    If p2 > 0 Then    ' 2<=x<=3
                        If CASengine = 0 Then
                            If ea.ChrByIndex(p1 + 1) = "=" Then
                                s = Left$(s, p2 - 1) & "," & Right$(s, Len(s) - p1 - 1)
                            Else
                                s = Left$(s, p2 - 1) & "," & Right$(s, Len(s) - p1)
                            End If
                        End If
                    End If
                End If
                If Len(s) > 0 Then defenk = defenk & s & ","
            Loop While s <> ""
            defenk = Left$(defenk, Len(defenk) - 1)
            ForgetList = ForgetList & defenk & ","
            GetAssumeText = "assume(" & defenk & ")$"
        End If
    End If

End Function

Sub AddToKillDefs(DefS As String)
    Dim arr As Variant
    Dim j As Integer
    Dim s As String
    If DefS <> "" Then
        arr = Split(DefS, ListSeparator)
        For j = 0 To UBound(arr)
            s = arr(j)
            s = Split(s, "=")(0)
            s = Split(s, "(")(0)
            KillDef = KillDef & s & ","
        Next
    End If
End Sub

Function IsSolved(ByVal Ligning As String, variabel As String) As Boolean
' checks if the equation is completely solved for the variable
' starts it with variable= ?
' includes variable on the right side of = sign
    On Error GoTo slut
    Dim ea As New ExpressionAnalyser
    Dim Var As String
    Dim arr As Variant
    If CASengine > 0 Then
        ea.SetTuborgBrackets
    End If
    Ligning = RemoveBrackets(Ligning)
    arr = Split(Ligning, "=")

    If UBound(arr) = 1 Then    ' there is exactly one = sign
        ea.text = arr(1)    'right side
        IsSolved = True
        Do
            Var = ea.GetNextVar
            ea.pos = ea.pos + 1
            If Var = variabel Then
                IsSolved = False    'variable found on the right hand side then it is not solved anyway
                GoTo slut
            End If
        Loop Until Var = ""

        If Left$(Ligning, 1 + Len(variabel)) = variabel & "=" Then  '  not enough alone
            IsSolved = True
        ElseIf Len(variabel) > 1 And Mid$(Ligning, Len(variabel) + 1, 1) = "=" Then    ' if length of variable same as left side. Necessary for variables like 1+r as it can be changed to r+1. But probably not 100% sure
            IsSolved = True
        Else    ' it must be assumed that it is not solved if it does not match anything above
            IsSolved = False
        End If
    Else
        IsSolved = False
    End If

    GoTo slut
fejl:
    IsSolved = False
slut:
End Function

Function IsAllSolved(ByVal ligninger As String, ByVal variabel As String, Optional ByVal Sep As String = ",") As String
' checks if a list of equations is completely solved
' true if one is solved, false if none. All if all.
    Dim oneSolved As Boolean
    Dim allsolved As Boolean
    Dim arr As Variant
    Dim i As Integer
    oneSolved = False
    allsolved = True
    If ligninger = "" Then
        IsAllSolved = "false"
        Exit Function
    End If

    If Not LmSet Then
        If InStr(ligninger, "%if(") > 0 Then GoTo slut

        ligninger = RemoveBrackets(ligninger)
        arr = Split(ligninger, Sep)

        For i = 0 To UBound(arr)
            If IsSolved(arr(i), variabel) Then
                oneSolved = True
            Else
                allsolved = False
            End If
        Next
        
        If InStr(ligninger, ChrW$(8743)) > 0 Then ' there may be systems of equations if vector equations have been submitted
            arr = Split(ligninger, ChrW$(8743))
            For i = 0 To UBound(arr)
                If IsSolved(arr(i), variabel) Then
                    oneSolved = True
                Else
                    allsolved = False
                End If
            Next
        End If
        
    End If
slut:
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
    On Error GoTo fejl

    If Kommando <> "" Then
        KommandoerStreng = Replace(Kommando, vbCr, "")
    Else
        For i = 0 To UBound(Kommandoer)
            KommandoerStreng = KommandoerStreng & Kommandoer(i) & ListSeparator
        Next
        KommandoerStreng = Left$(KommandoerStreng, Len(KommandoerStreng) - 1)
    End If

    GoTo slut
fejl:
    KommandoerStreng = ""
slut:
End Property

Public Property Get AntalKom() As Integer
' number of equations/expressions
    On Error GoTo fejl
    AntalKom = UBound(Kommandoer) + 1
    GoTo slut
fejl:
    If Kommando <> vbNullString Then
        AntalKom = 1
    Else
        AntalKom = 0
    End If
slut:
End Property

Public Property Get DefName(Index As Integer) As String
    On Error Resume Next
    DefName = ldefname(Index)
End Property

Public Property Get DefValue(Index As Integer) As String
    On Error Resume Next
    DefValue = ldefvalue(Index)
End Property
Public Property Get DefOMath(Index As Integer) As OMath
    On Error Resume Next
    DefOMath = MathOArr(Index)
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

Function GetMaximaSettingsString() As String
'dll API makes these settings itself

    Dim SettingsString As String, arr() As String, i As Integer
    
    SettingsString = "["
    If MaximaUnits Then
        SettingsString = SettingsString & "keepfloat:false,setunits([N,J,W,Pa,C,V,F,Ohm,T,H,K]),"
        If Not OutUnits = vbNullString Then
            arr = Split(OutUnits, ",")
            For i = 0 To UBound(arr) ' It has to be done this way because you cannot run setunits on a list where the same type of unit appears twice. e.g. eV and J
                SettingsString = SettingsString & "setunits(" & omax.ConvertUnits(arr(i)) & "),"
            Next
        End If
    Else
        SettingsString = SettingsString & "keepfloat:true,"
    End If
    
    If MaximaComplex Then
        SettingsString = SettingsString & "domain:complex,realonly:false,"
    Else
        SettingsString = SettingsString & "domain:real,"
    End If
    
'    If MaximaExact = 2 Then
'        SettingsString = SettingsString & "dectalall(udt):=dectalallNum(udt),"
'    Else
'        SettingsString = SettingsString & "dectalall(udt):=block([qout],qout:errcatch(apply1(udt,floatrul)),if qout=[] then udt else qout[1]),"
'    End If
    
    SettingsString = SettingsString & "DecOutType:" & MaximaDecOutType - 1 & ","
    SettingsString = SettingsString & "NoSigFig:" & MaximaCifre & ","
    SettingsString = SettingsString & "fpprintprec:" & MaximaCifre & ","
            
    If Right$(SettingsString, 1) = "," Then SettingsString = Left$(SettingsString, Len(SettingsString) - 1)
    SettingsString = SettingsString & "]$"
     
    GetMaximaSettingsString = SettingsString
End Function
