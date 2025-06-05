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
'            MaximaInputStreng = MaximaInputStreng & "[" & Left(Replace(DefS, "$", ","), Len(DefS) - 1) & "]$"
            MaximaInputStreng = Replace(DefS, "$", ";")
        End If
        MaximaInputStreng = MaximaInputStreng & "fejl;"
    End If

    PrepareNewCommand = True
    GoTo slut

Fejl:
    MsgBox TT.ErrorGeneral & vbCrLf & " preparing maxima", vbOKOnly, TT.Error
    PrepareNewCommand = False
slut:

End Function
Private Sub Class_Terminate()
'Closes commandprompt
    On Error Resume Next
End Sub

Public Sub ResetMaxima()
    Dim kd As String
    
    If DllConnType = 2 Then Exit Sub ' wsh starts anew every time so reset is not necessary
    
#If Mac Then
    KillDef = ""
    ForgetList = ""
    Exit Sub
#Else
    If Len(KillDef) > 0 Then
        KillDef = Left(KillDef, Len(KillDef) - 1)    'remove last comma
        kd = kd & "errcatch(kill(" & KillDef & "))"
        KillDef = ""
    End If
    If Len(ForgetList) > 0 Then
        ForgetList = Left(ForgetList, Len(ForgetList) - 1)    'remove last comma
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
    Set UFwait2 = New UserFormWaitForMaxima

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
            RestartMaxima
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
            Unload UFwait2
            MaxProc.CloseProcess
'            MaxProc.Finished = 1
            MaxProc.StartMaximaProcess
            Exit Sub
        End If

        MaxProc.ExecuteMaximaCommand InputStreng, 0
        UFwait2.Label_progress.Caption = UFwait2.Label_progress.Caption & "*"
    End If
    
    WaitForMaximaForm StopTime

    AnswerQuestion StopTime

    ShowDebug
    '    MsgBox (MaxProc.LastMaximaOutput)

    If Reset Then
        FindLastMaximaOutput ConvertWordSymbols
        ResetMaxima
    End If
    On Error Resume Next
    Unload UFwait2
    GoTo slut
Fejl:
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
    Dim i As Integer, k As String, Arr As Variant, s As String
    
    Expr = Replace(Expr, """", "") ' The answer will typically be run through the num function which adds "" to numbers
    
    If basedim Then ' all units are reverted to base SI-units
        k = Expr & ",numer;applyb1(%,unitrule),numer;dectalallNum(%),numer;"
    Else
        k = "numer:true$uforget(append(globalbaseunitlisting,globalderivedunitlisting))$"
        If Not OutUnits = vbNullString Then
            Arr = Split(OutUnits, ",")
            For i = 0 To UBound(Arr) ' It has to be done this way because you cannot run setunits on a list where the same type of unit appears twice. e.g. eV and J
                k = k & "setunits(" & omax.ConvertUnits(Arr(i)) & ")$"
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
        ConvertMaximaUnits = Replace(ConvertMaximaUnits, VBA.ChrW(215), "~")    ' vector produkt
'        ConvertMaximaUnits = Left(ConvertMaximaUnits, Pos - 1) & "vecmag(" & Mid(ea.text, Pos + 1, posb - Pos - 1) & ")" & right(ea.text, Len(ea.text) - posb)
        ConvertMaximaUnits = Replace(ConvertMaximaUnits, "^", "^^")
        ConvertMaximaUnits = Replace(ConvertMaximaUnits, "*", " ## ")
        Dim ea As New ExpressionAnalyser
        ea.Text = ConvertMaximaUnits
        ea.ReplaceVar "abs", "vecmag"
        ConvertMaximaUnits = ea.Text
    End If

End Function
Sub WaitForMaximaForm(StopTime As Integer)
#If Mac Then
#Else
    Dim i As Integer
    Dim UFwait2 As New UserFormWaitForMaxima
    
    i = 0
    Do While MaxProc.Finished = 0 And i < 11
        Wait (0.1)
        i = i + 1
    Loop
    If MaxProc.Finished = 0 Then
        '    ufwait2.omax = omax
        On Error Resume Next
        UFwait2.Show vbModeless
        DoEvents
        On Error GoTo Fejl
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
Fejl:
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
        
        UserFormDebug.Label_time.Caption = UserFormDebug.Label_time.Caption & "Samlet tid: " & timer - tid & vbCrLf
        UserFormDebug.Show
    End If
End Sub
#If Mac Then
#Else
Sub AnswerQuestion(StopTime As Integer, Optional ShowWaitForm As Boolean = True)
    Dim svar As String
    Dim answered As Boolean

    Do While MaxProc.Question = 1
        Dim Pos As Integer
        Dim pos2 As Integer
        Pos = InStr(MaxProc.LastMaximaOutput, "positive, negative or zero?")
        pos2 = InStr(MaxProc.LastMaximaOutput, "positive or negative?")
        If Pos > 0 Or pos2 > 0 Then
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
            svar = InputBox(MaxProc.LastMaximaOutput, "spr", "")    ' MaxProc.Question must be adjusted
        ElseIf InStr(MaxProc.LastMaximaOutput, "variables to solve for?") Then
            Dim antalvar As Integer, Var As String, ea As New ExpressionAnalyser, sprtext As String
            Pos = InStr(MaxProc.LastMaximaOutput, " variables to solve for?")
            antalvar = val(Mid(MaxProc.LastMaximaOutput, Pos - 2, 2))
            sprtext = TT.A(416) & " " & antalvar & " " & TT.A(417)
            ea.Text = Right(MaxProc.LastMaximaOutput, Len(MaxProc.LastMaximaOutput) - Pos - 24)
            ea.SetSquareBrackets
            ea.Pos = 0
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
        If Right(MaximaInputStrengSec, 1) = ";" Then    ' at answerquestion is set automatically;
            MaximaInputStrengSec = Left(MaximaInputStrengSec, Len(MaximaInputStrengSec) - 1)
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
    Dim Pos As Integer, pos2 As Integer, Assume As String, InputString As String

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
    If Left(s, 5) = "Fejln" Then
        WT = InputBox(TT.A(892) & WT & "s." & TT.A(893), "Stopped", 20)
        If Trim(WT) <> "" Then
            WT = Trim(WT)
            GoTo start
        Else
            StopNow = True
        End If
    End If
    If Assume = vbNullString Then
        Pos = InStr(s, "positive, negative or zero?")
        pos2 = InStr(s, "positive or negative?")
    
        If Pos > 0 Or pos2 > 0 Then
            Dim UF As New UserFormAskSign
            If Pos <= 0 And pos2 > 0 Then Pos = pos2
            pos2 = InStr(s, "Is ")
            Var = Mid(s, pos2 + 3, Pos - pos2 - 4)
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
            '                Var = Mid(s, p + 3, p2 - p - 3)
            '            Else
            '                Var = "x"
            '            End If
            '        Else
            '            Var = "x"
            '        End If
            '        MsgBox "The problem could not be solved, because the sign of a variable is unknown. You can probably solve the problem by inserting a definition. Example:" & vbCr & " 'Define: " & Var & ">0'" & vbCrLf & "Maxima Asked:" & vbCrLf & right(s, Len(s) - p)
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
        KillDef = Left(KillDef, Len(KillDef) - 1)    'remove last comma
        MaximaInputStreng = MaximaInputStreng & "errcatch(kill(" & KillDef & "))$"
        KillDef = ""
    End If
    If Len(ForgetList) > 0 Then
        ForgetList = Left(ForgetList, Len(ForgetList) - 1)    'remove last comma
        MaximaInputStreng = MaximaInputStreng & "forget(" & ForgetList & ")$"
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
    Dim Text As String
    If MaxCmd <> "" Then
        Kommando = MaxCmd    ' VBA.LCase maybe
    End If

    Application.ScreenUpdating = False
    If TranslateToMaxima Then
        MaximaInputStreng = MaximaInputStreng & CodeForMaxima(Kommando) & ";"
    Else
        Text = Kommando
        Text = Replace(Text, VBA.ChrW(9633), "")    ' characters before parentheses that make them invisible. Possibly problematic as the parentheses are not removed. also makes fractions small
        Text = Replace(Text, VBA.ChrW(9508), "")    ' blank space character at least [ ]
        Text = Replace(Text, VBA.ChrW(11), "")    ' shift-enter
        Text = Replace(Text, vbLf, "")    ' shift-enter and enter
        Text = Replace(Text, vbCrLf, "")
        Text = Replace(Text, vbCr, "")
        Text = Replace(Text, VBA.ChrW(183), "*")    ' dot is replaced with times
        Text = Replace(Text, VBA.ChrW(8901), "*")    ' \cdot , don't know why there seem to be two
        Text = Replace(Text, VBA.ChrW(8729), "*")    '  \cdot replaced with times
        Text = Replace(Text, VBA.ChrW(8226), "*")    ' thick dot replaced with times
        Text = Replace(Text, VBA.ChrW(247), "/")    '
        Text = Replace(Text, VBA.ChrW(8800), "#")    ' not equal to
        Text = Replace(Text, VBA.ChrW(8804), "<=")    '
        Text = Replace(Text, VBA.ChrW(8805), ">=")    '
        Text = Replace(Text, VBA.ChrW(160), "")    ' before accent e.g. vector arrow 160 8407
        Text = Replace(Text, VBA.ChrW(8289), "") ' function symbol
        Text = Replace(Text, VBA.ChrW(9618), "") ' function symbol
        Text = Replace(Text, VBA.ChrW(12310), "(") ' hidden brackets
        Text = Replace(Text, VBA.ChrW(12311), ")") ' hidden brackets
        
        Kommando = Text
        Kommando = Replace(Kommando, VBA.ChrW(8289), "")
        MaximaInputStreng = MaximaInputStreng & Kommando & ";"
    End If
    Call RunMaxima
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
        If (InStr(ikommando, "log" & VBA.ChrW(8289)) > 0 Or InStr(ikommando, "10" & VBA.ChrW(12311) & "^") > 0) Then
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
            MaximaOutput = "L=" & VBA.ChrW(8709)
        Else
            MaximaOutput = variabel & VBA.ChrW(8712) & VBA.ChrW(8709)
        End If
    ElseIf MaximaOutput = "all" Or MaximaOutput = "[all]" Or MaximaOutput = "[" & variabel & "=k_1]" Then
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
                MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8744) & "    ")    ' Comma replaced by or
            ElseIf ListSeparator = "," And InStr(MaximaOutput, ",") Then
                MaximaOutput = RemoveBrackets(MaximaOutput)
                MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8744) & "    ")    ' Comma replaced by or
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
    Dim Arr As Variant
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
            MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8744) & "    " & variabel & VBA.ChrW(&H2248))    ' Comma replaced by or
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

    MaximaOutput = Trim(MaximaOutput)
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
            MaximaOutput = variabel & VBA.ChrW(&H2248) & RemoveBrackets(MaximaOutput)
        End If
        MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8744) & "    " & variabel & VBA.ChrW(&H2248))    ' Comma replaced by or
    End If

slut:
End Sub
Sub Nsolve(ByVal variabel As String, ByVal qn As Integer, ByVal qm As Integer, ByVal qtid As Integer, ByVal maxsol As Integer, Optional ByVal qnn As Integer = 30, Optional ByVal qmm As Integer = 30, Optional newt As Boolean = False)
    If variabel = "" Then Exit Sub
    '    eps = "10^-" & MaximaCifre
    If TempDefs <> "" Then
        MaximaInputStreng = AddDefinition("definer:" & TempDefs) & MaximaInputStreng
    End If

    Kommando = "nsolve(" & CodeForMaxima(Kommando) & "," & variabel & "," & qn & "," & qm & "," & qtid & "," & maxsol & "," & qnn & "," & qmm & "," & VBA.LCase(newt) & ")"
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
            MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8744) & "    ")    ' Comma replaced by or
            MaximaOutput = Replace(MaximaOutput, "=", VBA.ChrW(&H2248))    ' cirka equal
            '    If MaximaOutput <> "" Then
            '        MaximaOutput = variabel & VBA.ChrW(&H2248) & RemoveBrackets(MaximaOutput)
            '    End If
            '    MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8744) & "    " & variabel & VBA.ChrW(&H2248)) ' Comma replaced by or
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
        If (InStr(ikommando, "log" & VBA.ChrW(8289)) > 0 Or InStr(ikommando, "10" & VBA.ChrW(12311) & "^") > 0) Then
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
    Dim Arr As Variant
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
    Kommando = Trim(Kommando)

    If Right(Kommando, 1) = "=" Then
        Kommando = Left(Kommando, Len(Kommando) - 1)
    End If

start:
    StartKommando = Kommando    ' command before it is tampered with
    StartKommando3 = Kommando
    
    pkommando = Kommando    ' command before it is tampered with
    pkommando = Replace(pkommando, VBA.ChrW(12310), "") ' special invisible brackets are removed
    pkommando = Replace(pkommando, VBA.ChrW(12311), "")
    Kommando = CodeForMaxima(Kommando)
    If Kommando = vbNullString And ConvertErrorText <> vbNullString Then
        GoTo slut
    End If
    inkom = "%"    ' save command

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
            eql = VBA.ChrW(&H2248)
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
            eqsg = VBA.ChrW(&H2248)
        End If
        If moutput(nr) <> "" Then
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
        If o = "" Then Exit Sub    ' only 1 solution

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
' for geogebra output for systems of equations
    Dim p As Integer, s As String, i As Integer, j As Integer
    Dim ea As New ExpressionAnalyser
    Dim Arr() As String, Arr2() As String
    
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
        Arr2 = Split(Arr(i), "&")
        If UBound(Arr2) > 0 Then
            If UBound(Arr) > 0 Then MaximaOutput = MaximaOutput & "("
            For j = 0 To UBound(Arr2)
                MaximaOutput = MaximaOutput & Arr2(j) & "    " & VBA.ChrW(8743) & "    "
            Next
            MaximaOutput = Left(MaximaOutput, Len(MaximaOutput) - 9)
            If UBound(Arr) > 0 Then MaximaOutput = MaximaOutput & ")"
        End If
        MaximaOutput = MaximaOutput & "    " & VBA.ChrW(8744) & "    "
    Next
    If Len(MaximaOutput) > 9 Then MaximaOutput = Left(MaximaOutput, Len(MaximaOutput) - 9)
    
End Sub
Sub ConvertOutputToSolved2(Optional Variable As String)
' converts output from [[x=1,y=3],[x=3,y=5]] to (x=1 and x=3) v ... or L={(1,3),(3,5)}
' intended for symbolic solution and mnewton and SolveIneq only
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
            ' notation L={}
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

        Else    ' notation x=...

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
            If o = "" Then Exit Sub    ' only 1 solution

            If InStr(MaximaOutput, VBA.ChrW(8743)) > 0 Then    ' only parentheses implication sign
                MaximaOutput = "(" & MaximaOutput & ")"
            End If

            o = Replace(o, ListSeparator, "    " & VBA.ChrW(8743) & "    ")
            If InStr(o, VBA.ChrW(8743)) > 0 Then    ' only parentheses implication sign
                o = "(" & o & ")"
            End If

            MaximaOutput = MaximaOutput & "   " & VBA.ChrW(8744) & "   " & o

            Do
                o = ea.GetNextBracketContent
                If o = "" Then Exit Do
                o = Replace(o, ListSeparator, "    " & VBA.ChrW(8743) & "    ")
                If InStr(o, VBA.ChrW(8743)) > 0 Then    ' only parentheses implication sign
                    o = "(" & o & ")"
                End If
                MaximaOutput = MaximaOutput & "   " & VBA.ChrW(8744) & "   " & o
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
        ea.Text = Kommando
        ea.ReplaceVar Split(funktion, "(")(0), "qvar"
        Kommando = ea.Text
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
            MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8744) & "    ")    ' Comma replaced by or
        End If
    End If
slut:
End Sub
Sub Plot2D(forskrifter As String, punkter As String, Var As String, xmin As String, xmax As String, Ymin As String, Ymax As String, xaksetitel As String, yaksetitel As String, Optional SaveFile As Boolean = False)
' functions separated by commas
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
        Text = Text & ",key=""" & ConvertToAscii(Arr(i)) & """,color=" & GetNextColor() & ",explicit(" & fkt & "," & Var & "," & xmin & "," & xmax & ")"
    Next
    If punkter <> "" Then
        Text = Text & ",key=""punkter"",point_type=filled_circle,point_size=1,points(" & punkter & ")"
    End If

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
    Dim WaitTime As Integer
    If df = "" Then
#If Mac Then
'        text = "gnuplot_command:""/Applications/WordMat/gnuplot/gnuplot""$set_plot_option([gnuplot_term, aqua])$draw2d("
    Text = "gnuplot_command:""/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/MaximaWM/maxima/bin/gnuplot""$set_plot_option([gnuplot_term, aqua])$draw2d("
'    text = "gnuplot_command:""/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/gnuplot/gnuplot""$set_plot_option([gnuplot_term, aqua])$draw2d("
'        waittime = -20
#Else
        Text = "draw2d("
#End If
        WaitTime = -400
    Else
#If Mac Then
    Text = "gnuplot_command:""/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/MaximaWM/maxima/bin/gnuplot""$set_plot_option([gnuplot_term, aqua])$load(drawdf)$drawdf(" & df & ","
'text = "gnuplot_command:""/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/gnuplot/gnuplot""$set_plot_option([gnuplot_term, aqua])$load(drawdf)$drawdf(" & df & ","
'        waittime = -25
#Else
        Text = "load(drawdf)$drawdf(" & df & ","
#End If
        WaitTime = -600
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
            '    text = text & ",user_preamble = ""set arrow from graph 0.95, first 0 to graph 1, first 0""" ' arrow on x-axis not nice
            '            text = text & ",user_preamble = ""set grid lt -1 """ ' black lines solid
            Text = Text & ",user_preamble = ""set grid lt 0 lc -1 lw 1"""    ' black dotted lines
            '           text = text & ",user_preamble = ""set grid lt 0 lw " & resolution & """" ' linewidth>1 does not display correctly on horizontal lines
            ' lw 2 is linewidth 2, lt is linetype, lc is linecolor
        End If
    End If
    If xaksetitel <> "" Then
        Text = Text & ",xlabel=""" & xaksetitel & """"
    End If
    If yaksetitel <> "" Then
        Text = Text & ",ylabel=""" & yaksetitel & """"
    End If

    Text = Text & "),numer;" ' numer introduced because there was an expression with numerical integral, where it just waited without. There is no need for exact expression when plotting.
    MaximaInputStreng = MaximaInputStreng & Text
    
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
Sub PlotDF(lign As String, IndepVar As String, DepVar As String, xmin As String, xmax As String, Ymin As String, Ymax As String, xtrajec As String, ytrajec As String, parametre As String)
'plot of curve integrals
    Dim Text As String

    Text = "plotdf("
    Text = Text & lign

    If xmin <> "" And xmax <> "" Then
        Text = Text & ",[" & IndepVar & "," & DepVar & "],[" & IndepVar & "," & xmin & "," & xmax & "]"
    End If
    If Ymin <> "" And Ymax <> "" Then
        Text = Text & ",[" & DepVar & "," & Ymin & "," & Ymax & "]"
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

Sub Draw3D(graphobj As String, Optional antalobj As Integer = 1)
'    text = "load(draw)$draw3d("
    Dim Text As String
    Dim tid As Integer

    tid = 20 + antalobj * 10

#If Mac Then
'        Text = "gnuplot_command:""/Library/'Application Support'/Microsoft/Office365/'User Content.localized'/Add-Ins.localized/WordMat/gnuplot/gnuplot""$set_plot_option([gnuplot_term, aqua])$draw3d("
'        Text = Text & "terminal=pdf"
'        Text = Text & ",file_name=""" & Replace(GetTempDir(), "\", "/") & "WordMatGraf""," & MacDrawDims(500, 350) & ","
#Else
    Text = "draw3d("
#End If
    Text = Text & graphobj

    Text = Text & ");"
    MaximaInputStreng = MaximaInputStreng & Text

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
Sub Wait(pausetime As Variant)
    Dim start
    start = timer    ' Set start time.
    Do While timer < start + pausetime
        DoEvents    ' Yield to other processes.
    Loop
End Sub

Function CodeForMaxima(ByVal Text As String, Optional CasType As Integer = 0) As String
' Castype=0 is maxima, CasType=1/2 is GeoGebra (geogebra friendly. still needs convertToGeogebraSyntax)
    Dim ea As New ExpressionAnalyser
    Dim ea2 As New ExpressionAnalyser
    Dim Pos As Long
    Dim posb As Long
    Dim rod As Integer
    Dim pos2 As Long
    Dim pos3 As Long
    Dim pos4 As Long
    Dim pos5 As Long
    Dim posog As Long
    Dim Arr As Variant, Arr2() As String, Arr3() As String
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

    On Error GoTo Fejl
    ea.SetNormalBrackets
    ea2.StartBracket = VBA.ChrW(12310)
    ea2.EndBracket = VBA.ChrW(12311)

    Text = Trim(Text)

    '    For j = 1 To Len(text)
    '    i = AscW(Mid(text, j, 1))
    '    s = s & Mid(text, j, 1) & " - " & i & vbCrLf
    '    Next
    '    MsgBox s

    Text = Replace(Text, VBA.ChrW(9633), "")    ' characters before parentheses that make them invisible. Possibly problematic as the parentheses are not removed. also makes fractions small
    Text = Replace(Text, VBA.ChrW(11), "")    ' shift-enter
    Text = Replace(Text, vbLf, "")    ' shift-enter and enter
    Text = Replace(Text, vbCrLf, "")
    Text = Replace(Text, vbCr, "")
    Text = Replace(Text, """", "")    ' apostrophe is removed
    Text = Replace(Text, "+-", "-")    ' plus-minus is minus
    Text = Replace(Text, VBA.ChrW(8212), "+")    'double minus is plus
    Text = Replace(Text, VBA.ChrW(8211), "-")    'A special minus sign that is slightly shorter. May appear when copying/pasting from a book.
    Text = Replace(Text, ChrW(8711) & VBA.ChrW(183), "divq ")    ' nabla mulltiply
    Text = Replace(Text, ChrW(8711) & VBA.ChrW(215), "curlq ")    ' nabla cross - curl
    Text = Replace(Text, ChrW(8711) & "^2", "laplaceq ")    ' nabla
    Text = Replace(Text, ChrW(8711), "gradq ")    ' nabla
    Text = Replace(Text, VBA.ChrW(183), "*")    ' dot is replaced with times
    Text = Replace(Text, VBA.ChrW(8901), "*")    ' \cdot , don't know why there seem to be two
    Text = Replace(Text, VBA.ChrW(8729), "*")    ' \cdot is replaced with times
    Text = Replace(Text, VBA.ChrW(8226), "*")    ' thick dot replaced with times
    Text = Replace(Text, VBA.ChrW(247), "/")    '
    Text = Replace(Text, VBA.ChrW(180) & VBA.ChrW(180), "^''")    ' there are three different apostrophes they are all understood as diff
    Text = Replace(Text, VBA.ChrW(96) & VBA.ChrW(96), "^''")    ' apostrophe
    Text = Replace(Text, VBA.ChrW(180), "^'")    ' apostrophe
    Text = Replace(Text, VBA.ChrW(96), "^'")    ' apostrophe
    Text = Replace(Text, VBA.ChrW(&H393) & "_incomplete ", "gamma_incomplete")
    Text = Replace(Text, VBA.ChrW(915), "gamma")    ' big gamma
    Text = Replace(Text, VBA.ChrW(8800), "#")    ' not equal
    Text = Replace(Text, VBA.ChrW(8804), "<=")    '
    Text = Replace(Text, VBA.ChrW(8805), ">=")    '
    Text = Replace(Text, VBA.ChrW(8710), "increment")    ' speciel delta increment
    Text = Replace(Text, "^ " & VBA.ChrW(176) & "C", "degC")    ' speciel degree(\circ) og C
    Text = Replace(Text, "^ " & VBA.ChrW(8728) & "C", "degC")    ' speciel degree(\circ) og C
    Text = Replace(Text, VBA.ChrW(8451), "degC")    ' specielt oC
    Text = Replace(Text, VBA.ChrW(8728), "tgrader")    ' degree symbol
    Text = Replace(Text, "%", "/100")    ' percentage symbol
    Text = Replace(Text, VBA.ChrW(160), "")    ' before accent e.g. vector arrow 160 8407
    If Not MaximaComplex Then
        Text = Replace(Text, VBA.ChrW(8736), "Symangle")    '
    End If
    If CasType = 0 Then
        Text = Replace(Text, VBA.ChrW(176), "tgrader")    ' degree symbol
    End If

    ' not changed back
    Text = Replace(Text, ChrW(230), "ae")
    Text = Replace(Text, ChrW(248), "oe")
    Text = Replace(Text, ChrW(229), "aa")
    Text = Replace(Text, ChrW(198), "AE")
    Text = Replace(Text, ChrW(216), "OE")
    Text = Replace(Text, ChrW(197), "AA")

    'greek letters
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
    Text = Replace(Text, VBA.ChrW(969), "omega")    ' small omega
    
    ea.Text = Text
    ' remove double spaces ' introduced here 1.24. Is also just before replacespacebymult. Introduced to perhaps catch some errors with many spaces
    Do While InStr(ea.Text, "  ") > 0
        ea.Text = Replace(ea.Text, "  ", " ")
    Loop

    ' half symbol becomes 1/2 or text Symhalf depending on context. Symhalf if letter in front
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
            Text = Left(Text, Pos - 1) & "generalized_lambert_w(0," & Right(Text, Len(Text) - Pos - 4)
        End If
    Loop While Pos > 0
    'lambertW-1
    Do
        Pos = InStr(Text, "W_(-1)")
        If Pos > 0 Then
            ea.Text = Text
            Text = Left(Text, Pos - 1) & "generalized_lambert_w(-1," & Right(Text, Len(Text) - Pos - 7)
        End If
    Loop While Pos > 0

    ' autodetection of matrix brackets
    If InStr(Text, "[") > 0 Then
        matrixstartbracket = "["
        matrixendbracket = "]"
    End If

    ' vector symbol for SymVecta
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
                ea.Text = Left(ea.Text, ea.Pos - 2) & s & "SymVecta" & Right(ea.Text, Len(ea.Text) - Pos)
            Else
                '            AddDefinition ("definer: declare(" & ea.ChrByIndex(pos - 1) & "SymVecta,nonscalar)")
                s = ea.ChrByIndex(Pos - 1)
                MaximaInputStreng = MaximaInputStreng & "declare(" & s & "SymVecta,nonscalar)$"
                KillDef = KillDef & s & "SymVecta,"
                ea.Text = Left(ea.Text, Pos - 1) & "SymVecta" & Right(ea.Text, Len(ea.Text) - Pos)
            End If
            Pos = Pos + 7
        End If
    Loop While Pos > 0

    ' vector character 2 (semi arrow) to SymVectb
    Pos = 1
    Do
        Pos = InStr(Pos, ea.Text, VBA.ChrW(8401))
        If Pos > 0 Then
            matvekt = True
            If ea.ChrByIndex(Pos - 1) = ")" Then
                s = ea.GetPrevBracketContent(Pos)
                ea.Text = Left(ea.Text, ea.Pos - 2) & s & "SymVectb" & Right(ea.Text, Len(ea.Text) - Pos)
            Else
                ea.Text = Left(ea.Text, Pos - 1) & "SymVectb" & Right(ea.Text, Len(ea.Text) - Pos)
            End If
            Pos = Pos + 7
        End If
    Loop While Pos > 0

    ' matrix symbol for SymVectc
    Pos = 1
    Do
        Pos = InStr(Pos, ea.Text, VBA.ChrW(773))
        If Pos > 0 Then
            matvekt = True
            If ea.ChrByIndex(Pos - 1) = ")" Then
                s = ea.GetPrevBracketContent(Pos)
                ea.Text = Left(ea.Text, ea.Pos - 2) & s & "SymVectc" & Right(ea.Text, Len(ea.Text) - Pos)
            Else
                ea.Text = Left(ea.Text, Pos - 1) & "SymVectc" & Right(ea.Text, Len(ea.Text) - Pos)
            End If
            Pos = Pos + 7
        End If
    Loop While Pos > 0


    ' matrix symbol \overbar for SymVectc
    
    Pos = 1
    Do
        Pos = InStr(Pos, ea.Text, VBA.ChrW(175))
        If Pos > 0 Then
            matvekt = True
            s = ea.GetNextVar(Pos)
            If Len(ea.Text) = ea.Pos Then
                ea.Text = Left(ea.Text, Pos - 1) & " " & s & "SymVectc"
            Else
                ea.Text = Left(ea.Text, Pos - 1) & " " & s & "SymVectc" & Right(ea.Text, Len(ea.Text) - ea.Pos)
            End If
            Pos = Pos + 7
        End If
    Loop While Pos > 0

    Text = ea.Text

    If InStr(Text, VBA.ChrW(9608)) > 0 And InStr(Text, VBA.ChrW(9508)) <= 0 Then matvekt = True    '9608=matrix symbol, however, it is also used for piecewise defined functions, where 9508 is used as the terminating character
    If InStr(Text, VBA.ChrW(9632)) > 0 Then matvekt = True    '9632= an other vector symbol
    If InStr(Text, VBA.ChrW(166)) > 0 Then matvekt = True    ' used for vectors entered from the equation menu
    
    Text = Replace(Text, VBA.ChrW(9508), "")    ' empty space character at all times [ ] used for piecewise def. functions moved here 1.26 due to problems with vekt

    If matvekt Then  '9632?
        Text = Replace(Text, VBA.ChrW(215), "~")    ' vectorproduct
    Else
        Text = Replace(Text, VBA.ChrW(215), "*")    ' cross replaced with multiply
    End If

    'cross vector \hat
    Pos = 1
    Do
        Pos = InStr(Pos, Text, VBA.ChrW(770))
        If Pos > 0 Then
            ea.Text = Text
            If ea.ChrByIndex(Pos - 1) = ")" Then
                s = ea.GetPrevBracketContent(Pos)
                Text = Left(Text, ea.Pos - 2) & "tvvec(" & s & ")" & Right(Text, Len(Text) - Pos)
                Pos = Pos + 4
            Else
                s = ea.GetPrevVar(Pos)
                If Pos - Len(s) - 2 < 0 Then
                  MsgBox "An error orrcured with hat operator. "
                  GoTo Fejl
                Else
                   Text = Left(Text, Pos - Len(s) - 2) & "tvvec(" & Mid(Text, Pos - Len(s) - 1, Len(s)) & ")" & Right(Text, Len(Text) - Pos)
                End If
                Pos = Pos + 5
            End If
        End If
    Loop While Pos > 0

    ' trig with large letters
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

    ' remove +- symbol
    Pos = InStr(Text, VBA.ChrW(177))
    If Pos > 0 Then
        pm = InputBox(TT.A(421), TT.A(422), "-")
        If pm = "+" Or VBA.LCase(pm) = "plus" Then
            Text = Replace(Text, VBA.ChrW(177), "+")
        ElseIf pm = "-" Or VBA.LCase(pm) = "minus" Then
            Text = Replace(Text, VBA.ChrW(177), "-")
        Else
            MsgBox2 TT.A(423)
            Text = Replace(Text, VBA.ChrW(177), "+")
        End If
    End If

    ' commas and list separator
    If Not MaximaSeparator Then
        ea.Text = Text
        ea.ConvertDecSeparator
        Text = ea.Text
        '        text = Replace(text, ",", ".")
    End If
    Text = Replace(Text, ";", ",")

'  piecewise defined functions defined with a curly bracket. fork functions
   Dim ass As String, k As Long
    ea.Text = Text
    ea.SetNormalBrackets
    i = 0
    Do
        Pos = InStr(ea.Text, "{" & VBA.ChrW(9608))
        If Pos > 0 Then

'            pos2 = InStr(pos, ea.Text, VBA.ChrW(9508))
'            expr = Mid(ea.Text, pos + 2, pos2 - pos - 1)
            Expr = ea.GetNextBracketContent(Pos)
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
                        If Len(ass) > 0 Or InStr(Arr3(k), "<") > 0 Or InStr(Arr3(k), ">") > 0 Or InStr(Arr3(k), VBA.ChrW(8712)) > 0 Then
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
                  ass = Replace(ass, "" & ChrW(8743) & " ", "&&and&&")
                  ass = Replace(ass, " " & ChrW(8743) & " ", "&&and&&")
                  ass = Replace(ass, "" & ChrW(8744) & " ", "&&or&&")
                  ass = Replace(ass, " " & ChrW(8744) & " ", "&&or&&")
                  ass = Replace(ass, " AND ", "&&and&&")
                  ass = Replace(ass, " and ", "&&and&&")
                  ass = Replace(ass, " OR ", "&&or&&")
                  ass = Replace(ass, " or ", "&&or&&")
                  If CasType = 0 Then
                    s = s & sqbrAssump(SplitAssump(Trim(ass))) & "&&then&&(" & Trim(Arr3(0)) & ")"
                  Else
                    s = s & ass & "&&then&&(" & Trim(Arr3(0)) & ")"
                  End If
               Else
                    s = s & "0=0&&then&&(" & Trim(Arr3(0)) & ")"
                End If
                If j < UBound(Arr2) Then s = s & "&&elseif&&"
            Next
            ea.Text = Left(ea.Text, Pos - 1) & s & ")" & Right(ea.Text, Len(ea.Text) - Pos - Len(Expr) - 3)
            i = i + 1
        End If
    Loop While Pos > 0 And i < 1000
    Text = ea.Text
    

    ' |x| -> abs(x) '
    If CASengine = 0 Then
    ea.Text = Text
    i = 0
    Do
        Pos = InStr(ea.Text, VBA.ChrW(124))
        If Pos <= 0 Then Pos = InStr(ea.Text, VBA.ChrW(8214))
        If Pos > 0 Then
            posb = Pos
            Do
                i = i + 1    ' incorrect input can otherwise create an infinite loop
                posb = posb + 1
                c = ea.ChrByIndex(posb)
                If c = "|" Or c = VBA.ChrW(8214) Then
                    If Not (ea.IsOperator(ea.ChrByIndex(posb - 1), 4)) Then
                        If matvekt Or InStr(ea.Text, VBA.ChrW(9632)) > 0 Then
                            If CasType = 0 Then
                                ea.Text = Left(ea.Text, Pos - 1) & "vecmag(" & Mid(ea.Text, Pos + 1, posb - Pos - 1) & ")" & Right(ea.Text, Len(ea.Text) - posb)
                            Else
                                ea.Text = Left(ea.Text, Pos - 1) & "length(" & Mid(ea.Text, Pos + 1, posb - Pos - 1) & ")" & Right(ea.Text, Len(ea.Text) - posb)
                            End If
                        Else
                            ea.Text = Left(ea.Text, Pos - 1) & "abs(" & Mid(ea.Text, Pos + 1, posb - Pos - 1) & ")" & Right(ea.Text, Len(ea.Text) - posb)
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
        MsgBox TT.A(424), vbOKOnly, TT.A(105)
    End If
    i = 0
    Text = ea.Text
   End If

    ' find implied parentheses after ^ and / ' (must be after diff and other functions with comma)
    ea.Text = Text
    ' insert parentheses at e.g. log 2 -> log(2)
    ea.InsertFunctionBrackets ' v. 1.29 moved from just after d/dx notation so that log a/b does not become log(a)/b. Tested with RunTestseq, and log a/b+2 added test
    ea.InsertBracketAfter ("/")
    ea.InsertBracketBefore ("/")
        
    ' d/dx notation
    ea.Text = ConvertDifferentials(ea.Text, VBA.ChrW(8518)) ' \dd
    If Left(ea.Text, 5) = "Fejl!" Then
        Text = ea.Text
        ConvertErrorText = Text
        GoTo slut
    End If
    ea.Text = ConvertDifferentials(ea.Text, VBA.ChrW(8706)) ' \partial
    If Left(ea.Text, 5) = "Fejl!" Then
        Text = ea.Text
        ConvertErrorText = Text
        GoTo slut
    End If
    If dAsDiffChr Then
        ea.Text = ConvertDifferentials(ea.Text, "d") ' normal d
        If Left(ea.Text, 5) = "Fejl!" Then
            Text = ea.Text
            ConvertErrorText = Text
            GoTo slut
        End If
    End If
        
    ea.Text = Replace(ea.Text, VBA.ChrW(12310), "")    ' special invisible brackets are removed   moved from before trig with uppercase letters to make insertfunctionbrackets work
    ea.Text = Replace(ea.Text, VBA.ChrW(12311), "")    ' special invisible brackets are removed

    ea.Text = Replace(ea.Text, "cosh^(-1)", "acosh")
    ea.Text = Replace(ea.Text, "sinh^(-1)", "asinh")
    ea.Text = Replace(ea.Text, "tanh^(-1)", "atanh")
    ea.Text = Replace(ea.Text, "cos^(-1) ", "acos")    ' in rare circumstances there may be a space instead of a 8289 character after
    ea.Text = Replace(ea.Text, "sin^(-1) ", "asin")    ' it will be understood as a multiplication sign if not removed
    ea.Text = Replace(ea.Text, "tan^(-1) ", "atan")
    ea.Text = Replace(ea.Text, "cos^(-1)", "acos")
    ea.Text = Replace(ea.Text, "sin^(-1)", "asin")
    ea.Text = Replace(ea.Text, "tan^(-1)", "atan")

    ea.ConvertTrigPow ("sin")    ' must be after inverse trig
    ea.ConvertTrigPow ("cos")
    ea.ConvertTrigPow ("tan")
    Text = ea.Text

    ' logarithms with bases other than 10 and e
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
                base = Right(base, Len(base) - 4)
            Else
                base = Mid(ea.Text, Pos + 4, pos2 - Pos - 4)
            End If
            ea.Text = Left(ea.Text, Pos + 2) & "base" & Right(ea.Text, Len(ea.Text) - Pos - Len(base) - 4)
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


    Text = Replace(Text, VBA.ChrW(8289), "")    ' symbol defining function removed from word syntax

    'logarithms
    If ConvertLnLog Then
        ea.Text = Text
        ea.ReplaceVar "log", "log10"
        Text = ea.Text
    End If

    ' 3 og 4 root
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
                Call ea.GetNextBracketContent    ' just to find the end parenthesis
                Call ea.InsertBeforePos("^(1/" & rod & ")") ' the method here can actually cause problems with e.g. \cbrt(4)^2 as this should be interpreted as cbrt(4^2) and not cbrt(4)^2
                ea.Text = Replace(ea.Text, VBA.ChrW(8728 + rod), "", 1, 1)
                Text = ea.Text
            End If
        Loop While Pos > 0
    Next

    'square root
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
                    Text = Left(ea.Text, Pos + 1) & Right(ea.Text, Len(ea.Text) - posog)
                ElseIf UBound(Arr) = -1 Then
                    MsgBox2 TT.A(169), vbOKOnly, TT.Error
                    GoTo Fejl
                End If
            End If
        End If
    Loop While Pos > 0

    ' sum symbol
    Do
        Pos = InStr(Text, VBA.ChrW(8721))
        If Pos > 0 Then
            ea.Text = Text
            ea.Pos = Pos
            subs = ea.FindChr("_")
            sups = ea.FindChr("^")
            ends = ea.FindChr(VBA.ChrW(9618))
            Arr = Split(Mid(Text, subs + 2, sups - subs - 3), "=")
            If UBound(Arr) < 1 Then    ' if definition is =, then translate to :
                Arr = Split(Mid(Text, subs + 2, sups - subs - 3), ":")
            End If
            If UBound(Arr) < 1 Then
                MsgBox2 TT.A(425), vbOKOnly, TT.Error
                GoTo videre
            End If
            Var = Arr(0)
            startn = Arr(1)
            slutn = Mid(Text, sups + 1, ends - sups - 1)
            sluts = ea.FindHierakiBreak(1)
            If ea.ChrByIndex(ea.Pos) = "+" Or ea.ChrByIndex(ea.Pos) = "-" Or ea.ChrByIndex(ea.Pos) = "*" Or ea.ChrByIndex(ea.Pos) = "/" Then sluts = sluts - 1
            Expr = Mid(Text, ends + 1, sluts - ends)
            Text = Left(Text, Pos - 1) & "sum(" & Expr & "," & Var & "," & startn & "," & slutn & ")" & Right(Text, Len(Text) - sluts)
        End If
    Loop While Pos > 0

    ' product symbol
    Do
        Pos = InStr(Text, VBA.ChrW(8719))
        If Pos > 0 Then
            ea.Text = Text
            ea.Pos = Pos
            subs = ea.FindChr("_")
            sups = ea.FindChr("^")
            ends = ea.FindChr(VBA.ChrW(9618))
            Arr = Split(Mid(Text, subs + 2, sups - subs - 3), "=")
            If UBound(Arr) < 1 Then    ' if definition is =, then translate to :
                Arr = Split(Mid(Text, subs + 2, sups - subs - 3), ":")
            End If
            If UBound(Arr) < 1 Then
                MsgBox2 TT.A(426), vbOKOnly, TT.Error
                GoTo videre
            End If
            Var = Arr(0)
            startn = Arr(1)
            slutn = Mid(Text, sups + 1, ends - sups - 1)
            sluts = ea.FindHierakiBreak(1)
            Expr = Mid(Text, ends + 1, sluts - ends)
            Text = Left(Text, Pos - 1) & "product(" & Expr & "," & Var & "," & startn & "," & slutn & ")" & Right(Text, Len(Text) - sluts)
        End If
    Loop While Pos > 0

videre:

    'limits
    Pos = 1
    Do
        Pos = InStr(Pos, Text, "lim")
        If Pos > 0 Then
            ea.Text = Text
            If ea.ChrByIndex(Pos + 3) = VBA.ChrW(9516) Or ea.ChrByIndex(Pos + 3) = "_" Then
                ea.Pos = Pos
                Var = ea.GetNextBracketContent
                ends = ea.Pos
                sluts = ea.FindHierakiBreak(2)
                Expr = Mid(Text, ends, sluts - ends + 1)

                Arr = Split(Var, VBA.ChrW(8594))
                Var = Trim(Arr(0))
                DiffEqVar = Var
                app = Trim(Arr(1))
                If Right(app, 2) = "^+" Then
                    app = Left(app, Len(app) - 2) & ",plus"
                ElseIf Right(app, 2) = "^-" Then
                    app = Left(app, Len(app) - 2) & ",minus"
                End If
                Text = Left(Text, Pos - 1) & "limit(" & Expr & "," & Var & "," & app & ")" & Right(Text, Len(Text) - sluts)
            Else
                Pos = Pos + 3
            End If
        End If
    Loop While Pos > 0


    'derivatives
    Dim subt As String, vararr() As String, s2 As String, f2 As String
    di = 0    ' if there are more f'(2) more variables must be used
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
            F = ea.GetPrevVar(Pos - 1)
            antalm = 0
            Do
                antalm = antalm + 1
            Loop Until ea.ChrByIndex(Pos + 1 + antalm) <> "'"
            If CasType = 0 Then
                antalmS = "round(" & antalm & ")"
            Else
                antalmS = CStr(antalm)
            End If
            If ea.ChrByIndex(Pos + antalm + 1) = " " Then    ' remove any spaces after '
                Text = Left(Text, Pos + 1) & Right(Text, Len(Text) - Pos - 2)
                ea.Text = Text
            End If

            If ea.ChrByIndex(Pos + antalm + 1) = "(" Then
                '            var = ea.ChrByIndex(pos + 3 + antalm)
                Var = ea.GetNextBracketContent(Pos + 1 + antalm)
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
               subt = Right(F, Len(F) - pos2)
               For i = 1 To Len(subt)
                  If Not ea.IsLetter(Mid(subt, i, 1)) Then
                     subt = ""
                     Exit For
                  End If
               Next
            End If

            If ea.ChrByIndex(Pos - 1) = ")" Then    ' if (...)'(x)
                ' find startparanthesis?
                Udtryk = ea.GetPrevBracketContent(Pos)
                If ea.ChrByIndex(ea.Pos - 2) = "_" Then    '  f_(...)'(x)
                    F = ea.GetPrevVar(ea.Pos - 1) & "(" & Udtryk & ")"
                    GoTo msat
                ElseIf ea.IsLetterPos(ea.Pos - 2) Then    ' sin(x)'
                    F = ea.GetPrevVar(ea.Pos - 1)
                    Text = Left(Text, ea.Pos - Len(F)) & "subst(" & Var & ",qkqg" & di & ",diff(" & F & "(qkqg" & di & "),qkqg" & di & "," & antalmS & "))" & Right(Text, Len(Text) - Pos - varl - antalm)
                    di = di + 1
                Else    ' (x^2)'
                    Text = Left(ea.Text, ea.Pos - 2) & "diff(" & Udtryk & "," & Var & "," & antalmS & ")" & Right(ea.Text, Len(ea.Text) - Pos - varl - antalm)
                End If
            ElseIf ea.ChrByIndex(Pos - 2 - Len(subt)) = ")" And Len(subt) > 0 And Len(subt) = antalm Then ' partiel (...)_xy''
                  s = ea.GetPrevBracketContent(Pos - 2)
                  For i = 1 To antalm
                     s = "diff(" & s & "," & Mid(subt, i, 1) & ",1)"
                  Next
                 Text = Left(ea.Text, ea.Pos - 2) & s & Right(ea.Text, Len(ea.Text) - Pos - varl - antalm)
            
            ElseIf Not ea.ChrByIndex(Pos + antalm + 1) = "(" Then    ' is y'  variable not function
                If Len(subt) > 0 And Len(subt) = antalm Then ' if partiel y'
                  s = Left(F, pos2 - 1)
                  For i = 1 To antalm
                     s = "diff(" & s & "," & Mid(subt, i, 1) & ",1)"
                  Next
                 Text = Left(Text, Pos - 1 - Len(F)) & s & Right(Text, Len(Text) - Pos - varl - antalm)
               Else
                Text = Left(Text, Pos - 1 - Len(F)) & "diff(" & F & "," & Var & "," & antalmS & ")" & Right(Text, Len(Text) - Pos - varl - antalm)
               End If
            ElseIf ea.ChrByIndex(Pos + antalm + 1) = "(" Then    ' if f'(x)
msat:
                If DiffEqVar = Var Or CasType = 1 Or CasType = 2 Then  'if diffligning then no subst
                  If Len(subt) > 0 And Len(subt) = antalm Then ' if partiel
                        Text = Left(Text, Pos - 1 - Len(F)) & "diff(" & F & "(" & Var & ")," & Var & "," & antalmS & ")" & Right(Text, Len(Text) - Pos - varl - antalm)
                    Else
                        Text = Left(Text, Pos - 1 - Len(F)) & "substitute(diff(" & F & "(qkqg" & di & "),qkqg" & di & "," & antalmS & "),qkqg" & di & "," & Var & ")" & Right(Text, Len(Text) - Pos - varl - antalm)
                        di = di + 1
                    End If
                Else    ' ikke diff ligning
                  If Len(subt) > 0 And Len(subt) = antalm Then ' if partiel
'                     s = Left(f, pos2 - 1) & "(" & var & ")"
                     f2 = Left(F, pos2 - 1)
                     s = "last(fundef(" & f2 & "))"
                     vararr = Split(Var, ",") ' ;?
                     For i = 1 To antalm
'                        s = "diff(" & s & "," & Mid(subt, i, 1) & ",1)"
                        s = "diff(" & s & "," & Mid(subt, i, 1) & ",1)"
                     Next
                     s2 = "subst(["
                     For i = 0 To UBound(vararr)
                        s2 = s2 & "part(first(fundef(" & f2 & "))," & i + 1 & ")=" & vararr(i) & ","
                     Next
                     s = Left(s2, Len(s2) - 1) & "]," & s & ")"
                    Text = Left(Text, Pos - 1 - Len(F)) & s & Right(Text, Len(Text) - Pos - varl - antalm)
                  Else ' normal not partiel f'(x)
                    Text = Left(Text, Pos - 1 - Len(F)) & "subst(" & Var & ",qkqg" & di & ",diff(" & F & "(qkqg" & di & "),qkqg" & di & "," & antalmS & "))" & Right(Text, Len(Text) - Pos - varl - antalm)
                    di = di + 1
                  End If
                End If
            Else
                GoTo Fejl
            End If
        End If
    Loop While Pos > 0

    'integrals
    Dim integrand As String
    Pos = InStr(Text, VBA.ChrW(8747))
    If Pos > 0 Then
        ea.Text = Text
        ea.Text = Replace(ea.Text, "diff", "qiff")    ' the d will cause problems as it will be interpreted as the end of the integral
        ea.Text = Replace(ea.Text, "round", "rounq")
        i = 0
        Do
            ea.Pos = 0
            Var = vbNullString
            ea.StartBracket = VBA.ChrW(8747)
            ea.EndBracket = VBA.ChrW(8518)    'differential d
            integrand = ea.GetNextBracketContent
            Pos = ea.Pos - Len(integrand) - 2  'pos set to start-integral symbol
            pos2 = ea.Pos
            If Len(integrand) = 0 Then ' if no differential d is used
                ea.EndBracket = "d"
                integrand = ea.GetNextBracketContent(Pos - 2)
                Pos = ea.Pos - Len(integrand) - 2    'pos is set to the starting integral symbol
                pos2 = ea.Pos
                If Not (ea.IsLetterPos(ea.Pos, False) And Not (ea.IsLetterPos(ea.Pos - 2, False))) Then ' if equal to d or as of variable/function that does not start with d, then it is not the end of the integral. Therefore, a simpler method is tried to find d
                    j = 0
                    Var = vbNullString
                    Do
                        j = j + 1
                        pos3 = InStr(ea.Pos, ea.Text, "d")
                        ea.Pos = pos3 + 1
                    Loop While j < 100 And Not (ea.IsLetterPos(ea.Pos, False) And Not (ea.IsLetterPos(ea.Pos - 2, False))) And Not pos3 = 0
                    If Pos > 0 Then
                        integrand = Mid(ea.Text, Pos + 1, pos3 - Pos - 1)
                        pos2 = ea.Pos
                    Else ' inden d fundet
                        ConvertErrorText = TT.A(895)
                        GoTo slut
                    End If
                End If
                If Var = vbNullString Then Var = ea.GetNextVar(ea.Pos - 1)
            Else
                Var = "d" & ea.GetNextVar(ea.Pos - 1) ' differential d replaced by d
                ea.Pos = ea.Pos - 1
            End If
            If Len(integrand) > 0 Then
                If Len(integrand) < 3 And Mid(ea.Text, pos2 - Len(Var), 2 + Len(Var)) = "(" & Var & ")" Then    ' if speciel notation with \int (dx)/(...)
                    ea.SetNormalBrackets
                    integrand = ea.GetNextBracketContent(pos2 + 1)
                    pos2 = ea.Pos - 1
                End If
                If Len(Var) < 2 Then
                    ConvertErrorText = TT.A(427)
                    GoTo slut
                End If
                Var = Right(Var, Len(Var) - 1)
                integrand = Trim(integrand)
                pos2 = pos2 + Len(Var)
                If Len(integrand) > 0 And Len(Var) > 0 Then
                    IntegrationVars = IntegrationVars & Var & ";"
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
                    If posn > 0 And poso > 0 And posn < poss And poso < poss Then ' definite integral
                        nedre = Mid(integrand, posn + 1, poso - posn - 1)
                        upper = Mid(integrand, poso + 1, poss - poso - 1)
                        '        text = Left(text, posn - 1) & Mid(text, poss + 1, pos2 - poss - 1) & "," & Mid(text, pos2 + 1, 1) & "," & nedre & "," & upper & ")" & Mid(text, pos2 + 2, Len(text) - pos2 - 1)
                        If CasType = 0 Then
                           integrand = Right(integrand, Len(integrand) - poss) & "," & Var & "," & nedre & "," & upper & ")"
                        Else ' geogebra only supports variable x
                           ea2.Text = Right(integrand, Len(integrand) - poss)
                           If Var <> "x" Then ea2.ReplaceVar Var, "x", 1
                           integrand = ea2.Text & "," & nedre & "," & upper & ")"
                        End If
                        If Left(integrand, 1) = VBA.ChrW(9618) Then    'remove function symbol
                            integrand = Right(integrand, Len(integrand) - 1)
                        End If
                        If MaximaExact = 2 Then
                            '                        ea.text = Replace(ea.text, VBA.ChrW(8747), "NIntegrate(", 1, 1)
                            ea.Text = Left(ea.Text, Pos - 1) & "NIntegrate(" & integrand & Right(ea.Text, ea.Length - pos2 - Len(Var) + 2)
                        Else
                            ea.Text = Left(ea.Text, Pos - 1) & "Integrate(" & integrand & Right(ea.Text, ea.Length - pos2 - Len(Var) + 2)
                            '                        ea.text = Replace(ea.text, VBA.ChrW(8747), "Integrate(", 1, 1)
                        End If
                    Else    ' indefinite integral
                        integrand = integrand & "," & Var & ")"
                        If Left(integrand, 1) = VBA.ChrW(9618) Then    'remove function symbol
                            integrand = Right(integrand, Len(integrand) - 1)
                        End If
                        ea.Text = Left(ea.Text, Pos - 1) & "integrate(" & integrand & Right(ea.Text, ea.Length - pos2 - Len(Var) + 2)
                    End If
                    i = i + 1
                End If
            Else
'                MsgBox2 TT.A(895), vbOKOnly, TT.Error
                ConvertErrorText = TT.A(895)
                GoTo slut
            End If
        Loop While i < 100 And InStr(ea.Text, VBA.ChrW(8747)) > 0 'Len(integrand) > 0
        ea.Text = Replace(ea.Text, "qiff", "diff")
        ea.Text = Replace(ea.Text, "rounq", "round")
        ea.SetNormalBrackets
        Text = ea.Text
    End If

    'double integrals
    Dim var2 As String
    Pos = InStr(Text, VBA.ChrW(8748))
    If Pos > 0 Then
        ea.Text = Text
        ea.Text = Replace(ea.Text, "diff", "qiff")    ' the d will cause problems as it will be interpreted as the end of the integral
        i = 0
        Do
            ea.Pos = 0
            ea.StartBracket = VBA.ChrW(8748)
            ea.EndBracket = VBA.ChrW(8518)    'differential d
            integrand = ea.GetNextBracketContent
            Pos = ea.Pos - Len(integrand)    'pos is set to the starting integral sign
            pos2 = ea.Pos
            If Len(integrand) = 0 Then
                ea.EndBracket = "d"
                integrand = ea.GetNextBracketContent
                Pos = ea.Pos - Len(integrand) - 2    'pos is set to the starting integral sign
                pos2 = ea.Pos
            End If
            If Len(integrand) > 0 Then
                Var = ea.GetNextVar(ea.Pos - 1)
                If Len(Var) < 2 Then
                    ConvertErrorText = TT.A(427)
                    GoTo slut
                End If
                Var = Right(Var, Len(Var) - 1)
                var2 = ea.GetNextVar(ea.Pos + 1)
                If Len(var2) < 2 Or (Left(var2, 1) <> "d" And Left(var2, 1) <> VBA.ChrW(8518)) Then
                    ConvertErrorText = TT.A(428)
                    GoTo slut
                End If
                var2 = Right(var2, Len(var2) - 1)
                integrand = Trim(integrand)
                pos2 = pos2 + Len(Var)
                If Len(integrand) > 0 And Len(Var) > 0 And Len(var2) > 0 Then
                    IntegrationVars = IntegrationVars & Var & ";"
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
                        ConvertErrorText = TT.A(429)
                        GoTo slut
                    Else    ' ubestemt integrale
                        integrand = integrand & "," & Var & ")," & var2 & ")"
                        If Left(integrand, 1) = VBA.ChrW(9618) Then    'remove function symbol
                            integrand = Right(integrand, Len(integrand) - 1)
                        End If
                        ea.Text = Left(ea.Text, Pos - 1) & "integrate(integrate(" & integrand & Right(ea.Text, ea.Length - pos2 - Len(Var) - Len(var2))
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
        ea.Text = Replace(ea.Text, "diff", "qiff")    ' the d will cause problems as it will be interpreted as the end of the integral
        i = 0
        Do
            ea.Pos = 0
            ea.StartBracket = VBA.ChrW(8749)
            ea.EndBracket = VBA.ChrW(8518)    'differential d
            integrand = ea.GetNextBracketContent()
            Pos = ea.Pos - Len(integrand)    'pos is set to the starting integral sign
            pos2 = ea.Pos
            If Len(integrand) = 0 Then
                ea.EndBracket = "d"
                integrand = ea.GetNextBracketContent
                Pos = ea.Pos - Len(integrand) - 2    'pos is set to the starting integral sign
                pos2 = ea.Pos
            End If
            If Len(integrand) > 0 Then
                Var = ea.GetNextVar(ea.Pos - 1)
                If Len(Var) < 2 Then
                    ConvertErrorText = TT.A(427)
                    GoTo slut
                End If
                Var = Right(Var, Len(Var) - 1)
                var2 = ea.GetNextVar(ea.Pos + 1)
                If Len(var2) < 2 Or (Left(var2, 1) <> "d" And Left(var2, 1) <> VBA.ChrW(8518)) Then
                    ConvertErrorText = TT.A(428)
                    GoTo slut
                End If
                var2 = Right(var2, Len(var2) - 1)
                var3 = ea.GetNextVar(ea.Pos + 1)
                If Len(var3) < 2 Or (Left(var3, 1) <> "d" And Left(var3, 1) <> VBA.ChrW(8518)) Then
                    ConvertErrorText = TT.A(428)
                    GoTo slut
                End If
                var3 = Right(var3, Len(var3) - 1)
                integrand = Trim(integrand)
                pos2 = pos2 + Len(Var)
                If Len(integrand) > 0 And Len(Var) > 0 And Len(var2) > 0 Then
                    IntegrationVars = IntegrationVars & Var & ";"
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
                        ConvertErrorText = TT.A(429)
                        GoTo slut
                    Else    ' ubestemt integrale
                        integrand = integrand & "," & Var & ")," & var2 & ")," & var3 & ")"
                        If Left(integrand, 1) = VBA.ChrW(9618) Then    'remove function symbol
                            integrand = Right(integrand, Len(integrand) - 1)
                        End If
                        ea.Text = Left(ea.Text, Pos - 1) & "integrate(integrate(integrate(" & integrand & Right(ea.Text, ea.Length - pos2 - Len(Var) - Len(var2) - Len(var3) - 2)
                    End If
                    i = i + 1
                End If
            End If
        Loop While i < 100 And Len(integrand) > 0
        ea.Text = Replace(ea.Text, "qiff", "diff")
        ea.SetNormalBrackets
        Text = ea.Text
    End If

' find implied parentheses after ^ and / ' (must be after diff and other functions with commas). Must be before ea.InsertMultSigns
' moved up to account for hidden parentheses
    ea.Text = Text
    ea.SetNormalBrackets
    ea.InsertBracketAfter ("^")
    Text = ea.Text

    ' konverter a_i - > a[i] a_i,j -> a[i,j]  if set in setting
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
                        ea.Text = Left(ea.Text, Pos - 1) & "[" & Expr & "]" & Right(ea.Text, Len(ea.Text) - Pos - Len(Expr) - 1)
                    Else
                        posb = ea.FindHierakiBreak(5)    ' changed 5 because f_1^'(x)
                        Expr = ea.GetTextInterval(Pos, posb)
                        ea.Text = Left(ea.Text, Pos - 1) & "[" & Expr & "]" & Right(ea.Text, Len(ea.Text) - Pos - Len(Expr) + 1)
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
    ' remove double spaces ' this was also introduced at the beginning of codeformaxima, so probably redundant here
    Do While InStr(ea.Text, "  ") > 0
        ea.Text = Replace(ea.Text, "  ", " ")
    Loop
    ea.ReplaceSpaceByMult    ' replaces spaces with * but not if illegal
    Text = ea.Text

    Text = Replace(Text, " ", "")     ' any remaining spaces are removed
    '    text = Replace(text, " ", "*")     ' must be after diff and integral ' problematic at (x^2 )' -> diff(x^2*,x,1)

    Text = Replace(Text, "det(", "determinant2(")    ' a little uncertain as there may be functions that end in "det"
    ' Matrices
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
                Text = Left(Text, slpos) & " ## " & Right(Text, Len(Text) - slpos - 1)
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
            Text = Left(Text, stpos - 1) & mtext & Right(Text, Len(Text) - slpos)
        End If
    Loop While Pos > 0

    'vectors
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
                If ea.ChrByIndex(stpos) = "[" Then
                    Text = Left(Text, stpos - 1) & mtext & Right(Text, Len(Text) - slpos) ' remove wrapping characters if [
                Else
                    Text = Left(Text, stpos) & mtext & Right(Text, Len(Text) - slpos + 1) ' preserve wraparound if ( or |
                End If
'                text = Left(text, stpos - 1) & mtext & right(text, Len(text) - slpos) ' changed 1.31 due to vector with || where brackets disappear
            Else
                Text = Left(Text, stpos - 1) & mtext
            End If
        End If
    Loop While Pos > 0

    'vectors with "(1¦2)" notation. When entering vectors via equation menu | Parentheses | Examples and stacks | ( _ _ )
    Do
        Pos = InStr(Text, VBA.ChrW(166))
        If Pos > 0 Then
            matvekt = True
            ea.Text = Text
            ea.Pos = Pos
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
            ea.Pos = ea.FindPrevBracket(Pos - 1)
            stpos = ea.Pos
            pos2 = stpos
            
            ea.StartBracket = "|" ' If you type || about a vector, Word can remove the () in the vector, leaving only []
            ea.EndBracket = "|"
            ea.Pos = ea.FindPrevBracket(Pos - 1)
            If ea.Pos > stpos Then
                stpos = ea.Pos
            Else
                ea.SetNormalBrackets
            End If
            
            ea.StartBracket = "["
            ea.EndBracket = "]"
            ea.Pos = ea.FindPrevBracket(Pos - 1)
            If ea.Pos > stpos Then
                stpos = ea.Pos
            Else
                ea.SetNormalBrackets
            End If
            
            matrix = ea.GetNextBracketContent(stpos)
            If Not InStr(matrix, VBA.ChrW(166)) > 0 And Pos > 0 Then
                Pos = Pos - 1
                GoTo hopBracket
            End If
            slpos = ea.Pos - 1
            ea.Pos = Pos
            Arr = Split(matrix, VBA.ChrW(166))
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
'                text = Left(text, stpos - 1) & mtext & right(text, Len(text) - slpos)
                Text = Left(Text, stpos) & mtext & Right(Text, Len(Text) - slpos + 1)
            Else
                Text = Left(Text, stpos - 1) & mtext
            End If
        End If
    Loop While Pos > 0
    ea.SetNormalBrackets

' insert implied multiplication signs between letter space and letter
' works but then you can't do functions with subscript because there are spaces
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


'insert implied multiplication signs' must be after remove space
    ea.Text = Text
    ea.Pos = 1
    ea.InsertMultSigns
    Text = ea.Text
    If Radians Then
        Text = Replace(Text, VBA.ChrW(8736), " CVinkel ")    ' if complex then angle is polar notation
    Else
        Text = Replace(Text, VBA.ChrW(8736), " CVinkelgrad ")    ' if complex then angle is polar notation
    End If

    Text = Replace(Text, "gradq*", "gradq ")
    Text = Replace(Text, "divq*", "divq ")
    Text = Replace(Text, "curlq*", "curlq ")
    Text = Replace(Text, "laplaceq*", "laplaceq ")


    If matvekt Then
        If CASengine = 0 Then
        Text = Replace(Text, "^", "^^")  ' defined matrices were not handled correctly with the above. Now trying to replace all and hope it doesn't cause other problems
        Text = Replace(Text, "*", " ## ")    ' really only important for matrix multiplication and problematic when * is not matrix multiplication, as there are then problems with hierarchy 2/3 . matrix([2],[3]) here 3 is multiplied in first. However solved with own _ operator instead of .
        Else
            Text = Replace(Text, ChrW(215), ChrW(8855))
            Text = Replace(Text, "~", ChrW(8855)) ' vectorproduct
        End If
    End If

    'replace e with %e and i with %i has been moved to after implied multiplication signs as % is not considered a letter
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

    ea.Text = Text
    If ConvertLnLog Then
        ea.ReplaceVar "ln", "log"
    End If
    ea.ReplaceVar "arg", "carg"
    Text = ea.Text


    ' trig functions converts to radians
    If Not Radians Then
        Text = ConvertDegreeToRad(Text, "sin")
        Text = ConvertDegreeToRad(Text, "cos")
        Text = ConvertDegreeToRad(Text, "tan")
        Text = ConvertDegreeToRad(Text, "sec")
        Text = ConvertDegreeToRad(Text, "cot")
        Text = ConvertDegreeToRad(Text, "csc")
    End If

    If CasType = 0 Then
        Text = Replace(Text, VBA.ChrW(960), "%pi")
        Text = Replace(Text, "-" & VBA.ChrW(8734), "minf")
        Text = Replace(Text, VBA.ChrW(8734), "inf")
    Else
'        text = Replace(text, "-" & VBA.ChrW(8734), "minf")
        Text = Replace(Text, VBA.ChrW(8734), "infinity")
    End If


    'units
    If MaximaUnits Then
        Text = ConvertUnits(Text)
    Else
        Text = Replace(Text, VBA.ChrW(937), "Omega")    ' omega
    End If

    Text = Replace(Text, "&&", " ") ' method to bypass automatic multiplication sign. Used with piecewise defined functions

    '    text = Replace(text, VBA.ChrW(8744), ",")  and
    '    text = Replace(text, VBA.ChrW(8743), ",") or

    GoTo slut:
Fejl:
    MsgBox TT.A(99), vbOKOnly, TT.Error
slut:
    If Len(ConvertErrorText) > 0 Then
        MsgBox2 ConvertErrorText, vbOKOnly, TT.Error
        CodeForMaxima = ""
    Else
      CodeForMaxima = Text
    End If
End Function
Function ConvertDifferentials(Text As String, DiffChr As String) As String
' d/dx notation
' converts d/dx to maxima diff(...,x) notation
' DiffChr is the symbol for d. It could be, for example, VBA.ChrW(8518) (\dd) or VBA.ChrW(...)(\partial) or perhaps a regular d
    
    Dim num As String, dem As String, vararr(9) As String, varind As Integer, antalm2 As Integer
    Dim udtryk1 As String, udtryk2 As String, Pos As Integer, pos3 As Integer, antalm As Integer
    Dim pos2 As Integer, Var As String, i As Integer, Udtryk As String, posstart As Integer, posslut As Integer
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
                GoTo slut
            End If
            Pos = ea.Pos
            varind = 0
            If Left(num, 2) = DiffChr & "^" Then
                antalm = Mid(num, 3, 1)
                pos2 = InStr(dem, DiffChr)
                Do
                    If Mid(dem, pos2 + 2, 1) = "^" Then
                        Var = Mid(dem, pos2 + 1, 1)
                        antalm2 = Mid(dem, pos2 + 3, 1)
                        For i = varind To varind + antalm2 - 1
                            vararr(varind) = Var
                            varind = varind + 1
                        Next
                    Else
                        vararr(varind) = Mid(dem, pos2 + 1, 1)
                        varind = varind + 1
                    End If
                    pos2 = InStr(pos2 + 1, dem, DiffChr)
                Loop While pos2 > 0
                If antalm <> varind Then
                    Text = TT.A(430)
                    GoTo slut
                End If
                ElseIf Left(num, 1) = DiffChr Then
                    antalm = 1
                    vararr(0) = Mid(dem, 2, 1)
                Else
                    Text = TT.A(430)
                    GoTo slut
            End If

            For i = 0 To antalm - 1
                udtryk1 = udtryk1 & "diff("
                udtryk2 = udtryk2 & "," & vararr(i) & ",1)"
            Next
            If Len(num) > 1 And Mid(num, 2, 1) <> "^" Then ' dy/dx or df(x)/dx
                Udtryk = Right(num, Len(num) - 1)
                If Left(Udtryk, 1) = "^" Then
                    Udtryk = Right(Udtryk, Len(Udtryk) - 1)
                End If
                Text = Left(Text, Pos - 2) & udtryk1 & Udtryk & udtryk2 & Right(Text, Len(Text) - pos3 + 1)
            ElseIf Mid(Text, pos3 + 1, 1) = "(" Then ' parentheses right after d/dx operator
                Udtryk = ea.GetNextBracketContent(pos3)
                Text = Left(Text, Pos - 2) & udtryk1 & Udtryk & udtryk2 & Right(Text, Len(Text) - ea.Pos + 1)
                ' Notation for df(x)/dx is missing here
            Else ' no parentheses after d/dx operator so the expression must be found via hierarchy
                posstart = pos3
                ea.Pos = pos3 + 1
                posslut = ea.FindHierakiBreak(2)
                Udtryk = Trim(Mid(Text, pos3, posslut - posstart + 1))
                Text = Left(Text, Pos - 2) & udtryk1 & Udtryk & udtryk2 & Right(Text, Len(Text) - posslut)
            End If
        End If
    Loop While Pos > 0
slut:
    ConvertDifferentials = Text
End Function
Function ConvertUnits(ByVal Text As String) As String
    Dim ea As New ExpressionAnalyser
    Text = Replace(Text, VBA.ChrW(937), "Ohm")    ' Omega is ohm
    ea.Text = Text
    ea.ReplaceVar VBA.ChrW(176) & "C", "degC"  ' degree symbol and C
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
    
    ConvertUnits = ea.Text

End Function
Function ConvertDegreeToRad(Text As String, trigfunc As String) As String
    Dim Pos, spos As Integer, s As String
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
                Else ' geogebra.
                    If InStr(s, VBA.ChrW(176)) <= 0 Then
                     ea.InsertBeforePos (")")
                     ea.Pos = Pos + Len(trigfunc)
                     ea.InsertAfterPos ("pi/180*(")
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
' inserts % in front of constant in expression if there is no letter before or after
    Dim Pos As Integer
    Dim ea As New ExpressionAnalyser

    ea.Text = Text
    Pos = 1
    Do
        Pos = InStr(Pos, ea.Text, tegn)
        If Pos > 0 Then
            If Not (ea.IsLetterPos(Pos + 1, False) Or ea.ChrByIndex(Pos + 1) = "[" Or ea.IsLetterPos(Pos - 1, False)) Then
                ea.Text = Left(ea.Text, Pos - 1) & "%" & Right(ea.Text, Len(ea.Text) - Pos + 1)
                Pos = Pos + 2
            Else
                Pos = Pos + 1
            End If
        End If
    Loop While Pos > 0

    CodeConstantPercentSign = ea.Text
End Function
Function ConvertToAscii(ByVal Text As String) As String
' converts from word input to ascii

    Text = Replace(Text, VBA.ChrW(12310), "")    ' special invisible brackets are removed
    Text = Replace(Text, VBA.ChrW(12311), "")    ' special invisible brackets are removed
    Text = Replace(Text, VBA.ChrW(9633), "")    ' characters before parentheses that make them invisible. Possibly problematic as the parentheses are not removed. also makes fractions small
    Text = Replace(Text, VBA.ChrW(9508), "")    ' blank space character. certainly at [ ]
    Text = Replace(Text, VBA.ChrW(11), "")    ' shift-enter
    Text = Replace(Text, vbLf, "")    ' shift-enter and enter
    Text = Replace(Text, vbCrLf, "")
    Text = Replace(Text, vbCr, "")
    Text = Replace(Text, """", "")    ' apostrophe is removed
    Text = Replace(Text, "^'", ChrW(180))    ' The regular apostrophe cannot be used, you must use ´ (accent facing upwards)
    Text = Replace(Text, "'", ChrW(180))    ' The regular apostrophe cannot be used
    Text = Replace(Text, "+-", "-")    ' plus-minus is minus
    Text = Replace(Text, VBA.ChrW(8289), "")    ' symbol defining function removed from word syntax
    Text = Replace(Text, VBA.ChrW(8212), "+")    'double minus is plus

    Text = Replace(Text, ChrW(183), "*")  ' SBCL cant show cdot
    Text = Replace(Text, ChrW(215), "*")  '
    Text = Replace(Text, "  ", " ")
    Text = Replace(Text, """", "")    ' apostrophe is removed
    Text = Replace(Text, "+-", "-")    ' plus-minus is minus
    Text = Replace(Text, "%e", "e")    '

    ConvertToAscii = Text
End Function
Function ConvertToWordSymbols(ByVal Text As String) As String
    On Error GoTo Fejl
    Dim gtegn As String
    Dim ea As New ExpressionAnalyser
    Dim Pos As Long, pos2 As Long
    Dim Udtryk As String
    Dim possep As Long
    Dim Arr As Variant
    Dim mtext As String
    Dim slutp As Integer
    Dim s As String, s2 As String

    ea.SetNormalBrackets
    gtegn = MaximaGangeTegn

    If CASengine = 0 Then
        Text = Replace(Text, " ", "")    ' space removed
    Else
        Text = Replace(Text, " + ", "+")
        Text = Replace(Text, " - ", "-")
        Text = Replace(Text, " * ", "*")
        Text = Replace(Text, " / ", "/")
    End If
    Text = Replace(Text, " . ", "*")    ' dot product
    Text = Replace(Text, """", "")    ' apostrophe is removed
    Text = Replace(Text, "+-", "-")    ' plus-minus is minus
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

'exponential notation positive' before insertbracketbefore otherwise 4E-13/x -> 4E-(13)/x but c^5.0E-1 -> c^0.5*10^-1 does not work
    Pos = 0
    Do
        Pos = InStr(Pos + 1, ea.Text, "e+")
        If Pos > 1 Then
            If ea.IsNumber(Pos - 1) Then
                ea.Pos = Pos + 1
                Do
                    ea.Pos = ea.Pos + 1
                Loop While ea.IsNumber(ea.Pos)
                ' space removed to make insert bracket work
                ea.Text = Left(ea.Text, Pos - 1) & gtegn & "10^" & Mid(ea.Text, Pos + 2, ea.Pos - Pos - 2) & " " & Right(ea.Text, Len(ea.Text) - ea.Pos + 1)
                'ea.text = Left(ea.text, pos - 1) & gtegn & "10^" & Mid(ea.text, pos + 2, ea.pos - pos - 2) & Right(text, Len(ea.text) - ea.pos + 1)
            End If
        End If
    Loop While Pos > 0

    ' exponential notation negative
    Pos = 0
    Do
        Pos = InStr(Pos + 1, ea.Text, "e-")
        If Pos > 1 Then
            If ea.IsNumber(Pos - 1) Then
                ea.Pos = Pos + 1
                Do
                    ea.Pos = ea.Pos + 1
                Loop While ea.IsNumber(ea.Pos)
                ea.Text = Left(ea.Text, Pos - 1) & gtegn & "10^" & Mid(ea.Text, Pos + 1, ea.Pos - Pos - 1) & " " & Right(ea.Text, Len(ea.Text) - ea.Pos + 1)
                '    ea.text = Left(ea.text, pos - 1) & gtegn & "10^" & Mid(ea.text, pos + 1, ea.pos - pos - 1) & Right(ea.text, Len(ea.text) - ea.pos + 1)
            End If
        End If
    Loop While Pos > 1

    ' unsigned exponential notation
    Pos = 0
    Do
        Pos = InStr(Pos + 1, ea.Text, "e")
        If Pos > 1 Then
            If ea.IsNumber(Pos - 1) And ea.IsNumber(Pos + 1) Then
                ea.Pos = Pos + 1
                Do
                    ea.Pos = ea.Pos + 1
                Loop While ea.IsNumber(ea.Pos)
                ea.Text = Left(ea.Text, Pos - 1) & gtegn & "10^" & Mid(ea.Text, Pos + 1, ea.Pos - Pos - 1) & " " & Right(ea.Text, Len(ea.Text) - ea.Pos + 1)
                '    ea.text = Left(ea.text, pos - 1) & gtegn & "10^" & Mid(ea.text, pos + 1, ea.pos - pos - 1) & Right(ea.text, Len(ea.text) - ea.pos + 1)
            End If
        End If
    Loop While Pos > 1

    ' must be after exponential as 1E-1*x goes wrong
    ' 1* removed
    Pos = InStr(ea.Text, "1*")
    Do While Pos > 0
        If Pos = 1 Then
            ea.Text = Left(ea.Text, Pos - 1) & Right(ea.Text, Len(ea.Text) - Pos - 1)    ' remove 1*
        ElseIf Mid(ea.Text, Pos - 1, 1) = "*" Then  ' times in front so always remove
            ea.Text = Left(ea.Text, Pos - 1) & Right(ea.Text, Len(ea.Text) - Pos - 1)    ' remove 1*
        ElseIf Mid(ea.Text, Pos - 1, 1) = "(" Then  ' parentheses in front so always remove
            ea.Text = Left(ea.Text, Pos - 1) & Right(ea.Text, Len(ea.Text) - Pos - 1)    ' remove 1*
        ElseIf Mid(ea.Text, Pos - 1, 1) = ")" Then  ' parentheses in front so always remove
            ea.Text = Left(ea.Text, Pos - 1) & Right(ea.Text, Len(ea.Text) - Pos - 1)    ' remove 1*
        ElseIf Mid(ea.Text, Pos - 1, 1) = "=" Then  ' equal in front so always remove
            ea.Text = Left(ea.Text, Pos - 1) & Right(ea.Text, Len(ea.Text) - Pos - 1)    ' remove 1*
        ElseIf Mid(ea.Text, Pos - 1, 1) = VBA.ChrW(8776) Then  ' cirka equal in front so always remove
            ea.Text = Left(ea.Text, Pos - 1) & Right(ea.Text, Len(ea.Text) - Pos - 1)    ' remove 1*
        ElseIf ea.IsNumber(Pos - 1) Then    ' number in front so do nothing

        ElseIf ea.IsLetterPos(Pos - 1, True) Then    ' letter in front so do nothing

        ElseIf Mid(ea.Text, Pos - 1, 1) = "." Then    ' .1* in front so do nothing

        ElseIf Pos > 2 Then
            If Mid(ea.Text, Pos - 2, 2) = "^-" Then    ' ^-1 dont remove
            Else
                ea.Text = Left(ea.Text, Pos - 1) & Right(ea.Text, Len(ea.Text) - Pos - 1)    ' remove 1*
            End If
        Else
            ea.Text = Left(ea.Text, Pos - 1) & Right(ea.Text, Len(ea.Text) - Pos - 1)    ' remove 1*
        End If
        Pos = InStr(Pos + 1, ea.Text, "1*")
    Loop

    ' 1,0* removed and ,0* ->*
    Pos = InStr(ea.Text, ".0*")
    Do While Pos > 0
        If ea.ChrByIndex(Pos - 1) = "1" Then
            If Not (ea.IsNumber(Pos - 2)) Or Pos = 2 Then
                ea.Text = Left(ea.Text, Pos - 2) & Right(ea.Text, Len(ea.Text) - Pos - 2)
            Else
                ea.Text = Left(ea.Text, Pos - 1) & Right(ea.Text, Len(ea.Text) - Pos - 1)    ',0* ->*
            End If
        Else
            ea.Text = Left(ea.Text, Pos - 1) & Right(ea.Text, Len(ea.Text) - Pos - 1)    ',0* ->*
        End If
        Pos = InStr(Pos + 1, ea.Text, ".0*")
    Loop
    ' number,non-number is removed e.g. 240, -> 240 and x*,45 ->x*0,45
    
'#If Mac Then
'** MAC   .0 on all numbers because float is used instead of dectalallnum
    Pos = InStr(ea.Text, ".0")
    Do While Pos > 0
            If Not (ea.IsNumber(Pos + 2)) Then
                ea.Text = Left(ea.Text, Pos - 1) & Right(ea.Text, Len(ea.Text) - Pos - 1)
            End If
        Pos = InStr(Pos + 1, ea.Text, ".0")
    Loop
'#End If

    ' converts a[i,j] -> a_(i,j)
    If MaximaIndex Then
        Pos = 1
        Do
            Pos = InStr(Pos + 1, ea.Text, "[")
            If Pos > 1 And ea.IsLetterPos(Pos - 1, False) Then
                ea.Pos = Pos
                ea.SetSquareBrackets
                Udtryk = ea.GetNextBracketContent(Pos)
                '       if ea.ChrByIndex(ea.pos)=
                ea.Text = Left(ea.Text, Pos - 1) & "_(" & Udtryk & ")" & Right(ea.Text, Len(ea.Text) - ea.Pos + 1)

            End If
        Loop While Pos > 1
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
    Text = Replace(Text, "#", VBA.ChrW(8800))    ' not equal

    If MaximaUnits Then
        Text = Replace(Text, VBA.ChrW(937), "Ohm")    ' Omega is ohm
        Text = Replace(Text, "%%m", "m")
        Text = Replace(Text, "%mu", VBA.ChrW(956))
        ea.Text = Text
        ea.ReplaceVar "degC", VBA.ChrW(176) & "C"  ' degree symbol and C
        ea.ReplaceVar "%timer", "timer"
        ea.ReplaceVar "%min", "min"
        ea.ReplaceVar "aar", ChrW(229) & "r"
        Text = ea.Text
    End If

    Text = Replace(Text, "*", gtegn)    ' VBA.ChrW(183)

    Text = Replace(Text, "exp(", "e^(")
    Text = Replace(Text, "log", "ln" & VBA.ChrW(8289))
    Text = Replace(Text, "lgog(", "log" & VBA.ChrW(8289) & "(")
    Text = Replace(Text, "sqrt", VBA.ChrW(8730))
    Text = Replace(Text, "cbrt", VBA.ChrW(8731)) ' from geogebra

    '    text = Replace(text, "E+", gtegn & "10^")
    '    text = Replace(text, "E-", gtegn & "10^-")
        
    'root of higher degree
    If CASengine > 0 Then 'only for GeoGebra
    Do
        Pos = InStr(Text, "nroot(")
        If Pos > 0 Then
            ea.Text = Text
            Udtryk = ea.GetNextBracketContent(Pos + 4)
            pos2 = ea.Pos
'            Arr = Split(udtryk, ";")
            ea.Text = Udtryk
            ea.SplitNonBracket Arr, ";"
            If UBound(Arr) > 0 Then
                Text = Left(Text, Pos - 1) & "" & VBA.ChrW(8730) & "(" & Arr(1) & "&" & Arr(0) & ")" & Right(Text, Len(Text) - pos2 + 1)
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
            Udtryk = ea.GetNextBracketContent(Pos + 7)
            Text = Left(Text, Pos - 1) & " " & Replace(Udtryk, ListSeparator, VBA.ChrW(8800)) & " " & Right(Text, Len(Text) - ea.Pos + 1)
        End If
    Loop While Pos > 0

    ' equal(x,y) -> x=y
    Do
        Pos = InStr(Text, "equal(")
        If Pos > 0 Then
            ea.Text = Text
            Udtryk = ea.GetNextBracketContent(Pos + 4)
            Text = Left(Text, Pos - 1) & " " & Replace(Udtryk, ListSeparator, "=") & " " & Right(Text, Len(Text) - ea.Pos + 1)
        End If
    Loop While Pos > 0

    ' abs(x) -> |x|
    Do
        Pos = InStr(Text, "abs(")
        If Pos > 0 Then
            ea.Text = Text
            Udtryk = ea.GetNextBracketContent(Pos)
            Text = Left(Text, Pos - 1) & VBA.ChrW(124) & Udtryk & VBA.ChrW(124) & Right(Text, Len(Text) - ea.Pos + 1)
        End If
    Loop While Pos > 0

    ' 'vecmag(x) -> |x|
    Do
        Pos = InStr(Text, "'vecmag(")
        If Pos > 0 Then
            ea.Text = Text
            Udtryk = ea.GetNextBracketContent(Pos + 6)
            Text = Left(Text, Pos - 1) & VBA.ChrW(124) & Udtryk & VBA.ChrW(124) & Right(Text, Len(Text) - ea.Pos + 1)
        End If
    Loop While Pos > 0

    'atan2(x,y) to atan(x/y)
    Pos = InStr(Text, "atan2")
    Do While Pos > 0
        ea.Text = Text
        ea.GetNextBracketContent (Pos)
        Text = Left(Text, ea.Pos - 1) & ")" & Right(Text, Len(Text) - ea.Pos + 1)    'closing parenthesis in denominator
        Text = Left(Text, Pos + 3) & "(" & Right(Text, Len(Text) - Pos - 4)    ' , atan2 to atan(
        possep = InStr(Pos + 4, Text, ListSeparator)
        Text = Left(Text, possep - 1) & ")/(" & Right(Text, Len(Text) - possep)    ' , to /

        Pos = InStr(Pos, Text, "atan2")
    Loop

    ' integrate to integral sign
    Pos = InStr(Text, "integrate")
    Do While Pos > 0
        If Mid(Text, Pos - 1, 1) = "'" Then
            Text = Left(Text, Pos - 2) & Right(Text, Len(Text) - Pos + 1)
            Pos = Pos - 1
        End If
        Dim udt As String
        ea.Text = Text
        ea.Pos = Pos + 8
        udt = ea.GetNextBracketContent()
        Arr = Split(udt, ListSeparator)
        If UBound(Arr) > 1 Then
            Text = Left(Text, Pos - 1) & VBA.ChrW(8747) & "_" & Arr(2) & "^" & Arr(3) & VBA.ChrW(9618) & Arr(0) & "  d" & Arr(1) & " " & Right(Text, Len(Text) - Pos - Len(udt) - 10)
        Else
            Text = Left(Text, Pos - 1) & VBA.ChrW(8747) & VBA.ChrW(9618) & Arr(0) & "  d" & Arr(1) & " " & Right(Text, Len(Text) - Pos - Len(udt) - 10)
        End If
        Pos = InStr(Pos, Text, "integrate")
    Loop

    ' differentiate
    Pos = InStr(Text, "diff")
    Do While Pos > 0
        If Mid(Text, Pos - 1, 1) = "'" Then
            Text = Left(Text, Pos - 2) & Right(Text, Len(Text) - Pos + 1)
            Pos = Pos - 1
        End If
        ea.Text = Text
        ea.Pos = Pos + 4
        udt = ea.GetNextBracketContent()
        Arr = Split(udt, ListSeparator)
        If UBound(Arr) > 1 Then
            If Arr(2) = "1" Then
                Text = Left(Text, Pos - 1) & VBA.ChrW(8518) & "/(" & VBA.ChrW(8518) & Arr(1) & ") (" & Arr(0) & ")" & " " & Right(Text, Len(Text) - Len(udt) - 5 - Pos)
            Else
                Text = Left(Text, Pos - 1) & VBA.ChrW(8518) & "/(" & VBA.ChrW(8518) & "^" & Arr(2) & " " & Arr(1) & ") (" & Arr(0) & ")" & " " & Right(Text, Len(Text) - Len(udt) - 6)
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
            Text = Left(Text, Pos - 2) & Right(Text, Len(Text) - Pos + 1)
            Pos = Pos - 1
        End If
        ea.Text = Text
        ea.Pos = Pos + 4
        udt = ea.GetNextBracketContent()
        Arr = Split(udt, ListSeparator)
        'lim" & VBA.ChrW(9516) & "(x" & VBA.ChrW(8594) & "1)" & VBA.ChrW(8289) & "(x+1)
        Text = Left(Text, Pos - 1) & "lim" & VBA.ChrW(9516) & "(" & Arr(1) & VBA.ChrW(8594) & Arr(2) & ")" & VBA.ChrW(8289) & "(" & Arr(0) & ") " & Right(Text, Len(Text) - Pos - Len(udt) - 6)
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

    ' matrices
    Pos = InStr(Text, "matrix")
    Do While Pos > 0

        ea.Text = Text
        mtext = ea.GetNextBracketContent(Pos + 5)
        mtext = Replace(mtext, "]" & ListSeparator & "[", "@")
        mtext = Replace(mtext, ListSeparator, "&")

        mtext = matrixstartbracket & VBA.ChrW(9632) & "(" & Mid(mtext, 2, Len(mtext) - 2) & ")" & matrixendbracket
        Text = Left(Text, Pos - 1) & mtext & Right(Text, Len(Text) - ea.Pos + 1)

        Pos = InStr(Pos + 1, Text, "matrix")
    Loop

    ' 'and i %if
    Pos = InStr(Text, "'and(")
    If Pos > 0 Then
        ea.Text = Text
        ea.Pos = Pos + 3
        udt = ea.GetNextBracketContent()
        udt = Replace(udt, ListSeparator, "  " & VBA.ChrW(8743) & "  ")
        Text = Left(Text, Pos - 1) & udt & Right(Text, Len(Text) - ea.Pos + 1)
    End If

    ' %if
    Pos = InStr(Text, "%if")
    Do While Pos > 0
        ea.Text = Text
        mtext = ea.GetNextBracketContent(Pos + 2)
        slutp = ea.Pos
        Arr = Split(mtext, ListSeparator)
        Text = Left(Text, Pos - 1) & Arr(1) & "   for   " & Arr(0) & Right(Text, Len(Text) - slutp + 1)
        Pos = InStr(Text, "%if")
    Loop

    'vecta_ -> vector arrow
    Pos = InStr(Text, "vecta_")
    Do While Pos > 0
        ea.Text = Text
        s = ea.GetNextVar(Pos + 6)
        Text = Left(Text, Pos - 1) & "(" & s & ")" & VBA.ChrW(160) & VBA.ChrW(8407) & Right(Text, Len(Text) - ea.Pos)
        Pos = InStr(Text, "vecta_")
    Loop
    'vectb_ -> semi vector arrow
    Pos = InStr(Text, "vectb_")
    Do While Pos > 0
        ea.Text = Text
        s = ea.GetNextVar(Pos + 6)
        Text = Left(Text, Pos - 1) & "(" & s & ")" & VBA.ChrW(160) & VBA.ChrW(8401) & Right(Text, Len(Text) - ea.Pos)
        Pos = InStr(Text, "vectb_")
    Loop
    'vectc_ -> matrix line
    Pos = InStr(Text, "vectc_")
    Do While Pos > 0
        ea.Text = Text
        s = ea.GetNextVar(Pos + 6)
        Text = Left(Text, Pos - 1) & "(" & s & ")" & VBA.ChrW(160) & VBA.ChrW(773) & Right(Text, Len(Text) - ea.Pos)
        Pos = InStr(Text, "vectc_")
    Loop

    'greek letters
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
    Text = Replace(Text, "eta", VBA.ChrW(951))    ' must be last as eta is included in others
    Text = Replace(Text, "increment", VBA.ChrW(8710))  ' specielt delta increment
    Text = Replace(Text, "Symhalf", VBA.ChrW(189))    ' _
    text = Replace(text, "degC", VBA.ChrW(8451))   ' speciel oC symbol
    Text = Replace(Text, "<=", VBA.ChrW(8804))  '
    Text = Replace(Text, ">=", VBA.ChrW(8805))    '
    Text = Replace(Text, "CVinkelO", VBA.ChrW(8736))    '
    Text = Replace(Text, "CVinkel", VBA.ChrW(8736))    '
    Text = Replace(Text, "Symangle", VBA.ChrW(8736))  '

    ea.Text = Text
    ' SymVecta to vector arrow
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
                ea.Text = Left(ea.Text, Pos - 1 - Len(s)) & "(" & s & ")" & VBA.ChrW(8407) & Right(ea.Text, Len(ea.Text) - Pos - 7)
            End If
            Pos = Pos + 1
        End If
    Loop While Pos > 0

    ' SymVectb to semi vector arrow
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
                ea.Text = Left(ea.Text, Pos - 1 - Len(s)) & "(" & s & ")" & VBA.ChrW(8401) & Right(ea.Text, Len(ea.Text) - Pos - 7)
            End If
            Pos = Pos + 1
        End If
    Loop While Pos > 0

    ' SymVectc to semi vector arrow
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
                ea.Text = Left(ea.Text, Pos - 1 - Len(s)) & "(" & s & ")" & VBA.ChrW(773) & Right(ea.Text, Len(ea.Text) - Pos - 7)
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
            ea.Text = Left(ea.Text, Pos - 1) & "W_(" & s2 & ") (" & Arr(1) & ")" & Right(ea.Text, Len(ea.Text) - ea.Pos + 1)
        End If
        End If
    Loop While Pos > 0

    ' These may result from limit value
    ea.ReplaceVar "inf", VBA.ChrW(8734)
    ea.ReplaceVar "minf", "-" & VBA.ChrW(8734)
    ea.ReplaceVar "infinity", VBA.ChrW(8734)
    ea.ReplaceVar "und", TT.A(431)
    ea.ReplaceVar "ind", TT.A(432)
    ea.ReplaceVar "tgrader", VBA.ChrW(176)
    Text = ea.Text

    Text = Replace(Text, "SPC", " ")

    'no complex expressions
    If Not MaximaComplex Then
        If InStr(Text, "%i") > 0 Or InStr(Text, VBA.ChrW(943)) > 0 Then
            Text = TT.A(433) & " " & VBA.ChrW(8477)
            GoTo slut
        End If
    Else
        Text = Replace(Text, "%i", "i")
    End If

    GoTo slut
Fejl:
    MsgBox TT.A(100) & vbCrLf & vbCrLf & Text, vbOKOnly, TT.Error
slut:
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
'only real solutions' problem that even though there are in in it can perhaps be reduced to real
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
        MaximaOutput = Replace(MaximaOutput, ListSeparator, "    " & VBA.ChrW(8744) & "    ")    ' Comma replaced by or
    End If

End Sub
Function ConvertToLatex(Text As String) As String
    Dim Pos As Long, pos1 As Long, pos2 As Long, pos3 As Long
    Dim ea As New ExpressionAnalyser
    Dim ea2 As New ExpressionAnalyser
    Dim s As String, t As String
    Dim Arr As Variant
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

    'various symbols
'    text = Replace(text, VBA.ChrW(183), "* ")    ' dot is replaced with times
    Text = Replace(Text, VBA.ChrW(8901), VBA.ChrW(183))     ' \cdot , don't know why there seem to be two
    Text = Replace(Text, VBA.ChrW(8729), VBA.ChrW(183))    ' \cdot replaced by multiply (this is \bullet)
    Text = Replace(Text, VBA.ChrW(8226), VBA.ChrW(183))    ' thick dot replaced with times
    
    Text = Replace(Text, VBA.ChrW(247), "\div ")    ' division sign with dots above/below


    Text = Replace(Text, vbCrLf, "")
    Text = Replace(Text, vbCr, "")
    Text = Replace(Text, vbLf, "")

    Text = Replace(Text, VBA.ChrW(9633), "")    ' characters before parentheses that make them invisible. Possibly problematic as the parentheses are not removed. also makes fractions small
    Text = Replace(Text, VBA.ChrW(9508), "")    ' blank space character certainly at [ ]
    Text = Replace(Text, VBA.ChrW(11), "")    ' shift-enter
    Text = Replace(Text, """", "")    ' apostrophe is removed
    Text = Replace(Text, VBA.ChrW(180) & VBA.ChrW(180), "^''")    ' there are three different apostrophes that look similar. they are all understood as diff
    Text = Replace(Text, VBA.ChrW(96) & VBA.ChrW(96), "^''")    ' there are three different apostrophes that look similar. they are all understood as diff
    Text = Replace(Text, VBA.ChrW(180), "^'")    ' there are three different apostrophes that look similar. they are all understood as diff
    Text = Replace(Text, VBA.ChrW(96), "^'")    ' there are three different apostrophes that look similar. they are all understood as diff
    '    text = Replace(text, VBA.ChrW(176), "\circ ") ' degree symbol
    Text = Replace(Text, VBA.ChrW(189), "\frac{1}{2} ")    ' _
    text = Replace(text, VBA.ChrW(8800), "\neq ")    ' not equal
    Text = Replace(Text, VBA.ChrW(177), "\pm ")    ' plus/minus
    Text = Replace(Text, VBA.ChrW(8723), "\mp ")    ' minus-plus
    Text = Replace(Text, VBA.ChrW(8804), "\leq ")    ' <=
    Text = Replace(Text, VBA.ChrW(8805), "\geq ")    ' >=
    Text = Replace(Text, VBA.ChrW(8776), "\approx ")    ' cirka equal
    Text = Replace(Text, VBA.ChrW(8773), "\cong ")    ' cirka equal
    Text = Replace(Text, VBA.ChrW(8656), "\Leftarrow ")    '
    Text = Replace(Text, VBA.ChrW(8658), "\Rightarrow ")    '
    Text = Replace(Text, VBA.ChrW(8660), "\Leftrightarrow ")    ' bi implication
    Text = Replace(Text, VBA.ChrW(8661), "\Updownarrow ")    ' bi implication up/down
    Text = Replace(Text, VBA.ChrW(8657), "\Uparrow ")    ' doublearrow up
    Text = Replace(Text, VBA.ChrW(8659), "\Downarrow ")    ' doublearrow down
    Text = Replace(Text, VBA.ChrW(8593), "\uparrow ")    '
    Text = Replace(Text, VBA.ChrW(8595), "\downarrow ")    '
    Text = Replace(Text, VBA.ChrW(8596), "\leftrightarrow ")    '
    Text = Replace(Text, VBA.ChrW(10187), "\partial ")    '
    Text = Replace(Text, VBA.ChrW(8469), "\mathbb{N} ")    ' symbol for natural numbers N
    Text = Replace(Text, VBA.ChrW(8484), "\mathbb{Z} ")    ' symbol for whole numbers Z
    Text = Replace(Text, VBA.ChrW(8474), "\mathbb{Q} ")    ' symbol for rationel numbers Q
    Text = Replace(Text, VBA.ChrW(8477), "\mathbb{R} ")    ' symbol for real numbers R
    Text = Replace(Text, VBA.ChrW(8450), "\mathbb{C} ")    ' symbol for complex numbers C
    Text = Replace(Text, VBA.ChrW(8463), "\hbar ")    ' planck constant/2pi
    Text = Replace(Text, VBA.ChrW(8704), "\forall ")    ' math for all symbol
    Text = Replace(Text, VBA.ChrW(8712), "\in ")    ' math belongs to symbol
    Text = Replace(Text, VBA.ChrW(8733), "\propto ")    ' propotional symbol
    Text = Replace(Text, VBA.ChrW(8810), "\ll ")    ' <<
    Text = Replace(Text, VBA.ChrW(8811), "\gg ")    ' >>
    Text = Replace(Text, VBA.ChrW(8704), "\exists ")    ' exist symobl
    Text = Replace(Text, VBA.ChrW(8708), "\exists ")    ' exist symobl
    Text = Replace(Text, VBA.ChrW(8713), "\notin ")    ' math does not belong to symbol
    Text = Replace(Text, VBA.ChrW(8709), "\emptyset ")    ' empty set symbol
    Text = Replace(Text, VBA.ChrW(8711), "\nabla ")    ' nabla operator
    Text = Replace(Text, VBA.ChrW(8801), "\equiv ")    ' equal to with 3 dashes
    Text = Replace(Text, VBA.ChrW(8710), "\inc ")    ' specielt delta increment
    Text = Replace(Text, VBA.ChrW(8943), "\ldots ")    ' ...
    Text = Replace(Text, VBA.ChrW(8942), "\vdots ")    ' 3 dots vertial
    Text = Replace(Text, VBA.ChrW(8944), "\rddots ")    ' 3 dots diagonally up to the right
    Text = Replace(Text, VBA.ChrW(8945), "\ddots ")    ' 3 dots diagonally down to the right
    Text = Replace(Text, VBA.ChrW(8756), "\therefore ")    ' 3 dots in triangle
    Text = Replace(Text, VBA.ChrW(8757), "\because ")    ' 3 prikker i trekant
    Text = Replace(Text, VBA.ChrW(8834), "\subset ")    ' belongs to set (real subset)
    Text = Replace(Text, VBA.ChrW(8838), "\subseteq ")    ' belongs to equal set
    Text = Replace(Text, VBA.ChrW(8745), "\cap ")    ' common set
    Text = Replace(Text, VBA.ChrW(8746), "\cup ")    ' unified set
    Text = Replace(Text, VBA.ChrW(8764), "\sim ")    'ca equal (a wave)
    Text = Replace(Text, VBA.ChrW(8771), "\simeq ")    ' ca equal (1 wave, 1 straight)
    Text = Replace(Text, VBA.ChrW(8855), "\otimes ")    ' set product
    Text = Replace(Text, VBA.ChrW(8614), "\mapsto ")    ' Arrow to function definition
    Text = Replace(Text, VBA.ChrW(8640), "\rightharpoonup ")    ' right arrow with only half arrow
    Text = Replace(Text, VBA.ChrW(8651), "\lrhar ")    ' harpoon right/left
    Text = Replace(Text, VBA.ChrW(8598), "\nwarrow ")    ' arrow up left
    Text = Replace(Text, VBA.ChrW(8599), "\nearrow ")    ' arrow up right
    Text = Replace(Text, VBA.ChrW(8600), "\nwarrow ")    ' arrow down right
    Text = Replace(Text, VBA.ChrW(8601), "\swarrow ")    ' down arrow left
    Text = Replace(Text, VBA.ChrW(8728), "\circ ")    ' degree symbol
    Text = Replace(Text, VBA.ChrW(8451), "^\circ\text{C} ")    ' speciel oC symbol
    Text = Replace(Text, VBA.ChrW(8457), "^\circ\text{F} ")    ' speciel oF symbol
    Text = Replace(Text, VBA.ChrW(176), "^\circ ")    '
    Text = Replace(Text, VBA.ChrW(8736), "\angle ")    '
    Text = Replace(Text, VBA.ChrW(160), "")    ' before accent e.g. vector arrow 160 8407
    Text = Replace(Text, VBA.ChrW(8743), "\wedge ")    ' and
    Text = Replace(Text, VBA.ChrW(8744), "\vee ")    ' or

'    text = Replace(text, VBA.ChrW(8721), "\sum ")    ' sum sign. Maybe it should be made more advanced
'    text = Replace(text, VBA.ChrW(8719), "\prod ")    ' product sign. Maybe needs to be made more advanced
    Text = Replace(Text, VBA.ChrW(8748), "\iint ")    ' double integral sign. Maybe needs to be made more advanced
    Text = Replace(Text, VBA.ChrW(8749), "\iiint ")    ' triple integral sign. Maybe needs to be made more advanced
    Text = Replace(Text, VBA.ChrW(8750), "\oint ")    ' curve integral sign. Maybe needs to be made more advanced


    ' Remove problematic invisible parentheses after and before / maybe not necessary for latex? at least problematic for x/10^2 -> x/(10)^2
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
    ea.Text = Text

    ea.Text = Replace(ea.Text, VBA.ChrW(12310), "")    ' special invisible brackets are removed moved from before trig with uppercase letters to make insertfunctionbrackets work
    ea.Text = Replace(ea.Text, VBA.ChrW(12311), "")    ' special invisible brackets are removed moved

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
            ea.Text = Left(ea.Text, Pos - 1) & "\lim " & Right(ea.Text, Len(ea.Text) - Pos - 2)
        End If
    Loop While Pos > 0

'    ea.text = Replace(ea.text, "cosh^(-1)", "cosh^{-1}")
'    ea.text = Replace(ea.text, "sinh^(-1)", "sinh^{-1}")
'    ea.text = Replace(ea.text, "tanh^(-1)", "tanh^{-1}")
'    ea.text = Replace(ea.text, "cos^(-1) ", "cos^{-1}")    ' in rare circumstances there may be a space instead of a 8289 character after
'    ea.text = Replace(ea.text, "sin^(-1) ", "sin^{-1}")    ' it will be understood as a multiplication sign if not removed
'    ea.text = Replace(ea.text, "tan^(-1) ", "tan^{-1}")
'    ea.text = Replace(ea.text, "cos^(-1)", "cos^{-1}")
'    ea.text = Replace(ea.text, "sin^(-1)", "sin^{-1}")
'    ea.text = Replace(ea.text, "tan^(-1)", "tan^{-1}")

    ea.Text = ConvertTrigPowLatex("sin", ea.Text) ' must be before function sign is removed
    ea.Text = ConvertTrigPowLatex("cos", ea.Text) ' must be before function sign is removed
    ea.Text = ConvertTrigPowLatex("tan", ea.Text) ' must be before function sign is removed

    ea.Text = Replace(ea.Text, VBA.ChrW(8289), "") ' hidden function symbol

    'fbox
    ea.Text = ConvertPreAccent(ea.Text, VBA.ChrW(9645), "\fbox{", "}")

    'overbrace
    ea.Text = ConvertPreAccent(ea.Text, VBA.ChrW(9182), "\overbrace{", "}")

    'underbrace
    ea.Text = ConvertPreAccent(ea.Text, VBA.ChrW(9183), "\underbrace{", "}")


    'vector-notation with arrow over letter
    ea.Text = ConvertAccent(ea.Text, VBA.ChrW(8407), "\overrightarrow{", "}")

    'cross vector \hat
    ea.Text = ConvertAccent(ea.Text, VBA.ChrW(770), "\widehat{", "}")
    
    '\overline
    ea.Text = ConvertAccent(ea.Text, VBA.ChrW(773), "\overline{", "}")
    'double overline
    ea.Text = ConvertAccent(ea.Text, VBA.ChrW(831), "\overline{\overline{", "}}")
    
    'semi arrow
    ea.Text = ConvertAccent(ea.Text, VBA.ChrW(8401), "\overset{\rightharpoonup}{", "}")

    'double arrow over
    ea.Text = ConvertAccent(ea.Text, VBA.ChrW(8417), "\overleftrightarrow{", "}")
    
    'left arrow over
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
            ea.Text = Left(ea.Text, pos1) & "\underset{" & t & "}{" & s & "}" & Right(ea.Text, Len(ea.Text) - pos2)
        End If
    Loop While Pos > 0

    ' sum symbol
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
            ea.Text = Left(ea.Text, Pos - 1) & "\sum " & Right(ea.Text, Len(ea.Text) - Pos)
        End If
    Loop While Pos > 0

    ' product symbol
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
            ea.Text = Left(ea.Text, Pos - 1) & "\prod " & Right(ea.Text, Len(ea.Text) - Pos)
        End If
    Loop While Pos > 0

    ' integral
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
            If Left(ll, 1) = "(" And Right(ll, 1) = ")" Then ll = Mid(ll, 2, Len(ll) - 2) ' Word usually inserts brackets that should not be seen. If there should be brackets, there is a double
            If Left(ul, 1) = "(" And Right(ul, 1) = ")" Then ul = Mid(ul, 2, Len(ul) - 2)
            ea.Text = Left(ea.Text, Pos - 1) & "\int_{" & ll & "}^{" & ul & "} " & Right(ea.Text, Len(ea.Text) - pos3)
        Else    ' indefinite
            ea.Text = Left(ea.Text, Pos - 1) & "\int " & Right(ea.Text, Len(ea.Text) - pos3)
        End If
        Pos = InStr(ea.Text, VBA.ChrW(8747))
    Loop

    ' ^
    Pos = InStr(ea.Text, "^")
    Do While Pos > 0
        If ea.ChrByIndex(Pos + 1) = "(" Then
            s = Trim(ea.GetNextBracketContent(Pos))
            If Not (Left(s, 1) = "{" And Right(s, 1) = "}") Then s = "{" & s & "}"
            ea.Text = Left(ea.Text, Pos) & s & Right(ea.Text, Len(ea.Text) - ea.Pos + 1)
        ElseIf ea.ChrByIndex(Pos + 1) = "{" Then ' if already converted
        Else
            ea.Pos = Pos + 1
            ea.FindHierakiBreak 4, False
            s = ea.SubString(Pos + 1, ea.Pos)
            If Not (Left(s, 1) = "{" And Right(s, 1) = "}") Then s = "{" & s & "}"
            ea.Text = Left(ea.Text, Pos) & s & Right(ea.Text, NotZero(ea.Length - ea.Pos))
        End If
        Pos = InStr(Pos + 1, ea.Text, "^")
    Loop

    '_
    Pos = InStr(ea.Text, "_")
    Do While Pos > 0
        If ea.ChrByIndex(Pos + 1) = "(" Then
            s = ea.GetNextBracketContent
            If Not (Left(s, 1) = "{" And Right(s, 1) = "}") Then s = "{" & s & "}"
            ea.Text = Left(ea.Text, Pos) & s & Right(ea.Text, Len(ea.Text) - ea.Pos + 1)
        Else
            ea.Pos = Pos + 1
            ea.FindHierakiBreak 4, False
            s = ea.SubString(Pos + 1, ea.Pos)
            If Not (Left(s, 1) = "{" And Right(s, 1) = "}") Then s = "{" & s & "}"
            ea.Text = Left(ea.Text, Pos) & s & Right(ea.Text, NotZero(ea.Length - ea.Pos))
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
        If Left(n, 1) = "(" And Right(n, 1) = ")" Then
            n = Mid(n, 2, Len(n) - 2)
        End If
        pos2 = pos2 + 1
        '        End If
        ea.Text = Left(ea.Text, pos1 - 1) & "\frac{" & t & "}{" & n & "}" & Right(ea.Text, NotZero(Len(ea.Text) - pos2 + 1))
        Pos = InStr(ea.Text, "/")
    Loop

    ' 3 og 4 root
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
                    ea.Text = Left(ea.Text, Pos - 1) & "\sqrt[" & rod & "]{" & s & "}" & Right(ea.Text, NotZero(ea.Length - ea.Pos + 1))
                End If
                ea.Pos = Pos
            End If
        Loop While Pos > 0
    Next

    'square root
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
                    ea.Text = Left(ea.Text, Pos - 1) & "\sqrt{" & Arr(0) & "}" & Right(ea.Text, NotZero(ea.Length - ea.Pos + 1))
                ElseIf UBound(Arr) = 1 Then
                    ea.Text = Left(ea.Text, Pos - 1) & "\sqrt[" & Arr(0) & "]{" & Arr(1) & "}" & Right(ea.Text, NotZero(ea.Length - ea.Pos + 1))
                End If
            End If
        End If
    Loop While Pos > 0

    ' matrices
    ea.Text = Replace(ea.Text, VBA.ChrW(9608), VBA.ChrW(9632))
    Pos = InStr(ea.Text, VBA.ChrW(9632))
    Do While Pos > 0
        ea.SetNormalBrackets
        s = ea.GetNextBracketContent(Pos)
        ea.Text = Left(ea.Text, Pos - 1) & "\begin{matrix}" & Replace(s, "@", "\\") & "\end{matrix}" & Right(ea.Text, NotZero(ea.Length - Pos - Len(s) - 2))
        Pos = InStr(ea.Text, VBA.ChrW(9632))
    Loop

    ' functions
    ea.ReplaceVar "sin", "\sin "
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

    ' units not in italics \,\text{cm}
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
        ea.Text = Replace(ea.Text, VBA.ChrW(8289), "")    ' function symbol
        s = ea.GetNextVar()
    Loop While s <> ""
    End If

    'greek letters
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

    ea.Text = Replace(ea.Text, VBA.ChrW(8592), "\leftarrow ")
    ea.Text = Replace(ea.Text, VBA.ChrW(8594), "\rightarrow ")
    ea.Text = Replace(ea.Text, VBA.ChrW(8652), "\rightleftharpoons ")    ' actually \leftrightharpoons but not supported by default Latex
    ea.Text = Replace(ea.Text, VBA.ChrW(8651), "\rightleftharpoons ")
    ea.Text = Replace(ea.Text, VBA.ChrW(8592), "\leftarrow ")
    ea.Text = Replace(ea.Text, VBA.ChrW(10236), "\longmapsto ")
    ea.Text = Replace(ea.Text, VBA.ChrW(8640), "\rightharpoonup ")
    ea.Text = Replace(ea.Text, VBA.ChrW(8641), "\rightharpoondown ")
    ea.Text = Replace(ea.Text, VBA.ChrW(183), "\cdot ")    ' dot is replaced with times
    ea.Text = Replace(ea.Text, VBA.ChrW(215), "\times ")    ' x multiplication symbol
    ea.Text = Replace(ea.Text, " \", "\") ' In Moodle, spaces are created. Therefore, they should preferably be removed. Spaces before \ are not necessary.

    ' brackets
    ea.Text = Replace(ea.Text, "(", "\left(")    '
    ea.Text = Replace(ea.Text, ")", "\right)")    '
    
    'square brackets ' not so simple because they are part of latex syntax
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
                ea.Text = Left(ea.Text, pos2 - 1) & "\left" & Right(ea.Text, Len(ea.Text) - pos2 + 1)
                ea.Text = Left(ea.Text, pos3 + 3) & "\right" & Right(ea.Text, Len(ea.Text) - pos3 - 3)
            End If
        Else
            pos1 = Pos + 1
        End If
    Loop While Pos > 0

    ConvertToLatex = ea.Text
End Function
Function ConvertTrigPowLatex(Trig As String, Text As String) As String
' converts  sin^2(x) -> sin^2(x)
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
        Text = Left(Text, Pos - 1) & "{" & pow & "}" & Right(Text, Len(Text) - pos2)
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
        If ea.ChrByIndex(Pos + 1) = " " Then ea.RemoveChar (Pos + 1) ' there may be spaces after the arrow which can be distracting
        If Mid(ea.Text, Pos - 1, 1) = ")" Then
            ea.SetNormalBrackets
            s = ea.GetPrevBracketContent(Pos)
            ea.Text = Left(ea.Text, Pos - Len(s) - 3) & aString & s & bString & Right(ea.Text, Len(ea.Text) - Pos)
        ElseIf Mid(ea.Text, Pos - 1, 1) = "}" Then
            ea.SetTuborgBrackets
            s = ea.GetPrevBracketContent(Pos)
            pos1 = ea.Pos - 2
            Do While ea.IsLetterPos(pos1, False)
                pos1 = pos1 - 1
            Loop
            If ea.ChrByIndex(pos1) = "\" Then
                ea.Text = Left(ea.Text, pos1 - 1) & aString & Mid(ea.Text, pos1, Pos - pos1) & bString & Right(ea.Text, Len(ea.Text) - Pos - Len(Accent) + 1)
            End If
        Else
            ea.Text = Left(ea.Text, Pos - 2) & aString & Mid(ea.Text, Pos - 1, 1) & bString & Right(ea.Text, Len(ea.Text) - Pos - Len(Accent) + 1)
        End If
        Pos = InStr(Pos + 1, ea.Text, Accent)
    Loop
    ConvertAccent = ea.Text
End Function
Function ConvertPreAccent(Text As String, Accent As String, aString, bString) As String
    Dim ea As New ExpressionAnalyser
    Dim Pos As Long, s As String
    ea.Text = Text
    Do
        Pos = InStr(ea.Text, Accent)
        If Pos > 0 Then
            If ea.ChrByIndex(Pos + 1) = "(" Then
                ea.SetNormalBrackets
                s = ea.GetNextBracketContent(Pos)
                ea.Text = Left(ea.Text, Pos - 1) & aString & s & bString & Right(ea.Text, Len(ea.Text) - ea.Pos + 1)
            Else
                s = ea.GetNextVar(Pos)
                ea.Text = Left(ea.Text, Pos - 1) & aString & s & bString & Right(ea.Text, Len(ea.Text) - ea.Pos)
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
    If StopNow Then GoTo slut

    '    MaximaOutput = Replace(MaximaOutput, "%i", "i")
    ConvertToLatexMaxima = MaximaOutput
slut:
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
    Text = Replace(Text, "\cdot", VBA.ChrW(183))    'dot multiplication

    'greek letters
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
    Text = Replace(Text, "\eta", VBA.ChrW(951))    ' must be last as eta is included in others
    Text = Replace(Text, "\inc", VBA.ChrW(8710))  ' speciel delta increment
    '    text = Replace(text, "degC", VBA.ChrW(8451))   ' speciel oC symbol
    '    text = Replace(text, "Symangle", VBA.ChrW(8736))  '

    ea.SetTuborgBrackets
    ea.Text = Text
    p = InStr(ea.Text, "\frac")
    Do While p > 0
        t = ea.GetNextBracketContent
        n = ea.GetNextBracketContent
        ea.Text = Left(ea.Text, p - 1) & "(" & t & ")/(" & n & ")" & Right(ea.Text, ea.Length - ea.Pos + 1)
        p = InStr(ea.Text, "\frac")
    Loop

    ConvertLatexToWord = ea.Text

End Function
Sub ReadSelection()
' inserts the selected equation as a command
' this method does not use tempdoc, but changes the selected equation to text, reads it and changes it back to prof layout. This may cause changes in the way it is written.
' the method is probably a bit faster though

    Dim antal As Integer
    Dim sstart As Long
    Dim sslut As Long
    Dim j As Integer, antalgenstart As Integer
    On Error GoTo Fejl

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

        If InStr(Kommando, VBA.ChrW(8743)) Then
            Kommandoer = Split(Kommando, VBA.ChrW(8743))
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
            Kommandoer(j) = Replace(mathcol(j + 1).Range.Text, ChrW(11), "")    ' remove shift-enter
        Next
        
    Oundo.EndCustomRecord
    ActiveDocument.Undo
        
hop:
        ReDim Preserve Kommandoer(0 To antal - 1)
        sr.Select
    End If
    GoTo slut
Fejl:
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
    ReadEquation2 = OM.Range.Text
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
        If Len(Selection.Text) < 2 Then
            Kommando = ""
            MsgBox "Mark what should be evaluated by Maxima.", vbOKOnly, "Choose"
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
        Selection.Text = Replace(Selection.Text, VBA.ChrW(8727), MaximaGangeTegn)    'remove *
        Selection.Text = Replace(Selection.Text, "*", MaximaGangeTegn)    'remove *
        If DecSeparator = "," Then
            Selection.Text = Replace(Selection.Text, ".", ",")    'When comma is the decseparator, . can be changed to comma.
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
    If r.Text = VBA.ChrW(11) Then    ' if there is shift-enter at the end, replace with regular return
        r.Text = VBA.ChrW(13)
    End If
End Sub
Sub InsertMaximaOutput()
' inserts maxima output into word document
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

    If Len(MaximaOutput) > 1000 Then
        If MsgBox2(TT.A(896) & Len(MaximaOutput) & VbCrLfMac & TT.A(897), vbYesNo, TT.A(846)) = vbNo Then
            MaximaOutput = Left(MaximaOutput, 1000) & " " & TT.A(119)
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
    'Do While Left(comm, 1) = vbLf
    '    comm = Right(comm, Len(comm) - 1)
    'Loop
    ' trim linefeeds behind
    'Do While Right(comm, 1) = vbLf
    '    comm = Left(comm, Len(comm) - 1)
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
        Dim Pos As Integer
        Dim ea As New ExpressionAnalyser
        Pos = InStr(KommentarOutput, "Assumptions:")
        ea.Text = KommentarOutput
        ea.Pos = Pos
        assum = Trim(omax.ConvertToWordSymbols(Replace(ea.GetNextBracketContent, ",", " , ")))
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
Fejl:
slut:
End Sub

Function FindMaximaOutput(Text As String) As String
' picks output2 from the entire maxima output
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

    FindMaximaOutput = text2

    GoTo slut
Fejl:
    opos = InStr(Text, "(%i2)")
    text2 = Right(Text, Len(Text) - opos - 4)
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
Fejl:
slut:
End Sub
Sub GenerateOutputArray(Optional Text As String)
    Dim Output As String
    Dim nr As Integer
    Dim AntalCifre As Integer
    Dim Pos As Long
    Dim pos2 As Long
    Dim posb As Long
    Dim ArrIndex As Integer 'mac
    ArrIndex = 1 'mac

    On Error GoTo Fejl
#If Mac Then
#Else
    If Text = "" Then Text = MaxProc.LastMaximaOutput
#End If
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
            GoTo slut
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
'#If Mac Then ' SBCL
        If Mid(Text, Pos + 3, 1) = "i" Then
            ' next is input
            Minput(ArrIndex) = Minput(ArrIndex) & Output
        Else
'            Output = Replace(Trim(Output), VbCrLfMac, "") ' v. 1.29 this now only for output, as input can be errors, with multiple lines. Here it is important with arrow
            Output = Trim(Output) ' replace vbcrlfmac removed 1.29 for the '23,,' the arrow showing the error was not placed correctly on the next line
            ' next is output
            moutput(ArrIndex) = moutput(ArrIndex) & Output
            ArrIndex = ArrIndex + 1 'mac
        End If
        Pos = pos2
'        pos2 = InStr(pos + 3, text, vbLf & "(%")
      pos2 = FindNextio(Pos + 3, Text) - 1
    Loop Until slutnu
    nr = ArrIndex
    If moutput(nr) = "" And Len(Minput(nr)) < 6 Then
        nr = nr - 1
    End If

    '    ShowOutputArray ' for test

    GoTo slut
Fejl:
slut:
   If nr < 1 Then nr = 1
    ReDim Preserve Minput(1 To nr) As String
    ReDim Preserve moutput(1 To nr) As String
End Sub

Function FindNextio(StartPos As Long, Text As String) As Long
On Error GoTo Fejl
Dim Pos As Long, io As String
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
   
GoTo slut
Fejl:
   FindNextio = 0
'    MsgBox "Fejl " & Err.Number & " (" & Err.Description & ") i procedure FindNextio, linje " & Erl & ".", vbOKOnly Or vbCritical Or vbSystemModal, TT.error
slut:
End Function
'#If Mac Then
Function RemovePrintetBool(s As String) As String
' removes the true/false which can be printed in output on Mac
' mac auto. A true or false may be printed at the end.
    Dim Arr As Variant
    Dim t As String, ant As Integer, i As Integer
    On Error GoTo Fejl
    s = TrimR(Trim(s), vbLf)
    s = TrimR(Trim(s), vbCr)
If DllConnType = 2 Then
    Arr = Split(s, vbLf) ' wsh is a bit different
Else
    Arr = Split(s, VbCrLfMac)
End If
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
GoTo slut
Fejl:
    MsgBox2 "removeprintetbool error", vbOKOnly, TT.Error
slut:
    RemovePrintetBool = s
End Function
'#End If
Function GetMaximaUnitOutput(Optional Text As String) As String
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
'    Text = MaxProc.LastMaximaOutput
#Else
    If Text = vbNullString Then
        Text = MaxProcUnit.LastMaximaOutput
    End If
#End If
    ReDim MinputU(1 To 100) As String
    ReDim moutputU(1 To 100) As String
    Dim slutnu As Boolean
'    MsgBox text
   Pos = FindNextio(1, Text) - 1
   pos2 = FindNextio(Pos + 2, Text) - 1
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
        ElseIf nr < 1000 Then
            AntalCifre = 3
        Else
            AntalCifre = 4
        End If
        Output = Replace(Trim(Mid(Text, posb + 1, pos2 - Pos - 4 - AntalCifre)), vbCrLf, "")

        If Mid(Text, Pos + 3, 1) = "i" Then
            ' next is input
            MinputU(nr) = MinputU(nr) & Output
        Else
            ' next is output
            moutputU(nr) = moutputU(nr) & Output
        End If
        Pos = pos2
        pos2 = FindNextio(Pos + 2, Text) - 1
    Loop Until slutnu

    If moutputU(nr) = "" And Len(MinputU(nr)) < 6 Then
        nr = nr - 1
    End If

'    ReDim Preserve MinputU(1 To nr) As String
'    ReDim Preserve moutputU(1 To nr) As String

    '    ShowOutputArray ' til test
    GetMaximaUnitOutput = moutputU(nr)

    GoTo slut
Fejl:
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
Function FindVariable(Optional Text As String, Optional MaximaCode As Boolean = True) As Boolean
' finds variables in the text and puts them in Vars-string
    Dim Var As String
    Dim i As Integer
    Dim j As Integer
    Dim arrint As Variant
    AntalVars = 0
    Vars = ""
    IntegrationVars = ""
    If Text = vbNullString Then
        Text = Kommando
        If Text = "" Then
            On Error GoTo Fejl
            For j = 0 To UBound(Kommandoer)
                Text = Text & "+" & Kommandoer(j)    ' there must only be characters between that cannot be included in the variable
            Next
        End If
    End If
    Dim ea As New ExpressionAnalyser
    If MaximaCode Then
        ea.Text = CodeForMaxima(Text)
    Else
        ea.Text = Text
    End If
    If ea.Text = vbNullString Then ' fejl
        GoTo Fejl
    End If
    ea.Pos = 1
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
    Var = Replace(Var, vbCrLf, "")
    Do While Var <> ""
        If ea.ChrByIndex(ea.Pos + 1) <> "(" And ea.ChrByIndex(ea.Pos + 1) <> "'" And ea.ChrByIndex(ea.Pos + 1) <> VBA.ChrW(8289) And Not (InStr(DefString, Var & ":") = 1 Or InStr(DefString, "$" & Var & ":") > 0) And InStr(Var, "qkqg") < 1 Then
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
        ea.Pos = ea.Pos + 1
        Var = ea.GetNextVar()
    Loop
    
    FindVariable = True
    GoTo slut
Fejl:
    FindVariable = False
slut:
End Function
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
    ReDim MathOArr(100)
    Set DefColl = New Collection
    start = Selection.Range.start
    sslut = Selection.Range.End
    Set ra = ActiveDocument.Range
    ra.End = start
    For i = ra.OMaths.Count To 1 Step -1
        s = ActiveDocument.Range.OMaths(i).Range.Text
        If InStr(s, VBA.ChrW(8788)) > 0 Or InStr(s, VBA.ChrW(8797)) > 0 Or InStr(s, ":") > 0 Or InStr(s, VBA.ChrW(8801)) > 0 Then    ' only look at fields with : triple with def= and especially := characters
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
Fejl:
    DefFejl = True
    '    DefString = ""
    FindDefinitions = ""
slut:
    If FindDefinitions <> "" Then
        '            FindDefinitions = "[" & Left(FindDefinitions, Len(FindDefinitions) - 1) & "]$"
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
    Dim Arr As Variant
    Dim deflign As Boolean

    ea.SetNormalBrackets
    s = Replace(s, VBA.ChrW(12310), "")    ' special invisible brackets are removed
    s = Replace(s, VBA.ChrW(12311), "")    ' special invisible brackets are removed
    s = Replace(s, vbCrLf, "")    ' enter can accur
    s = Replace(s, vbCr, "")
    slc = LCase(s)
 
    If InStr(s, ChrW(8788)) > 0 Or InStr(s, ChrW(8797)) > 0 Or InStr(s, ":=") > 0 Or InStr(slc, "definer:") > 0 Or InStr(slc, "define:") > 0 Or InStr(slc, "ligning:") > 0 Or InStr(slc, "equation:") > 0 Or InStr(slc, "slet def") > 0 Or InStr(slc, "delete def") > 0 Or InStr(s, VBA.ChrW(8801)) > 0 Or InStr(slc, LCase(TT.A(69))) > 0 Then
        If VBA.LCase(Left(s, 8)) = "definer:" Then
            s = Right(s, Len(s) - 8)
        ElseIf VBA.LCase(Left(s, 7)) = "define:" Then
            s = Right(s, Len(s) - 7)
        ElseIf VBA.LCase(Left(s, 16)) = "definer ligning:" Then
            s = Right(s, Len(s) - 16)
            deflign = True
        ElseIf VBA.LCase(Left(s, 16)) = "define equation:" Then
            s = Right(s, Len(s) - 16)
            deflign = True
        End If
        s = Replace(s, VBA.ChrW(8788), ":=")    ' \equiv and
        s = Replace(s, VBA.ChrW(8797), ":=")    ' def=
        s = Replace(s, VBA.ChrW(8801), ":=")    ' speciel := symbol

        s = Replace(s, VBA.ChrW(11), "")
       slc = LCase(s)
 
        '        If InStr(s, "block") < 1 Then
        '            s = Replace(s, VBA.ChrW(34), "") ' apostrophe
        '            s = Replace(s, "=", ":")
        '            s = Replace(s, "::", ":")
        '            s = Replace(s, "):", "):=")
        '        End If

        '        s = Replace(s, "):=:", "):=")
        If InStr(slc, "slet def") > 0 Or InStr(slc, "delete def") > 0 Or InStr(slc, LCase(TT.A(69))) Then
            s = Replace(Trim(Split(s, ":")(1)), vbCr, "")
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
            ea.Text = Replace(s, ";", " , ")
            ea.ConvertDecSeparator
            ea.Pos = 1
            Do
                defenk = Trim(ea.GetNextListItem(ea.Pos, ","))
                If Len(defenk) = 0 Then GoTo hop
                If InStr(defenk, "block") > 0 Then
                    defenk = Replace(defenk, ";", ",")
                    defenk = Replace(defenk, VBA.ChrW(9633), "")    ' characters before parentheses that make them invisible. Possibly problematic as the parentheses are not removed. also makes fractions small
                    defenk = Replace(defenk, VBA.ChrW(9508), "")    ' blank space character at least [ ]
                    defenk = Replace(defenk, VBA.ChrW(11), "")    ' shift-enter
                    defenk = Replace(defenk, vbLf, "")    ' shift-enter and enter
                    defenk = Replace(defenk, vbCrLf, "")
                    defenk = Replace(defenk, vbCr, "")
                    defenk = Replace(defenk, VBA.ChrW(183), "*")    ' dot is replaced with times
                    defenk = Replace(defenk, VBA.ChrW(8901), "*")    ' \cdot , don't know why there seem to be two
                    defenk = Replace(defenk, VBA.ChrW(8729), "*")    ' \cdot replaced by multiply
                    defenk = Replace(defenk, VBA.ChrW(8226), "*")    ' thick dot replaced with times
                    defenk = Replace(defenk, VBA.ChrW(247), "/")
                    defenk = Replace(defenk, VBA.ChrW(8800), "#")    ' not equal to
                    defenk = Replace(defenk, VBA.ChrW(8804), "<=")
                    defenk = Replace(defenk, VBA.ChrW(8805), ">=")
                    defenk = Replace(defenk, VBA.ChrW(160), "")    ' before accent e.g. vector arrow 160 8407
                    defenk = Replace(defenk, VBA.ChrW(8289), "") ' function symbol
                    defenk = Replace(defenk, VBA.ChrW(9618), "") ' function symbol
                    defenk = Replace(defenk, VBA.ChrW(12310), "(") ' hidden brackets
                    defenk = Replace(defenk, VBA.ChrW(12311), ")") ' hidden brackets
                Else
                    defenk = Replace(defenk, VBA.ChrW(34), "")    ' apostrof
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
                Arr = Split(defenk, ":")
                If UBound(Arr) = 0 Then
                    '                    If InStr(defenk, "setunits") > 0 Or InStr(defenk, "assume") > 0 Then
                    GoTo hop2
                    '                    End If
                End If
                If Not ValidateDef(Arr(0), Arr(1)) Then GoTo hop

                If Left(Arr(1), 1) = "=" Then Arr(1) = Right(Arr(1), Len(Arr(1)) - 1)
                If Arr(1) = "" Then GoTo hop
            
                If UBound(ldefname) < defindex Then
                    ReDim Preserve ldefname(defindex)
                    ReDim Preserve ldefvalue(defindex)
                    ReDim Preserve MathOArr(defindex)
                End If
                ldefname(defindex) = Arr(0)
                ldefvalue(defindex) = Arr(1)
                Set MathOArr(defindex) = MathObj
                defindex = defindex + 1
hop2:                    ' if setunit or assume or declare
                ea2.Text = defenk
                ea2.Pos = 1
                t = ea2.GetNextVar

                If InStr(Split(defenk, ":")(0), "diff") > 0 Then
                    MsgBox TT.A(116) & s
                ElseIf (t = "e" And InStr(defenk, "[") <= 0) Or t = "integrate" Or t = "sin" Or t = "cos" Or t = "tan" Or t = "diff" Or t = "solve" Or t = "at" Or t = "for" Or t = "in" Or t = "do" Or t = "if" Or t = "and" Or t = "or" Or t = "product" Or t = "step" Or t = "next" Or t = "from" Or t = "limit" Or t = "sum" Or t = "then" Or t = "and" Or t = "or" Or t = "else" Or t = "unless" Or t = "while" Then
                    MsgBox TT.A(845) & " " & t & " " & TT.A(434), vbOKOnly, TT.Error
                ElseIf t = "setunits" Or t = "assume" Or t = "declare" Then
                    '                    AddDefinition =  AddDefinition & CodeForMaxima(defenk) & "$"
                    AddDefinition = AddDefinition & defenk & ";"
                ElseIf t <> "" Then
                    KillDef = KillDef & t & ","
                    AddDefinition = AddDefinition & defenk & ";"
                End If
hop:
            Loop Until ea.Pos >= Len(ea.Text)
        End If
    End If

End Function
Function SplitAssump(s As String) As String
' splits  0<x<5  to  x>0 and x<5
   Dim ea As New ExpressionAnalyser
   Dim v As String, p As Integer
   Dim pre As String, pst As String
   
   ea.SetNormalBrackets
   
    s = Replace(s, VBA.ChrW(8804), "<=")    '
    s = Replace(s, VBA.ChrW(8805), ">=")    '
    s = Replace(s, " ", vbNullString)
   ea.Text = s
   
   v = ea.GetNextVar()
   If v = "" Then
      SplitAssump = s
      Exit Function
   End If
   
   p = InStr(s, "<=" & v & "<=")
   If p > 0 Then
      pre = Left(s, p - 1)
      pst = Right(s, Len(s) - p - 4)
      SplitAssump = v & ">=" & pre & "&&and&&" & v & "<=" & pst
'      SplitAssump = pre & "<=" & v & " and " & v & "<=" & pst
      Exit Function
   End If
   p = InStr(s, "<=" & v & "<")
   If p > 0 Then
      pre = Left(s, p - 1)
      pst = Right(s, Len(s) - p - 3)
      SplitAssump = pre & "<=" & v & "&&and&&" & v & "<" & pst
      Exit Function
   End If
   p = InStr(s, "<" & v & "<=")
   If p > 0 Then
      pre = Left(s, p - 1)
      pst = Right(s, Len(s) - p - 3)
      SplitAssump = pre & "<" & v & "&&and&&" & v & "<=" & pst
      Exit Function
   End If
   p = InStr(s, "<" & v & "<")
   If p > 0 Then
      pre = Left(s, p - 1)
      pst = Right(s, Len(s) - p - 2)
      SplitAssump = pre & "<" & v & "&&and&&" & v & "<" & pst
      Exit Function
   End If
   SplitAssump = s
End Function
Function sqbrAssump(s As String) As String
   Dim ea As ExpressionAnalyser, t As String
   Dim v As String, p As Integer, c As String
   Dim Arr() As String, l1 As String, l2 As String
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
   s = Right(s, Len(s) - p)
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
      sqbrAssump = v & l1 & Arr(0) & "&&and&&" & Arr(1) & l2 & v
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
    Dim Pos As Integer
    On Error GoTo Fejl
    ValidateDef = False

    Pos = InStr(LHS, "[")    ' You can write anything in subscript/index
    If Pos > 0 Then
        LHS = Left(LHS, Pos - 1)
    End If

    ' If Not ea.IsAllText Then   not good enough then, and; not taken into account
    If InStr(LHS, "+") > 0 Or InStr(LHS, "-") > 0 Or InStr(LHS, "*") > 0 Or InStr(LHS, "/") > 0 Then
        MsgBox TT.A(117) & vbCrLf & vbTab & ConvertToAscii(LHS & "=" & RHS) & vbCrLf & TT.A(118), vbOKOnly, TT.Error
        Exit Function
    End If

    If Right(RHS, 1) = "+" > 0 Then
        MsgBox TT.A(117) & vbCrLf & vbTab & ConvertToAscii(LHS & "=" & RHS) & vbCrLf & TT.A(118), vbOKOnly, TT.Error
        Exit Function
    End If


    ValidateDef = True
Fejl:
End Function
Function GetAssumeText(defenk) As String
' takes text from word and converts to assume and inserts into forget
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
                                s = Left(s, p2 - 1) & "," & Right(s, Len(s) - p1 - 1)
                            Else
                                s = Left(s, p2 - 1) & "," & Right(s, Len(s) - p1)
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
' checks if the equation is completely solved for the variable
' starts it with variable= ?
' includes variable on the right side of = sign
    On Error GoTo slut
    Dim ea As New ExpressionAnalyser
    Dim Var As String
    Dim Arr As Variant
    If CASengine > 0 Then
        ea.SetTuborgBrackets
    End If
    Ligning = RemoveBrackets(Ligning)
    Arr = Split(Ligning, "=")

    If UBound(Arr) = 1 Then    ' there is exactly one = sign
        ea.Text = Arr(1)    'right side
        IsSolved = True
        Do
            Var = ea.GetNextVar
            ea.Pos = ea.Pos + 1
            If Var = variabel Then
                IsSolved = False    'variable found on the right hand side then it is not solved anyway
                GoTo slut
            End If
        Loop Until Var = ""

        If Left(Ligning, 1 + Len(variabel)) = variabel & "=" Then  '  not enough alone
            IsSolved = True
        ElseIf Len(variabel) > 1 And Mid(Ligning, Len(variabel) + 1, 1) = "=" Then    ' if length of variable same as left side. Necessary for variables like 1+r as it can be changed to r+1. But probably not 100% sure
            IsSolved = True
        Else    ' it must be assumed that it is not solved if it does not match anything above
            IsSolved = False
        End If
    Else
        IsSolved = False
    End If

    GoTo slut
Fejl:
    IsSolved = False
slut:
End Function

Function IsAllSolved(ByVal ligninger As String, ByVal variabel As String, Optional ByVal Sep As String = ",") As String
' checks if a list of equations is completely solved
' true if one is solved, false if none. All if all.
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
        If InStr(ligninger, "%if(") > 0 Then GoTo slut

        ligninger = RemoveBrackets(ligninger)
        Arr = Split(ligninger, Sep)

        For i = 0 To UBound(Arr)
            If IsSolved(Arr(i), variabel) Then
                oneSolved = True
            Else
                allsolved = False
            End If
        Next
        
        If InStr(ligninger, VBA.ChrW(8743)) > 0 Then ' there may be systems of equations if vector equations have been submitted
            Arr = Split(ligninger, VBA.ChrW(8743))
            For i = 0 To UBound(Arr)
                If IsSolved(Arr(i), variabel) Then
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
    On Error GoTo Fejl

    If Kommando <> "" Then
        KommandoerStreng = Replace(Kommando, vbCr, "")
    Else
        For i = 0 To UBound(Kommandoer)
            KommandoerStreng = KommandoerStreng & Kommandoer(i) & ListSeparator
        Next
        KommandoerStreng = Left(KommandoerStreng, Len(KommandoerStreng) - 1)
    End If

    GoTo slut
Fejl:
    KommandoerStreng = ""
slut:
End Property

Public Property Get AntalKom() As Integer
' number of equations/expressions
    On Error GoTo Fejl
    AntalKom = UBound(Kommandoer) + 1
    GoTo slut
Fejl:
    If Kommando <> "" Then
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

    Dim SettingsString As String, Arr() As String, i As Integer
    
    SettingsString = "["
    If MaximaUnits Then
        SettingsString = SettingsString & "keepfloat:false,setunits([N,J,W,Pa,C,V,F,Ohm,T,H,K]),"
        If Not OutUnits = vbNullString Then
            Arr = Split(OutUnits, ",")
            For i = 0 To UBound(Arr) ' It has to be done this way because you cannot run setunits on a list where the same type of unit appears twice. e.g. eV and J
                SettingsString = SettingsString & "setunits(" & omax.ConvertUnits(Arr(i)) & "),"
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
            
    If Right(SettingsString, 1) = "," Then SettingsString = Left(SettingsString, Len(SettingsString) - 1)
    SettingsString = SettingsString & "]$"
     
    GetMaximaSettingsString = SettingsString
End Function
