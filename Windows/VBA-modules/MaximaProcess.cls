VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MaximaProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' denne klasse benyttes ikke mere efter Maxima tilgås med terminal på Mac. Den er der kun for at der ikke kommer fejlkompilering
#If Mac Then
'Private Declare PtrSafe Function StartMaxima Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/libMaximaConnection.dylib" () As LongPtr
'Private Declare PtrSafe Function CloseMaxima Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/libMaximaConnection.dylib" () As LongPtr
'Private Declare PtrSafe Function SetMaximaPath Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/libMaximaConnection.dylib" (ByVal maximapath As String) As LongPtr
'Private Declare PtrSafe Function IsMaximaRunning Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/libMaximaConnection.dylib" () As LongPtr
'Private Declare PtrSafe Function SendToMaxima Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/libMaximaConnection.dylib" (ByVal s As String) As LongPtr
''Private Declare PtrSafe Function SendToMaximaWait Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/libMaximaConnection.dylib" (ByVal s As String, ByVal outstring As String, ByVal waittime As LongPtr) As LongPtr
'Private Declare PtrSafe Function ReadOutputIfAny Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/libMaximaConnection.dylib" (ByVal outstring As String) As LongPtr

'Private Declare PtrSafe Function InString Lib "Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/libMaximaConnection.dylib" (ByVal startpos As Integer, ByVal searchstring As String, ByVal findstring As String) As LongPtr
'Private Declare PtrSafe Function CountInString Lib "Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/libMaximaConnection.dylib" (ByVal startpos As Integer, ByVal searchstring As String, ByVal findstring As String) As Integer
Private Declare PtrSafe Function system Lib "/usr/lib/libc.dylib" (ByVal Command As String) As Long
Private Declare PtrSafe Function Sleep Lib "/usr/lib/libc.dylib" Alias "sleep" (ByVal seconds As Long) As Long
Private Declare PtrSafe Function uSleep Lib "/usr/lib/libc.dylib" Alias "usleep" (ByVal useconds As Long) As Long

Private mUnits As Integer
Private mExact As Integer
Private mQuestion As Integer
Private mComplex As Integer
Private lastOutput As String
Private mFinished As Integer
Private mError As Integer
Private mAntalCifre As Integer
Private mIndex As LongPtr

Sub StartMaximaProcess(Optional WaitTime As LongPtr = 15)
    Dim r As LongPtr
    mFinished = 0
    mQuestion = 0
'    If IsMaximaRunning() > 0 Then
''        MsgBox "Maxima already running"
'        Exit Sub
'    End If
'    SetMaximaPath "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/Maxima.app/Contents/Resources/maxima.sh"
'    SetMaximaPath "/Applications/WordMat/Maxima.app/Contents/Resources/maxima.sh"
'    If WaitTime = 0 Then
'        r = StartMaxima()
'    Else
'        r = StartMaximaWait(WaitTime)
    ' Mac: pga ændring af operator _, som ikke  virker på Mac. Skal vel være i image. Er nu puttet i image v. 1.03.
'        Dim Command As String
'        Command = "infix(""##"",115,115)$""##""(qwa,qwb):=block(if scalarp(qwa) or scalarp(qwb) then return(qwa*qwb) else return(qwa.qwb))$" & Command
'        ExecuteMaximaCommand Command, 5
'    End If
    mIndex = 1
        If r < 1 Then
            MsgBox "Error starting Maxima", vbOKOnly, Sprog.Error
            GoTo Slut
        End If
Slut:
End Sub
Sub CloseProcess()
  '  CloseMaxima
End Sub

Function StartMaximaWait(WaitTime As LongPtr) As LongPtr
    Dim L As LongPtr, i As LongPtr
    Dim OutputBuffer As String
    Dim Startupstring As String
    Dim retval As LongPtr

    i = WaitTime * 20
    OutputBuffer = GetEmptyString()
'    retval = StartMaxima()

    If retval < 1 Then
'        StartMaximaWait = retval
        Exit Function
    End If

'    Do While IsMaximaRunning() < 1
'        uSleep 50000
'    Loop
'    Do
'        uSleep 50000
'        L = ReadOutputIfAny(OutputBuffer)
'        If L > 0 Then
'            Startupstring = Startupstring & Left(OutputBuffer, CInt(L))
'        End If
'        i = i - 1
'    Loop While i > 0 And InStr(Startupstring, "(%i1)") < 1  'l < 1 And

'    MsgBox Startupstring
    If InStr(Startupstring, "(%i1)") > 0 Then
        StartMaximaWait = 1
        mFinished = 1
    Else
        StartMaximaWait = retval
    End If
End Function

Sub ExecuteMaximaCommand(Command As String, WaitForIt As Integer)
Dim L As LongPtr, nocmd As LongPtr
Dim ea As New ExpressionAnalyser
On Error GoTo Fejl
    mFinished = 0
    mQuestion = 0
    Command = "[" & GuiSettingsString & "]$" & Command & vbLf

    lastOutput = ""
    
    ea.Text = Command
    nocmd = ea.CountText(";") + ea.CountText("$")
    mIndex = mIndex + nocmd
        
'    If WaitForIt = 0 Then
'        L = SendToMaxima(Command)
'    Else
'        L = SendToMaxima(Command)
'        WaitForMaximaUntil WaitForIt
        
'        lastOutput = GetEmptyString()
'        l = SendToMaximaWait(command, lastOutput, WaitForIt)
'        If l > 0 Then
'            lastOutput = Left(lastOutput, l)
'            mFinished = 1
'        End If
'        CheckForQuestion
'    End If
GoTo Slut
Fejl:
    MsgBox "Maxima execute error", vbOKOnly, Sprog.Error
Slut:
End Sub
Sub AnswerQuestion(answer As String)
' check c#
    ExecuteMaximaCommand answer, 0
End Sub
Private Sub CheckForQuestion()
' skal udbygges se i c# kode
    If InStr(lastOutput, "?") Then
        mQuestion = 1
    End If
End Sub
Public Function LastMaximaOutput() As String
    LastMaximaOutput = lastOutput
End Function
Function GuiSettingsString() As String

If Complex Then
    GuiSettingsString = "domain:complex,realonly:false,"
Else
    GuiSettingsString = "domain:real,"
End If
If Exact = 2 Then
    GuiSettingsString = GuiSettingsString & "dectalall(udt):=dectalallNum(udt),"
Else
    GuiSettingsString = GuiSettingsString & "dectalall(udt):=block([qout],qout:errcatch(apply1(udt,floatrul)),if qout=[] then udt else qout[1]),"
End If
GuiSettingsString = GuiSettingsString & "fpprintprec:" & mAntalCifre & ",NoSigFig:" & mAntalCifre ' fpprintprec er mac spec

End Function
Function WaitForMaximaUntil(Optional StopTime As Integer = 6) As Integer
' tid angives i hele  s
' -1 = error     0=timeout    1=maxima finished
Dim i As Integer
Dim L As LongPtr
Dim stoptext As String
Dim OutputBuffer As String
Dim t As String
On Error GoTo Slut
    stoptext = "(%i" & mIndex & ")"
    StopTime = StopTime * 10
If StopTime = 0 Then StopTime = 1
    mQuestion = 0
    mError = 0
'    If showGUI Then
        'Mac: Extra parameter showGUI
'        If MaxProc.Finished = 0 And MaxProc.ErrCode = 0 And Not (omax.StopNow) And i < StopTime Then
'            Dim ufwaitstartup As New MacUserFormWaitStartup
'            ufwaitstartup.StopTime = StopTime
'            ufwaitstartup.Show
'        End If
'    Else
WaitForMaximaUntil = 1
    Do While Not (omax.StopNow) And i < StopTime
        OutputBuffer = GetEmptyString()
'        L = ReadOutputIfAny(OutputBuffer)
'        l = GetLastOutputVB(OutputBuffer, 10)
        If L > 0 Then
            t = Left(OutputBuffer, CInt(L))
            lastOutput = lastOutput & t
        End If
'        msgbox stoptext & vbLf & vbLf & lastOutput
        If InStr(lastOutput, stoptext) > 0 Then
            mFinished = 1
            GoTo Slut
        ElseIf InStr(lastOutput, "Console interrupt.") > 0 Then
            mFinished = 1
            mError = 1
            GoTo Slut
        ElseIf InStr(lastOutput, "Acceptable answers are") > 0 Then
            mFinished = 1
            mQuestion = 1
            GoTo Slut
        ElseIf InStr(lastOutput, "Nonalgebraic argument given to 'topoly'") > 0 Then
            mFinished = 1
            mError = 1
            GoTo Slut
        ElseIf InStr(lastOutput, "log: log(0) has been generated.") > 0 Then
            mFinished = 1
            mError = 1
            GoTo Slut
        ElseIf InStr(lastOutput, "incorrect syntax:") > 0 Then
            mFinished = 1
            mError = 1
            GoTo Slut
        ElseIf InStr(lastOutput, "positive, negative or zero?") > 0 Or InStr(lastOutput, "positive or negative?") > 0 Then
            mFinished = 1
            mQuestion = 1
            GoTo Slut
        ElseIf InStr(lastOutput, "variables to solve for?") > 0 Then
            mFinished = 1
            mError = 1
            GoTo Slut
        ElseIf InStr(lastOutput, "syntax error") > 0 Then
            mFinished = 1
            GoTo Slut
        End If
        i = i + 1
        uSleep 100000 ' vent 1/10 s
    Loop
    WaitForMaximaUntil = 0 ' timeout
    If mError > 0 Then
        WaitForMaximaUntil = -1 'error
        MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
    End If
'    End If
GoTo Slut
Fejl:
        WaitForMaximaUntil = -1 'error
Slut:
End Function

Public Sub Reset(extrakommando As String)
' bruges til at køre killdefs
Dim Command As String

    If extrakommando <> vbNullString Then
        Command = extrakommando & ","
    End If
'    command = "[" & command & "reset(linenum,%,features,%rnum)," & GuiSettingsString & "]$"
    Command = "[" & Command & "reset(linenum,features,%rnum)]$"
    mIndex = 0 '    SetCommandIndex 0
    
    ExecuteMaximaCommand Command, 1 ' ændret fra 5
End Sub
Function GetEmptyString() As String
'    GetEmptyString = String(2000, vbNullChar)  'this will make strlen not work in c
    GetEmptyString = String(2000, " ")
End Function

Public Property Get Units() As Integer
    Units = mUnits
End Property
Public Property Let Units(ByVal vNewValue As Integer)
    mUnits = vNewValue
End Property
Public Property Get Exact() As Integer
    Exact = mExact
End Property
Public Property Let Exact(ByVal vNewValue As Integer)
    mExact = vNewValue
End Property
Public Property Get Complex() As Integer
    Complex = mComplex
End Property
Public Property Let Complex(ByVal vNewValue As Integer)
    mComplex = vNewValue
End Property
Public Property Get ErrCode() As Integer
    ErrCode = 0
End Property
Public Property Get Finished() As Integer
    Finished = mFinished
End Property
Public Property Let AntalCifre(ByVal vNewValue As Integer)
    mAntalCifre = vNewValue
End Property
Public Property Get Question() As Integer
    Question = mQuestion
End Property
Public Property Let Question(ByVal vNewValue As Integer)
    mQuestion = vNewValue
End Property
Public Function MaximaStarted() As Boolean
    Dim i As LongPtr
'    i = IsMaximaRunning()
    If i = 0 Then
        MaximaStarted = False
    Else
        MaximaStarted = True
    End If
End Function

Public Function Commandindex() As LongPtr
    Commandindex = mIndex
End Function
Public Function SetCommandindex(mi As LongPtr)
     mIndex = mi
End Function

#End If

