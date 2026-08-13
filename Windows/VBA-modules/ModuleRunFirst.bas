Attribute VB_Name = "ModuleRunFirst"
Option Explicit

Private HasStarted As Boolean
Private WMRunTime As Single

Dim oAppClass As New oAppClass ' is also in P, so the risk of lost tempdoc is less
#If Mac Then
#Else
    Private Declare PtrSafe Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As LongPtr, ByVal bInitialOwner As LongPtr, ByVal lpName As String) As LongPtr
#End If

Sub RunFirst()
      ' Should be run on startup of WordMat
          Dim s As String
          
10        If Abs(Timer() - WMRunTime) > 24# * 3600 Then
20            On Error Resume Next
30            Err.Clear
40            Application.Run macroname:="Popstart"
50            If Err.Number = -2147352573 Then
60                If TT.LangNo = 1 Then
70                    OpenLink "https://www.eduap.com/da/partnerskab/"
80                Else
90                    OpenLink "https://www.eduap.com/partnerskab/"
100               End If
110           End If
120           Err.Clear
130           On Error GoTo TheEnd
140           DoEvents
150           WMRunTime = Timer()
160       End If
          
170       If HasStarted Then Exit Sub
          
180       On Error Resume Next
190       Err.Clear
200       Application.Run macroname:="Popstart"
210       If Err.Number = -2147352573 Then
220           If TT.LangNo = 1 Then
230               OpenLink "https://www.eduap.com/da/partnerskab/"
240           Else
250               OpenLink "https://www.eduap.com/partnerskab/"
260           End If
270       End If
280       Err.Clear
290       On Error GoTo TheEnd
300       DoEvents
310       WMRunTime = Timer()
320       AntalB = Antalberegninger

330       SetMathAutoCorrect
340       ChangeAutoHyphen ' so 1-(-1) does not translate to 1--1 dash

350       Set oAppClass.oApp = Word.Application
#If Mac Then
#Else
360       CreateMutex 0&, 0&, "WordMatMutex"
#End If
          Dim RSF As Integer, SettingsLoadedOK As Boolean
370       RSF = ReadSettingsFromFile
380       If RSF > 0 Then
390           If RSF = 2 Then
400               SettingsLoadedOK = LoadSettingsFromData
410           ElseIf RSF = 3 Then
420               SettingsLoadedOK = LoadSettingsFromWMfolder
430           End If
440       End If
          
450       If Not SettingsLoadedOK Then
460           SetAllDefaultRegistrySettings ' if new user
470           ReadAllSettingsFromRegistry
480       End If
          

490       If AppVersion <> RegAppVersion Then ' if this is the first time WordMat is started after an update, then here you can set the settings that need to be changed
500           If val(RegAppVersion) <= 1.33 Then
510               SettShortcutAltM = KeybShortcut.InsertNewEquation
520               SettShortcutAltM2 = -1
530               SettShortcutAltB = KeybShortcut.beregnudtryk
540               SettShortcutAltL = KeybShortcut.SolveEquation
550               SettShortcutAltP = KeybShortcut.ShowGraph
560               SettShortcutAltD = KeybShortcut.Define
570               SettShortcutAltS = KeybShortcut.sletdef
580               SettShortcutAltF = KeybShortcut.Formelsamling
590               SettShortcutAltO = KeybShortcut.OmskrivUdtryk
600               SettShortcutAltR = KeybShortcut.PrevResult
610               SettShortcutAltJ = KeybShortcut.SettingsForm
620               SettShortcutAltN = -1
630               SettShortcutAltE = -1
640               SettShortcutAltT = KeybShortcut.ConvertEquationToLatex
650               SettShortcutAltQ = -1
660           End If
670           If val(RegAppVersion) <= 1.34 Then
680               OutputColor = wdGreen
690           End If
700           If val(RegAppVersion) < 1.37 Then
710               ShowAssum = True
720           End If
730           If val(RegAppVersion) < 1.4 Then
740               SettShortcutAltG = KeybShortcut.ShowGraph
750               If Not QActivePartnership Then
760                   s = GetRegSettingString("ShowMenus")
770                   s = "10" & Right(s, Len(s) - 2)
780                   SetRegSettingString "ShowMenus", s
790               End If
800           End If
810           RegAppVersion = AppVersion
820       End If
830       If SettCheckForUpdate Then CheckForUpdateSilent

840       GoTo slut
TheEnd:
850       MsgBox2 "A startup error occured. WordMat will probably work, but please show this to support: " & "Err. number: " & Err.Number & vbCrLf & Err.Description & vbCrLf & "Linenumber: " & Erl, vbOKOnly, TT.Error
slut:
860       HasStarted = True
End Sub

Sub SetMaxProc()
#If Mac Then
#Else
    If DllConnType > 1 Then Exit Sub ' not when using wsh
    
    If MaxProc Is Nothing Then
'        On Error Resume Next
        Err.Clear
        Set MaxProc = GetMaxProc() 'CreateObject("MaximaProcessClass")
        If Not MaxProc Is Nothing Then GetMaxProc.SetMaximaPath GetMaximaPath()
        If Err.Number <> 0 Then
            Err.Clear
            If QActivePartnership(False, True) Then
                If DllConnType = 0 Or DllConnType = 1 Then
                    If MsgBox2(TT.A(885), vbYesNo, TT.Error) = vbYes Then
                        DllConnType = 2
                    End If
                End If
            Else
                MsgBox2 TT.A(54), vbOKOnly, TT.Error
            End If
        End If
    End If
    
#End If

End Sub
