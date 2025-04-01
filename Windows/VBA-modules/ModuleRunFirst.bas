Attribute VB_Name = "ModuleRunFirst"
Option Explicit

Private HasStarted As Boolean

#If Mac Then
#Else
    Dim oAppClass As New oAppClass ' is also in P, so the risk of lost tempdoc is less
    Private Declare PtrSafe Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As LongPtr, ByVal bInitialOwner As LongPtr, ByVal lpName As String) As LongPtr
#End If

Sub RunFirst()
    ' Should be run on startup of WordMat
    If HasStarted Then Exit Sub

    ChangeAutoHyphen ' so 1-(-1) does not translate to 1--1 dash

#If Mac Then
#Else
    Set oAppClass.oApp = Word.Application
    CreateMutex 0&, 0&, "WordMatMutex"
#End If
    Dim RSF As Integer
    RSF = ReadSettingsFromFile
    If RSF > 0 Then
        If RSF = 2 Then
            LoadSettingsFromData
        ElseIf RSF = 3 Then
            LoadSettingsFromWMfolder
        End If
    End If

    SetAllDefaultRegistrySettings ' if new user

    ReadAllSettingsFromRegistry
    AntalB = Antalberegninger

    If AppVersion <> RegAppVersion Then ' if this is the first time WordMat is started after an update, then here you can set the settings that need to be changed
        If val(AppVersion) = 1.34 Then
            SettShortcutAltM = KeybShortcut.InsertNewEquation
            SettShortcutAltM2 = -1
            SettShortcutAltB = KeybShortcut.beregnudtryk
            SettShortcutAltL = KeybShortcut.SolveEquation
            SettShortcutAltP = KeybShortcut.ShowGraph
            SettShortcutAltD = KeybShortcut.Define
            SettShortcutAltS = KeybShortcut.sletdef
            SettShortcutAltF = KeybShortcut.Formelsamling
            SettShortcutAltO = KeybShortcut.OmskrivUdtryk
            SettShortcutAltR = KeybShortcut.PrevResult
            SettShortcutAltJ = KeybShortcut.SettingsForm
            SettShortcutAltN = -1
            SettShortcutAltE = -1
            SettShortcutAltT = KeybShortcut.ConvertEquationToLatex
            SettShortcutAltQ = -1
        End If
        RegAppVersion = AppVersion
    End If

    If SettCheckForUpdate Then CheckForUpdateSilent

    HasStarted = True

TheEnd:
End Sub

Sub SetMaxProc()
#If Mac Then
#Else
    If MaxProc Is Nothing Then
'        On Error Resume Next
        Err.Clear
        Set MaxProc = GetMaxProc() 'CreateObject("MaximaProcessClass")
        If Not MaxProc Is Nothing Then GetMaxProc.SetMaximaPath GetMaximaPath()
        If Err.Number <> 0 Then
            Err.Clear
            If QActivePartnership(False, True) Then
                If DllConnType = 0 Or DllConnType = 1 Then
                    If MsgBox2(Sprog.A(885), vbYesNo, Sprog.Error) = vbYes Then
                        DllConnType = 2
                    End If
                End If
            Else
                MsgBox2 Sprog.A(54), vbOKOnly, Sprog.Error
            End If
        End If
    End If
#End If

End Sub
