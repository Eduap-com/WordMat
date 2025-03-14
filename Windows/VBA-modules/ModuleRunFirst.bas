Attribute VB_Name = "ModuleRunFirst"
Option Explicit

Private HasStarted As Boolean

#If Mac Then
#Else
    Dim oAppClass As New oAppClass ' er også i P, så risiko for tabt tempdoc er mindre
    Private Declare PtrSafe Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As LongPtr, ByVal bInitialOwner As LongPtr, ByVal lpName As String) As LongPtr
#End If

Sub RunFirst()
' Should be run on startup of WordMat
If HasStarted Then Exit Sub

ChangeAutoHyphen ' så 1-(-1) ikke oversættes til  1--1 tænkestreg

#If Mac Then
#Else
    Set oAppClass.oApp = Word.Application
    CreateMutex 0&, 0&, "WordMatMutex"
#End If

SetAllDefaultRegistrySettings ' hvis ny bruger

ReadAllSettingsFromRegistry
AntalB = Antalberegninger

If AppVersion <> RegAppVersion Then ' hvis det er første gang WordMat startes efter en opdatering, Så her kan sættes indstillinger der skal ændres
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

HasStarted = True

End Sub

Sub SetMaxProc()
#If Mac Then
#Else
    If MaxProc Is Nothing Then
'        On Error Resume Next
        Err.Clear
        Set MaxProc = GetMaxProc() 'CreateObject("MaximaProcessClass")
        If Not GetMaxProc Is Nothing Then GetMaxProc.SetMaximaPath GetMaximaPath()
        If Err.Number <> 0 Then
            Err.Clear
            If QActivePartnership(False, True) Then
                If DllConnType = 0 Or DllConnType = 1 Then
                    If MsgBox2("Kan ikke forbinde til Maxima. Vil du anvende metoden 'WSH' i stedet?" & VbCrLfMac & VbCrLfMac & "(Denne indstilling findes under avanceret i Indstillinger)", vbYesNo, Sprog.Error) = vbYes Then
                        DllConnType = 2
                    End If
'                    If MsgBox2("Kan ikke forbinde til Maxima. Vil du anvende metoden 'dll direct' i stedet?" & VbCrLfMac & VbCrLfMac & "(Denne indstilling findes under avanceret i Indstillinger)", vbYesNo, Sprog.Error) = vbYes Then
'                        DllConnType = 1
'                    End If
'                ElseIf DllConnType = 1 Then
                End If
            Else
                MsgBox2 Sprog.A(54), vbOKOnly, Sprog.Error
            End If
        End If
    End If
#End If

End Sub
