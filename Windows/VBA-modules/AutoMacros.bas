Attribute VB_Name = "AutoMacros"
Option Explicit
#If Mac Then
#Else
    Dim oAppClass As New oAppClass ' er ogs� i P, s� risiko for tabt tempdoc er mindre
    Private Declare PtrSafe Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As LongPtr, ByVal bInitialOwner As LongPtr, ByVal lpName As String) As LongPtr
#End If

Sub AutoExec()
Attribute AutoExec.VB_ProcData.VB_Invoke_Func = "WordMat.AutoMacros.AutoExec"
' denne k�res kun hvis filen er sat som globalskabelon. Alts� ikke hvis den bare �bnes

ChangeAutoHyphen ' s� 1-(-1) ikke overs�ttes til  1--1 t�nkestreg

#If Mac Then
#Else
    Set oAppClass.oApp = Word.Application
    CreateMutex 0&, 0&, "WordMatMutex"
#End If

'    LavRCMenu ' fjernet da det gav problemer med udskrivning af flere dokumenter p� engang
                ' det for�rsagede �ndringer i normal.dot. Nu rykket til preparemaxima
'    CustomizationContext = ActiveDocument.AttachedTemplate ' kan man ikke p� dette tidspunkt i opstart
SetAllDefaultRegistrySettings ' hvis ny bruger

ReadAllSettingsFromRegistry
AntalB = Antalberegninger

If AppVersion <> RegAppVersion Then ' hvis det er f�rste gang WordMat startes efter en opdatering, S� her kan s�ttes indstillinger der skal �ndres
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

 TriangleNAS = "A"
 TriangleNBS = "B"
 TriangleNCS = "C"
 TriangleNAV = "a"
 TriangleNBV = "b"
 TriangleNCV = "c"
 TriangleSett1 = 3
 TriangleSett2 = 2
 TriangleSett3 = False
 TriangleSett4 = False

'If AutoStart Then ' WordMat start ret hurtigt op nu. Der er ikke meget fordel ved denne. Den giver bare potentielt problemer.
'    PrepareMaximaNoSplash
'End If
End Sub

Sub AutoExit()
End Sub

'Sub AutoClose()
'' hver gang dokument lukkes, men kun n�r filen er �bnet - ikke n�r den er sat som global skabelonen. S� kaldes den slet ikke. S�dan burde det dog ikke v�re og det er det heller ikke med autoexec
'
''Dim d As Variant
'Exit Sub  ' n�dvendig n�r der er appclass?
'
'On Error Resume Next
''tempDoc.Close (False)
''LukTempDoc
''
''    For Each d In Application.Documents
''        If d.BuiltInDocumentProperties("Title") = "MMtempDoc" Then
''       d.Close (False)
''       End If
''    Next
''
'''    Set WebV = Nothing
'''    Set MaxProc = Nothing
'''    Set MaxProcUnit = Nothing
'
'
'    If Application.Documents.Count <= 2 Then
'        LukTempDoc
'        MaxProc.CloseProcess
'        cxl.CloseExcel
'
'#If Mac Then
'#Else
'        Set WebV = Nothing
'        Set MaxProc = Nothing
'        Set MaxProcUnit = Nothing
'#End If
'        '    For Each d In Application.Documents ' f�r Word til altid at sp�rge om der ikke skal gemmes
'        '       If d.BuiltInDocumentProperties("Title") = "MMtempDoc" Then
'        '           d.Close (False)
'        '       End If
'        '    Next
'        '    SletRCMenu
'    End If
'
'End Sub


