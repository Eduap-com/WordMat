Attribute VB_Name = "ModuleKeyboardShortcuts"
Option Explicit
Private TapTime As Single

Sub ExecuteKeyboardShortcut(ShortcutVal As KeybShortcut)
    RunFirst
    Select Case ShortcutVal
    Case KeybShortcut.InsertNewEquation
        NewEquation
    Case KeybShortcut.NewNumEquation
        InsertNumberedEquation
    Case KeybShortcut.beregnudtryk
        beregn
    Case KeybShortcut.SolveEquation
        MaximaSolve
    Case KeybShortcut.Define
        InsertDefiner
    Case KeybShortcut.sletdef
        InsertSletDef
    Case KeybShortcut.ShowGraph
        StandardPlot
    Case KeybShortcut.Formelsamling
        RunFirst
        Application.Run macroname:="WMPShowFormler"
    Case KeybShortcut.OmskrivUdtryk
        Omskriv
    Case KeybShortcut.SolveDiffEq
        SolveDE
    Case KeybShortcut.ExecuteMaximaCommand
        MaximaCommand
    Case KeybShortcut.PrevResult
        ForrigeResultat
    Case KeybShortcut.SettingsForm
        MaximaSettings
    Case KeybShortcut.ToggleNumExact
        ToggleNum
    Case KeybShortcut.ToggleUnitsOnOff
        ToggleUnits
    Case KeybShortcut.ConvertEquationToLatex
        ToggleLatex
    Case KeybShortcut.OpenLatexPDF
        SaveDocToLatexPdf
    Case KeybShortcut.InsertRefToEqution
        InsertEquationRef
    Case Else
        UserFormShortcuts.Show
    End Select
End Sub

Sub PressAltM()
    If SettShortcutAltM2 <> KeybShortcut.NoShortcut Then
        If Timer() - TapTime < 0.8 Then
            ExecuteKeyboardShortcut SettShortcutAltM2
            GoTo slut
        End If
        TapTime = Timer()
    End If
    
    ExecuteKeyboardShortcut SettShortcutAltM
slut:
End Sub
Sub PressAltB()
    ExecuteKeyboardShortcut SettShortcutAltB
End Sub
Sub PressAltL()
    ExecuteKeyboardShortcut SettShortcutAltL
End Sub
Sub PressAltP()
    ExecuteKeyboardShortcut SettShortcutAltP
End Sub
Sub PressAltD()
    ExecuteKeyboardShortcut SettShortcutAltD
End Sub
Sub PressAltS()
    ExecuteKeyboardShortcut SettShortcutAltS
End Sub
Sub PressAltF()
    ExecuteKeyboardShortcut SettShortcutAltF
End Sub
Sub PressAltO()
    ExecuteKeyboardShortcut SettShortcutAltO
End Sub
Sub PressAltR()
    ExecuteKeyboardShortcut SettShortcutAltR
End Sub
Sub PressAltJ()
    ExecuteKeyboardShortcut SettShortcutAltJ
End Sub
Sub PressAltN()
    ExecuteKeyboardShortcut SettShortcutAltN
End Sub
Sub PressAltE()
    ExecuteKeyboardShortcut SettShortcutAltE
End Sub
Sub PressAltT()
    ExecuteKeyboardShortcut SettShortcutAltT
End Sub
Sub PressAltQ()
    ExecuteKeyboardShortcut SettShortcutAltQ
End Sub
Sub PressAltG()
    ExecuteKeyboardShortcut SettShortcutAltG
End Sub
Sub PressAltGr()
    ExecuteKeyboardShortcut SettShortcutAltGr
End Sub


