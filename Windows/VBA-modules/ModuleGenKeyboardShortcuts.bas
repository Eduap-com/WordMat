Attribute VB_Name = "ModuleGenKeyboardShortcuts"
Option Explicit

Public Sub GenerateKeyboardShortcutsWordMat()
' stores KeyboardShortcuts in WordMat.dotm, but only if the wordMat.dotm file itself is opened
    GenerateKeyboardShortcutsPar False
End Sub
Public Sub GenerateKeyboardShortcutsPar(Optional NormalDotmOK As Boolean = False)
    Dim Wd As WdKey, WT As Template
    Dim GemT As Template
    
    Set GemT = CustomizationContext
    
    DeleteKeyboardShortcutsInNormalDotm
    
    Set WT = GetWordMatTemplate(NormalDotmOK)
    If WT Is Nothing Then
        MsgBox "The open template is not wordmat*.dotm", vbOKOnly, Sprog.Error
        GoTo slut
    End If
    
    CustomizationContext = WT
    
    KeyBindings.ClearAll

On Error Resume Next
'#If Mac Then
'    Wd = wdKeyControl
'#Else
    Wd = wdKeyAlt ' 1024 for windows, 2048 for mac
'#End If
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyG, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="PressAltG"

#If Mac Then
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyReturn, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="PressAltGr"
#Else
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyReturn, Wd, wdKeyControl), KeyCategory:=wdKeyCategoryCommand, Command:="PressAltGr"
#End If

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyB, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="PressAltB"

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyD, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="PressAltD"
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="PressAltF"
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyL, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="PressAltL"
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyS, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="PressAltS"
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyJ, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="PressAltJ"
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyP, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="PressAltP"
        
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyM, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="PressAltM"
        
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyR, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="PressAltR"
        
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyE, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="PressAltE"
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyO, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="PressAltO"
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyN, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="PressAltN"
        
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyT, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="PressAltT"
        
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyQ, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="PressAltQ"
            
slut:
    Set CustomizationContext = GemT

End Sub


Public Sub CheckKeyboardShortcuts()
' to manually check if everything is ok with ks
    CheckKeyboardShortcutsPar False
End Sub
Public Sub TestCheckKeyboardShortcutsNoninteractive()
    MsgBox CheckKeyboardShortcutsPar(True)
End Sub
Public Function CheckKeyboardShortcutsNoninteractive() As String
' is used by the test module to check if ks is set correctly. It is not important whether Normal-dotm is set.
    CheckKeyboardShortcutsNoninteractive = CheckKeyboardShortcutsPar(True)
End Function
Function CheckKeyboardShortcutsPar(Optional NonInteractive As Boolean = False) As String
' Checks if Keyboard shortcuts are saved correctly in WordMat.dotm. and if anything is saved in normal.dotm
    Dim WT As Template
    Dim KB As KeyBinding
    Dim GemT As Template, s As String
    Dim KeybInNormal As Boolean, KBerr As Boolean
    On Error GoTo slut
    Set GemT = CustomizationContext
        
    Set WT = GetWordMatTemplate(False)
    If WT Is Nothing Then
        CheckKeyboardShortcutsPar = "No template named wordmat*.dotm could be found." & vbCrLf
        If Not NonInteractive Then
            MsgBox "It doesn't look like you have opened wordmat.dotm, but it is running as a global template. Shortcuts are shown for" & ActiveDocument.AttachedTemplate & "", vbOKOnly, "No WordMat template"
            Set WT = ActiveDocument.AttachedTemplate
        Else
            GoTo slut
        End If
    End If
    
    CustomizationContext = NormalTemplate
    For Each KB In KeyBindings
        If KeyBindings.Count > 10 Then
#If Mac Then
            If KB.KeyString = "Option+B" Then
                KeybInNormal = True
                Exit For
            End If
#Else
            If KB.Command = "WordMat.Maxima.Beregn" Then
                KeybInNormal = True
                Exit For
            End If
#End If
        End If
    Next
    If KeybInNormal Then
        CheckKeyboardShortcutsPar = CheckKeyboardShortcutsPar & "Warning: WordMat keyboard shortcuts have been set in Normal.dotm" & vbCrLf
        If Not NonInteractive Then
            MsgBox "WordMat keyboard shortcuts are set in Normal.dotm", vbOKOnly Or vbInformation, "Warning"
            DeleteNormalDotm
        End If
        GoTo slut
    End If
    
    CustomizationContext = WT
    
    If Not NonInteractive Then
        s = "CustomizationContext:  " & CustomizationContext & VbCrLfMac
        If CustomizationContext = ActiveDocument.AttachedTemplate Then
            s = s & "It is an active document" & VbCrLfMac
        Else
            s = s & "It is a global template and not an active document." & VbCrLfMac
        End If
        s = s & vbCrLf
        s = s & "No. of keybindings: " & KeyBindings.Count & VbCrLfMac & VbCrLfMac
        s = s & "Keybindings:" & VbCrLfMac
    End If
    On Error Resume Next
    
    For Each KB In KeyBindings
        Err.Clear
        s = s & "  " & KB.KeyString & " ->" & KB.Command & VbCrLfMac
        If Err.Number > 0 Then
            s = s & "  ??? ->" & KB.Command & VbCrLfMac
            KBerr = True
        End If
    Next
    
    If Not NonInteractive Then
        MsgBox s, vbOKOnly, "KeyBindings"
    ElseIf KeyBindings.Count < 10 Then
        CheckKeyboardShortcutsPar = CheckKeyboardShortcutsPar & "There is only " & KeyBindings.Count & " keyboard shortcuts in WordMat*.dotm. You should run GenerateKeyboardShortcutsWordMat." & vbCrLf
    ElseIf KBerr Then
        CheckKeyboardShortcutsPar = CheckKeyboardShortcutsPar & "There are problems with the keyboard shortcuts in WordMat*.dotm. You should run GenerateKeyboardShortcutsWordMat on Mac." & vbCrLf
    End If
    
slut:
    On Error Resume Next
    If Not GemT Is Nothing Then CustomizationContext = GemT

End Function

