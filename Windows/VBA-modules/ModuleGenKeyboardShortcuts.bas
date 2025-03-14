Attribute VB_Name = "ModuleGenKeyboardShortcuts"
Option Explicit

Public Sub GenerateKeyboardShortcutsWordMat()
' gemmer KeyboardShortcuts i WordMat.dotm, men kun hvis det er selve wordMat.dotm filen der er åbnet
    GenerateKeyboardShortcutsPar False
End Sub
Public Sub GenerateKeyboardShortcutsNormalDotm()
' gemmer KeyboardShortcuts i WordMat.dotm, hvis det er selve wordMat.dotm filen der er åbnet. Hvis ikke gemmes i normal.dotm.
' Det kan give problemer ved en opdatering at benytte denne metode
    GenerateKeyboardShortcutsPar True
End Sub
Public Sub GenerateKeyboardShortcutsPar(Optional NormalDotmOK As Boolean = False)
    Dim Wd As WdKey, WT As Template
    Dim GemT As Template
    
    Set GemT = CustomizationContext
    
    DeleteKeyboardShortcutsInNormalDotm
    
    Set WT = GetWordMatTemplate(NormalDotmOK)
    If WT Is Nothing Then
'        If MsgBox("Der kunne ikke findes nogen skabelon der hed wordmat*.dotm. Vil du anvende " & ActiveDocument.AttachedTemplate & "?", vbYesNo, "Ingen WordMat skabelon") = vbYes Then
'            Set WT = ActiveDocument.AttachedTemplate
'        Else
'            GoTo slut
'        End If
        MsgBox "Den åbne skabelon er ikke wordmat*.dotm", vbOKOnly, "Fejl"
        GoTo slut
    End If
    
    CustomizationContext = WT
    
    KeyBindings.ClearAll

On Error Resume Next
'#If Mac Then
'    Wd = wdKeyControl
'#Else
    Wd = wdKeyAlt ' 1024 på windows, 2048 på mac
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
' til manuelt kald af om alt er ok med ks
    CheckKeyboardShortcutsPar False
End Sub
Public Sub TestCheckKeyboardShortcutsNoninteractive()
    MsgBox CheckKeyboardShortcutsPar(True)
End Sub
Public Function CheckKeyboardShortcutsNoninteractive() As String
' bruges at test-modulet til at checke om ks er sat rigtigt. Det er ikke vigtigt om Normal-dotm er sat.
    CheckKeyboardShortcutsNoninteractive = CheckKeyboardShortcutsPar(True)
End Function
Function CheckKeyboardShortcutsPar(Optional NonInteractive As Boolean = False) As String
    ' Checker om Keyboard shortcuts er gemt correct i WordMat.dotm.  og om der er gemt noget i normal.dotm
    Dim WT As Template
    Dim KB As KeyBinding
    Dim GemT As Template, s As String
    Dim KeybInNormal As Boolean, KBerr As Boolean
    On Error GoTo slut
    Set GemT = CustomizationContext
        
    Set WT = GetWordMatTemplate(False)
    If WT Is Nothing Then
        CheckKeyboardShortcutsPar = "Der kunne ikke findes nogen skabelon, der hed wordmat*.dotm" & vbCrLf
        If Not NonInteractive Then
            MsgBox "Det ser ikke ud til at du har åbnet wordmat.dotm, men kører som global skabelon. Genveje vises for " & ActiveDocument.AttachedTemplate & "", vbOKOnly, "Ingen WordMat skabelon"
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
        CheckKeyboardShortcutsPar = CheckKeyboardShortcutsPar & "Advarsel: Der er sat WordMat tastaturgenveje i Normal.dotm" & vbCrLf
        If Not NonInteractive Then
            MsgBox "Der er sat WordMat tastaturgenveje i Normal.dotm", vbOKOnly Or vbInformation, "Advarsel"
            DeleteNormalDotm
        End If
        GoTo slut
    End If
    
    CustomizationContext = WT
    
    If Not NonInteractive Then
        s = "CustomizationContext:  " & CustomizationContext & VbCrLfMac
        If CustomizationContext = ActiveDocument.AttachedTemplate Then
            s = s & "Det er aktivt dokument" & VbCrLfMac
        Else
            s = s & "Det er global skabelon og ikke aktivt dokument" & VbCrLfMac
        End If
        s = s & vbCrLf
        s = s & "Antal keybindings: " & KeyBindings.Count & VbCrLfMac & VbCrLfMac
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
        CheckKeyboardShortcutsPar = CheckKeyboardShortcutsPar & "Der er kun " & KeyBindings.Count & " tastaturveje i WordMat*.dotm. Der skal nok køres GenerateKeyboardShortcutsWordMat." & vbCrLf
    ElseIf KBerr Then
        CheckKeyboardShortcutsPar = CheckKeyboardShortcutsPar & "Der er problemer med Genvejene i WordMat*.dotm. Der skal nok køres GenerateKeyboardShortcutsWordMat på Mac." & vbCrLf
    End If
    
slut:
    On Error Resume Next
    If Not GemT Is Nothing Then CustomizationContext = GemT

End Function

