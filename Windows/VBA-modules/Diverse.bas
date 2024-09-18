Attribute VB_Name = "Diverse"
Option Explicit
Public TimeText As String
Public cxl As CExcel
Public ProgramFilesDir
Public DocumentsDir
Dim SaveTime As Single
Dim BackupAnswer As Integer
Private UserDir As String
Private tmpdir As String
Private TapTime As Single
#If Mac Then
    Private m_tempDoc As Document
#Else
Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
#End If

'Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

'Private Declare Function FindWindow Lib "user32.dll" _
'Alias "FindWindowA" ( _
'ByVal lpClassName As String, _
'ByVal lpWindowName As String) As Long

'Sub LockWindow()
''To turn it on just call it like this, passing it the hWnd of the window to lock.
''Dim nHwnd As Long
''MsgBox Application.ActiveDocument.Windows(1).Caption
''nHwnd = FindWindow("OpusApp", Application.Caption)
'nHwnd = FindWindow("OpusApp", Application.ActiveDocument.Windows(1).Caption)
''nHwnd = FindWindow("OpusApp", "")
''nHwnd = FindWindow("", Application.ActiveDocument.Windows(1).Caption)
'LockWindowUpdate nHwnd
'End Sub
'Sub UnLockWindow()
''To turn it off just call it and pass it a zero.
'LockWindowUpdate 0
'End Sub
'Sub TestLock()
'LockWindow
'Wait (5)
'UnLockWindow
'End Sub
Sub Tstmsg()
    MsgBox2 "det er" & vbCrLf & "test" & vbCrLf & "test" & vbCrLf & "test" & vbCrLf & "test" & vbCrLf & "test" & vbCrLf & "test", vbOKCancel, "d"
End Sub
#If Mac Then
#Else
Sub UnitImageTest()
    Dim i As Integer
'    MsgBox omax.ConvertMaximaUnits("2+3")
    
        PrepareMaxima
'    If MaxProcUnit Is Nothing Then
'        Set MaxProcUnit = CreateObject("MaximaProcessClass")
'        MaxProcUnit.Units = 1
'        MaxProcUnit.StartMaximaProcess
'    Do While MaxProcUnit.Finished = 0 And MaxProcUnit.ErrCode = 0 And i < 30
'        Wait (0.1)
'        i = i + 1
'    Loop
'    End If

    'fejl;2*km+200*m,numer:false;
    MaxProcUnit.ExecuteMaximaCommand "fejl;applyunitrule([t=11.7*(J/W)]);", 0
    Do While MaxProcUnit.Finished = 0 And MaxProcUnit.ErrCode = 0 And i < 30
        Wait (0.1)
        i = i + 1
    Loop

'    Do While MaxProcUnit.Finished = 0 And MaxProcUnit.ErrCode = 0 And i < 30
'        Wait (0.1)
'        i = i + 1
'    Loop
    
    MsgBox MaxProcUnit.LastMaximaOutput
End Sub
#End If

Function FileExists(FullFileName As String) As Boolean
' returns TRUE if the file or folder exists
    On Error GoTo Err
    FileExists = False
    FileExists = Len(Dir(FullFileName)) > 0 Or Len(Dir(FullFileName, vbDirectory)) > 0
    Exit Function
Err:
End Function

Function GetTempDir() As String
#If Mac Then
'    If startupdrive = vbNullString Then getappdir
    If UserDir = vbNullString Then UserDir = MacScript("return POSIX path of (path to home folder)")
'    If userdir = vbNullString Then userdir = MacScript("return (path to home folder) as string")
    tmpdir = UserDir & "WordMat/"
    On Error Resume Next
    MkDir (tmpdir)
    GetTempDir = tmpdir
#Else
    GetTempDir = Environ("TEMP") & "\"
#End If

End Function
Sub OpretTempdoc()
' men kun hvis ikke eksisterer allerede
'#If Mac Then
'    Call tempDoc
'#Else

    Exit Sub
' indtil v1.30:

Dim D As Document
If tempDoc Is Nothing Then
For Each D In Application.Documents
    If D.BuiltInDocumentProperties("Title") = "MMtempDoc" Then
        Set tempDoc = D
        Exit For
    End If
Next
End If

If tempDoc Is Nothing Then
    Set tempDoc = Documents.Add(, , , False)
'    tempDoc.ActiveWindow.View.Draft = True ' giver måske en hastighedsforbedring, men har ikke kunnet måle det
    tempDoc.BuiltInDocumentProperties("Title") = "MMtempDoc"
End If

If Not tempDoc.BuiltInDocumentProperties("Title") = "MMtempDoc" Then
    tempDoc.Close SaveChanges:=wdDoNotSaveChanges
    Set tempDoc = Documents.Add(, , , False)
    tempDoc.BuiltInDocumentProperties("Title") = "MMtempDoc"
End If
'#End If
End Sub
Sub ShowTempDoc()
    MsgBox tempDoc.Range.text
End Sub
Sub LukTempDoc()
On Error GoTo slut
'#If Mac Then ' fjernet 1.29
'    If Not m_tempDoc Is Nothing Then
'        If Word.Application.IsObjectValid(m_tempDoc) Then
'            m_tempDoc.Close (False)
'        End If
'    End If
'    Set m_tempDoc = Nothing
'#Else
    tempDoc.Close False
    Set tempDoc = Nothing ' added v. 1.11
'#End If
slut:
End Sub

'#If Mac Then ' fjernet 1.29 håndteres nu ens på mac og windows
'Function tempDoc() As Document
'    'Mac: User may have closed the document, so the variable tempDoc is now a function
'    Dim farRight As Integer
'    On Error Resume Next
''    farRight = ScreenWidth - 1 ' just inside the screen hides most
'    If Not m_tempDoc Is Nothing Then
'        If Word.Application.IsObjectValid(m_tempDoc) Then
''            If m_tempDoc.ActiveWindow.Left <> farRight Then m_tempDoc.ActiveWindow.Left = farRight
'            Set tempDoc = m_tempDoc
'            Exit Function
'        Else
'            Set m_tempDoc = Nothing
'        End If
'    End If
'    Dim activeDoc As Document
'    Set activeDoc = ActiveDocument
'
'' men kun hvis ikke eksisterer allerede
'Dim d As Document
'If m_tempDoc Is Nothing Then
'For Each d In Application.Documents
''    If d.BuiltInDocumentProperties("Title") = "MMtempDoc" Then
'    If d.ActiveWindow.Caption = "WordMatTempDoc" Then
'        Set m_tempDoc = d
'        Exit For
'    End If
'Next
'End If
'
'If m_tempDoc Is Nothing Then
'    Set m_tempDoc = Documents.Add(, , , False)
''    m_tempDoc.ActiveWindow.Left = farRight
''    m_tempDoc.BuiltInDocumentProperties("Title") = "MMtempDoc" ' på mac gav denne problemer. Der blev skiftet fokus til tempdoc nogle sekunder senere. Måske fordi den er meget langsom
'    'Mac: Visible=False?
'    m_tempDoc.ActiveWindow.Caption = "WordMatTempDoc"
'
'    m_tempDoc.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = Sprog.A(680) '"Do NOT edit this document or close or it. WordMat needs it for calculations. Anything you enter here will be deleted."
'    'Note: Update 14.2.5 for Office 2011 allows document to be placed outside screen
'    'm_tempDoc.ActiveWindow.WindowState = wdWindowStateMinimize
'    m_tempDoc.Saved = True
'End If
'
'' fjernet 26/1-17
''If Not m_tempDoc.BuiltInDocumentProperties("Title") = "MMtempDoc" Then
''    m_tempDoc.Close SaveChanges:=wdDoNotSaveChanges
''    m_tempDoc.ActiveWindow.Left = farRight
''    Set m_tempDoc = Documents.Add(, , , False)
''    m_tempDoc.BuiltInDocumentProperties("Title") = "MMtempDoc"
''
''    'Mac: Visible=False?
''    m_tempDoc.ActiveWindow.Caption = "WordMatTempDoc"
''    'Note: Update 14.2.5 for Office 2011 allows document to be placed outside screen
''    'm_tempDoc.ActiveWindow.WindowState = wdWindowStateMinimize
''    m_tempDoc.Saved = True
''End If
'    'Mac:
'    Set tempDoc = m_tempDoc
'    If Not activeDoc Is Nothing Then activeDoc.Activate
'End Function
'Sub SetTempDocSaved()
'    On Error Resume Next
'    m_tempDoc.Saved = True
'End Sub
'#End If

Sub ChangeAutoHyphen()
    Options.AutoFormatAsYouTypeReplaceFarEastDashes = False
    Options.AutoFormatAsYouTypeReplaceSymbols = False
End Sub

Sub ShowCustomizationContext()
'    MsgBox CustomizationContext & vbCrLf & ActiveDocument.AttachedTemplate
    MsgBox Templates(4)
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
    CustomizationContext = GemT

End Function
Function GetWordMatTemplate(Optional NormalDotmOK As Boolean = False) As Template
    ' Hvis det aktuelle dokument hedder wordmat*.dotm så returneres den som template
    ' Ellers søges alle globale skabeloner igennem efter om der er en der hedder wordmat*.dotm
    If Len(ActiveDocument.AttachedTemplate) > 10 Then
        If LCase(Left(ActiveDocument.AttachedTemplate, 7)) = "wordmat" And LCase(right(ActiveDocument.AttachedTemplate, 5)) = ".dotm" Then
            Set GetWordMatTemplate = ActiveDocument.AttachedTemplate
            Exit Function
        End If
    End If
    If NormalDotmOK Then
        Set GetWordMatTemplate = NormalTemplate
    End If

' Det duer ikke at ændre wordmat.dotm hvis filen ikke er åbnet direkte. Den kan ikke gemmes.
'    For Each WT In Application.Templates
'        If LCase(Left(WT, 7)) = "wordmat" And LCase(right(WT, 5)) = ".dotm" Then
'            Set GetWordMatTemplate = WT
'            Exit Function
'        End If
'    Next
End Function
Public Sub GenerateKeyboardShortcutsNormalDotm()
' gemmer KeyboardShortcuts i WordMat.dotm, hvis det er selve wordMat.dotm filen der er åbnet. Hvis ikke gemmes i normal.dotm.
' Det kan give problemer ved en opdatering at benytte denne metode
    GenerateKeyboardShortcutsPar True
End Sub
Public Sub GenerateKeyboardShortcutsWordMat()
' gemmer KeyboardShortcuts i WordMat.dotm, men kun hvis det er selve wordMat.dotm filen der er åbnet
    GenerateKeyboardShortcutsPar False
End Sub
Public Sub GenerateKeyboardShortcutsPar(Optional NormalDotmOK As Boolean = False)
    Dim Wd As WdKey, WT As Template
    Dim WdMac As WdKey
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
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyG, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="Gange"
        
If Sprog.SprogNr = 1 Then
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyB, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="beregn"
Else
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyC, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="beregn"
End If

#If Mac Then
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyReturn, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="beregn"
#Else
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyReturn, Wd, wdKeyControl), KeyCategory:=wdKeyCategoryCommand, Command:="beregn"
#End If

If Sprog.SprogNr = 1 Then
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyL, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="MaximaSolve"
Else
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyE, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="MaximaSolve"
End If
    
If Sprog.SprogNr = 1 Then
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyS, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="InsertSletDef"
Else
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="InsertSletDef"
End If
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyD, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="InsertDefiner"
    
If Sprog.SprogNr = 1 Then
'#If Mac Then ' alt+i bruges til numerisk tegn på mac, så hellere ikke genvej til indstillinger
'#Else
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyJ, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="MaximaSettings"
'#End If
Else
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyO, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="MaximaSettings"
End If
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyP, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="StandardPlot"
        
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyM, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="NewEquation"
        
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyR, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="ForrigeResultat"
        
If Sprog.SprogNr = 1 Then
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyE, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="ToggleUnits"
Else
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyU, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="ToggleUnits"
End If
    
If Sprog.SprogNr = 1 Then
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyO, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="Omskriv"
Else
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyS, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="Omskriv"
End If
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyN, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="ToggleNum"
        
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyT, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="ToggleLatex"
        
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyQ, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="SaveDocToLatexPdf()"
        
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF, Wd), KeyCategory:=wdKeyCategoryCommand, Command:="WMPShowFormler"
    
slut:
    Set CustomizationContext = GemT

End Sub


Function GetProgramFilesDir() As String
    ' bruges ikke af maxima mere da det er dll-filen der står for det nu.
    ' bruges af de Worddokumenter mm. der skal findes
    'MsgBox GetProgFilesPath
    On Error GoTo fejl
#If Mac Then
    GetProgramFilesDir = "/Applications/"
#Else
    If ProgramFilesDir <> "" Then
        GetProgramFilesDir = ProgramFilesDir
    Else
        GetProgramFilesDir = RegKeyRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\ProgramFilesDir")
        If Dir(GetProgramFilesDir & "\WordMat", vbDirectory) = "" Then
            GetProgramFilesDir = RegKeyRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\ProgramW6432Dir")
        End If
        If Dir(GetProgramFilesDir & "\WordMat", vbDirectory) = "" Then
            GetProgramFilesDir = Environ("ProgramFiles")
        End If
        If Dir(GetProgramFilesDir & "\WordMat", vbDirectory) = "" Then
            GetProgramFilesDir = RegKeyRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\ProgramFilesDir (x86)")
        End If
        ProgramFilesDir = GetProgramFilesDir
    End If
#End If

    GoTo slut
fejl:
    MsgBox Sprog.A(110), vbOKOnly, Sprog.Error
slut:
    'MsgBox GetProgramFilesDir
End Function
Function GetDocumentsDir() As String
On Error GoTo fejl
  If DocumentsDir <> "" Then
    GetDocumentsDir = DocumentsDir
  Else
#If Mac Then
    GetDocumentsDir = MacScript("return POSIX path of (path to documents folder) as string")
#Else
 GetDocumentsDir = RegKeyRead("HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal")
 If Dir(GetDocumentsDir, vbDirectory) = "" Then
     GetDocumentsDir = "c:\"
 End If
#End If
 DocumentsDir = GetDocumentsDir
 End If
 
GoTo slut
fejl:
    MsgBox Sprog.A(110), vbOKOnly, Sprog.Error
slut:
'MsgBox GetProgramFilesDir
End Function

Function GetDownloadsFolder() As String
#If Mac Then
    GetDownloadsFolder = RunScript("GetDownloadsFolder", vbNullString)
#Else
    GetDownloadsFolder = RegKeyRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\{374DE290-123F-4565-9164-39C4925E467B}")
    GetDownloadsFolder = Replace(GetDownloadsFolder, "%USERPROFILE%", Environ$("USERPROFILE"))
#End If
End Function
Function RegKeyRead(i_RegKey As String) As String
'eks syntaks
'"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\ProgramFilesDir"
#If Mac Then
    RegKeyRead = GetSetting("com.wordmat", "defaults", i_RegKey)
#Else
Dim myWS As Object

  On Error Resume Next
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'read key from registry
  RegKeyRead = myWS.RegRead(i_RegKey)
#End If
End Function

Function RegKeyExists(i_RegKey As String) As Boolean
#If Mac Then
    RegKeyExists = True
    If GetSetting("com.wordmat", "defaults", i_RegKey) = "" Then RegKeyExists = False
#Else
Dim myWS As Object

  On Error GoTo ErrorHandler
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'try to read the registry key
  myWS.RegRead i_RegKey
  'key was found
  RegKeyExists = True
  Exit Function

ErrorHandler:
  'key was not found
  RegKeyExists = False
#End If
End Function

'sets the registry key i_RegKey to the
'value i_Value with type i_Type
'if i_Type is omitted, the value will be saved as string
'if i_RegKey wasn't found, a new registry key will be created
Sub RegKeySave(ByVal i_RegKey As String, _
               ByVal i_Value As String, _
      Optional ByVal i_Type As String = "REG_SZ")
#If Mac Then
    SaveSetting "com.wordmat", "defaults", i_RegKey, i_Value
#Else
Dim myWS As Object
    On Error Resume Next
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'write registry key
  myWS.RegWrite i_RegKey, i_Value, i_Type
#End If
End Sub

'deletes i_RegKey from the registry
'returns True if the deletion was successful,
'and False if not (the key couldn't be found)
Function RegKeyDelete(i_RegKey As String) As Boolean
#If Mac Then
    DeleteSetting "com.wordmat", "defaults", i_RegKey
#Else
Dim myWS As Object

  On Error GoTo ErrorHandler
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'delete registry key
    On Error Resume Next
  myWS.RegDelete i_RegKey
  'deletion was successful
  RegKeyDelete = True
  Exit Function

ErrorHandler:
  'deletion wasn't successful
  RegKeyDelete = False
#End If
End Function
Sub TestLink()
    OpenLink "https://www.eduap.com"
End Sub
Sub TestLink2()
' virker ikke, men det burde være sådan
   ' ActiveDocument.FollowHyperlink Address:="file://C:\Program Files (x86)/WordMat/geogebra-math-apps/GeoGebra/HTML5/5.0/GeoGebra.html", Method:=msoMethodGet, NewWindow:=True, ExtraInfo:="command=f(x)=x"
   'virker heller ikke5
'    CreateObject("Shell.Application").Open "C:\Program Files (x86)\WordMat\geogebra-math-apps\GeoGebra\HTML5\5.0\GeoGebra.html"
'    RunDefaultProgram "C:\Program Files (x86)\WordMat\geogebra-math-apps\GeoGebra\HTML5\5.0\GeoGebra.html?id=3"
'    shell """" & GetProgramFilesDir & "\Microsoft\Edge\Application\msedge.exe"" ""file://C:\Program Files (x86)/WordMat/geogebra-math-apps/GeoGebra/HTML5/5.0/GeoGebra.html?command=f(x)=x""", vbNormalFocus
'    CreateObject("Shell.Application").Open CVar(GetProgramFilesDir & "\Microsoft\Edge\Application\msedge.exe ""file://C:\Program Files (x86)/WordMat/geogebra-math-apps/GeoGebra/HTML5/5.0/GeoGebra.html?command=f(x)=x""")
'    CreateObject("Shell.Application").Open GetProgramFilesDir & "\Microsoft\Edge\Application\msedge.exe 'file://C:\Program Files (x86)/WordMat/geogebra-math-apps/GeoGebra/HTML5/5.0/GeoGebra.html?command=f(x)=x'"
   Dim shellcmd As String
'   shellcmd = "cmd /K ""C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"" ""file://C:\Program Files (x86)/WordMat/geogebra-math-apps/GeoGebra/HTML5/5.0/GeoGebra.html?command=f(x)=x""" '/K holder cmd åben /C lukker
   shellcmd = "cmd /S /K """"C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"" ""file://C:\\Program Files (x86)\\WordMat\\geogebra-math-apps\\GeoGebra\\HTML5\\5.0\\GeoGebra.html?command=f(x)=x"""""
'   shellcmd = "cmd /K C:\Program\\ Files\\ (x86)\Microsoft\Edge\Application\msedge.exe ""file://C:\Program Files (x86)/WordMat/geogebra-math-apps/GeoGebra/HTML5/5.0/GeoGebra.html?command=f(x)=x""" '/K holder cmd åben /C lukker
'    Debug.Print shellcmd
'   shell shellcmd, vbNormalFocus
End Sub
Sub OpenLink(Link As String, Optional Script As Boolean = False)
' obs: Script er altid true på mac for at forhindre advarsel
On Error Resume Next

#If Mac Then
    Script = True
    If Script Then
        RunScript "OpenLink", Link
    Else
        ActiveDocument.FollowHyperlink Address:=Link, NewWindow:=True
    End If
#Else
' ActiveDocument.FollowHyperlink fjerner parametre som fx. ?command=...   Derfor kan det være nødvendigt at bruge script
    If Script Then
        Shell """" & GetProgramFilesDir & "\Microsoft\Edge\Application\msedge.exe"" """ & Link & """", vbNormalFocus ' giver problemer med bitdefender
'        shell "cmd /S /C """"" & GetProgramFilesDir & "\Microsoft\Edge\Application\msedge.exe"" """ & Link & """""", vbNormalFocus ' Denne bliver ikke fanget ved install, men bitdefender blokerer den ved kørsel
    Else
        ActiveDocument.FollowHyperlink Address:=Link, NewWindow:=True ' hvis linket ikke virker så sker der bare ingen ting
    End If
#End If
fejl:
End Sub

Sub InsertSletDef()
    Dim gemfontsize As Integer
    Dim gemitalic As Boolean
    Dim gemfontcolor As Integer
    Dim gemsb As Integer
    Dim gemsa As Integer
    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord
            
    gemfontsize = Selection.Font.Size
    gemitalic = Selection.Font.Italic
    gemfontcolor = Selection.Font.ColorIndex
    gemsb = Selection.ParagraphFormat.SpaceBefore
    gemsa = Selection.ParagraphFormat.SpaceAfter
            
            
            Selection.Font.Size = 8
            Selection.Font.ColorIndex = wdGray50
            
    insertribformel "", Sprog.A(69) & ":"
            
'            Selection.TypeParagraph
            Selection.Font.Size = gemfontsize
            Selection.Font.Italic = gemitalic
            Selection.Font.ColorIndex = gemfontcolor
    With Selection.ParagraphFormat
        .SpaceBefore = gemsb
'        .SpaceBeforeAuto = False
        .SpaceAfter = gemsa
'        .SpaceAfterAuto = False
    End With
    Oundo.EndCustomRecord
End Sub

Sub InsertDefiner()
    On Error GoTo fejl

    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord

    Application.ScreenUpdating = False
    If Selection.OMaths.Count > 0 Then
        Selection.OMaths(1).Range.Select
        Selection.Collapse wdCollapseStart
        Selection.InsertAfter (Sprog.A(62) & ": ")
    Else
        Selection.InsertAfter (Sprog.A(62) & ": ")
        Selection.OMaths.Add Range:=Selection.Range
        Selection.OMaths.BuildUp
    '    Selection.OMaths(1).BuildUp
    End If
    Selection.Collapse wdCollapseEnd
        
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
    Oundo.EndCustomRecord
End Sub

Sub ForrigeResultat()

    Dim ra As Range
    Dim sr As Range
    Dim r As Range
    Dim s As String
    Dim start As Integer
    Dim sslut As Integer
    Dim matfeltno As Integer
    Dim hopover As Boolean
    Application.ScreenUpdating = False
    
    On Error Resume Next
    If Selection.OMaths.Count = 0 Then GoTo slut
    
    Dim scrollpos As Double
    scrollpos = ActiveWindow.VerticalPercentScrolled

    If omax Is Nothing Then
        DoEvents
        Set omax = New CMaxima
    End If
    
    Set sr = Selection.Range
    If ResPos1 = Selection.Range.start Then ' hvis gentaget tryk
        If ResIndex < 0 Then
            ResFeltIndex = ResFeltIndex + 1
            ResIndex = 0
        Else
            ResIndex = ResIndex + 1
        End If
'        ActiveDocument.Range(ResPos1, ResPos2).text = ""
    Else
        ResFeltIndex = 0
        ResIndex = 0
    End If
    
    If ResIndex < 0 Then ResIndex = 0
    On Error GoTo fejl
    start = Selection.Range.start
    sslut = Selection.Range.End
    Set ra = ActiveDocument.Range
    ra.End = sslut + 1
    matfeltno = ra.OMaths.Count
    Do
        If ResFeltIndex >= matfeltno - 1 Then
            If ActiveDocument.Range.OMaths(matfeltno).Range.text = Selection.Range.text Then
                Selection.text = ""
                Selection.OMaths.Add Range:=Selection.Range
            Else
                Selection.text = ""
            End If
            GoTo fejl
        End If
'        ActiveDocument.Range.OMaths(matfeltno - 1 - ResFeltIndex).Range.Select
        Set r = ActiveDocument.Range.OMaths(matfeltno - 1 - ResFeltIndex).Range
        If Len(r.text) = 0 Then
            ResFeltIndex = ResFeltIndex + 1
            ResIndex = 0
            GoTo slut
        End If
'        s = omax.ReadEquation2(r)
        s = omax.ReadEquation(r)
        hopover = False
        If InStr(VBA.LCase(s), "defin") > 0 Then
            ResFeltIndex = ResFeltIndex + 1
            ResIndex = 0
            hopover = True
        Else
            s = KlipTilLigmed(s, ResIndex)
            If Len(s) = 1 Or s = "f(x)" Then
                If ResIndex < 0 Then
                    ResFeltIndex = ResFeltIndex + 1
                    ResIndex = 0
                Else
                    ResIndex = ResIndex + 1
                End If
                hopover = True
            ElseIf s = VBA.ChrW(8661) Then
                ResFeltIndex = ResFeltIndex + 1
                ResIndex = 0
                hopover = True
            End If
        End If
Loop While hopover
    
    sr.Select
    ResPos1 = Selection.Range.start
    If Selection.Range.text = "Skriv ligningen her." Then
        ResPos1 = ResPos1 - 1 ' hvis tom i forvejen er selection af eller anden grund 1 tegn for meget
    End If
    s = Replace(s, VBA.ChrW(8289), "") ' funktionstegn  sin(x) bliver ellers til si*n(x). også problem med andre funktioner
    Selection.text = s
    
'    Dim ml As Integer
'    ml = Len(ActiveDocument.Range.OMaths(matfeltno).Range.text)
'    Selection.TypeText s
'    ResPos2 = Selection.start
'    ActiveDocument.Range.OMaths(ra.OMaths.Count).BuildUp
'    ResPos2 = ResPos1 + Len(ActiveDocument.Range.OMaths(matfeltno).Range.text) - ml
GoTo slut
fejl:
    ResIndex = 0
    ResFeltIndex = 0
    ResPos2 = 0
    ResPos1 = 0
slut:
'    Selection.End = sslut ' slut skal være først eller går det galt
'    Selection.start = start
'    Call sr.Move(wdCharacter, Len(s))
'    sr.Select
'    ActiveDocument.Range(ResPos1, ResPos1).Select
'    sr.Select
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub

Function KlipTilLigmed(text As String, ByVal indeks As Integer) As String
' returnerer sidste del af texten til første position talt fra enden for = eller ca. ligmed
' = i sumtegn ignoreres
    
    Dim posligmed As Integer
    Dim possumtegn As Integer
    Dim posca As Integer
    Dim poseller As Integer
    Dim Pos As Integer
    Dim Arr(20) As String
    Dim i As Integer
    
    Do ' gå tilbage til nærmeste ligmed
    posligmed = InStr(text, "=")
    possumtegn = InStr(text, VBA.ChrW(8721))
    posca = InStr(text, VBA.ChrW(8776))
    poseller = InStr(text, VBA.ChrW(8744))
    
    Pos = Len(text)
'    pos = posligmed
    If posligmed > 0 And posligmed < Pos Then Pos = posligmed
    If posca > 0 And posca < Pos Then Pos = posca
    If poseller > 0 And poseller < Pos Then Pos = poseller
    
    If possumtegn > 0 And possumtegn < Pos Then ' hvis sumtegn er der =tegn som del deraf
        Pos = 0
    End If
    If Pos = Len(text) Then Pos = 0
    If Pos > 0 Then
        Arr(i) = Left(text, Pos - 1)
        text = right(text, Len(text) - Pos)
        i = i + 1
    Else
        Arr(i) = text
    End If
    Loop While Pos > 0
    
    If indeks = i Then ResIndex = -1  ' global variabel markerer at der ikke er flere til venstre
    If i = 0 Then
        KlipTilLigmed = text
        ResIndex = -1
    Else
        KlipTilLigmed = Arr(i - indeks)
    End If
    
    ' fjern retur og mellemrum mm.
'    s = Replace(s, vbCrLf, "")
    KlipTilLigmed = Replace(KlipTilLigmed, vbCr, "")
    KlipTilLigmed = Replace(KlipTilLigmed, VBA.ChrW(11), "")
'    s = Replace(s, vbLf, "")
    KlipTilLigmed = Replace(KlipTilLigmed, VBA.ChrW(8744), "") 'eller tegn
'    KlipTilLigmed = Replace(KlipTilLigmed, " ", "")
    KlipTilLigmed = Trim(KlipTilLigmed)
    
    If InStr(KlipTilLigmed, "/") > 0 Then KlipTilLigmed = "  " & KlipTilLigmed ' hvorfor?
    
End Function

Function ReadEquationFast(Optional ir As Range) As String
' Oversætter selection der er omath til streng
    Dim sr As Range

    Selection.OMaths(1).Range.Select
    Set sr = Selection.Range
    sr.OMaths.BuildUp
    sr.OMaths.Linearize
    sr.OMaths(1).ConvertToNormalText
    
    ReadEquationFast = sr.OMaths(1).Range.text
    
    Selection.OMaths(1).ConvertToMathText
    Selection.OMaths(1).Range.Select
    Selection.OMaths.BuildUp
    Selection.Collapse (wdCollapseEnd)

    sr.Select
    
End Function

Sub OpenFormulae(FilNavn As String)
On Error GoTo fejl
#If Mac Then
    Documents.Open "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/WordDocs/" & FilNavn
#Else
    OpenWordFile "" & FilNavn
#End If
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Sub OpenWordFile(FilNavn As String)
    ' OpenWordFile ("Figurer.docx")

    Dim filnavn1 As String
#If Mac Then
    FilNavn = Replace(FilNavn, "\", "/")
    filnavn1 = GetWordMatDir() & "WordDocs/" & FilNavn
    Documents.Open filnavn1
#Else
    Dim filnavn2 As String
    Dim appdir As String
    Dim fs
    On Error GoTo fejl
    Set fs = CreateObject("Scripting.FileSystemObject")
    appdir = Environ("AppData")
    filnavn1 = appdir & "\WordMat\WordDocs\" & FilNavn

    If Dir(filnavn1) = vbNullString Then
        filnavn2 = GetProgramFilesDir & "\WordMat\WordDocs\" & FilNavn

        If Dir(filnavn2) <> vbNullString Then
            If Dir(appdir & "\WordMat\", vbDirectory) = vbNullString Then MkDir appdir & "\WordMat\"
            If Dir(appdir & "\WordMat\WordDocs\", vbDirectory) = vbNullString Then MkDir appdir & "\WordMat\WordDocs\"
            fs.CopyFile filnavn2, appdir & "\WordMat\WordDocs\"
        End If
    End If
    
    If Dir(filnavn1) <> "" Then
        Documents.Open fileName:=filnavn1
    ElseIf Dir(filnavn2) <> "" Then
        Documents.Open fileName:=filnavn2, ReadOnly:=True
    Else
        MsgBox Sprog.A(111) & FilNavn, vbOKOnly, Sprog.Error
    End If
#End If

    GoTo slut
fejl:
    MsgBox Sprog.A(111) & FilNavn, vbOKOnly, Sprog.Error
slut:

End Sub

Function GetRandomTip()
    Dim i As Integer
    Dim n As Integer
    Dim mindste As Integer
    Dim tip As String
    n = 29 ' antal tip
    mindste = 0
    
    If AntalB < 10 Then
        mindste = 0
        n = 5
    ElseIf AntalB < 20 Then
        mindste = 0
        n = 12
    ElseIf AntalB < 50 Then
        mindste = 0
        n = 15
    ElseIf AntalB < 100 Then
        mindste = 3
        n = 20
    ElseIf AntalB < 130 Then
        mindste = 0
        n = 29
    Else
        mindste = 3
        n = 29
    End If
    
    Randomize
    i = Int(Rnd(1) * (n - mindste) + mindste) ' tilfældigt tal 0-(n-1)
'hævet a " & VBA.ChrW(7491) & " hævet b " & VBA.ChrW(7495) & " hævet p  " & VBA.ChrW(7510) & "  hævet q " & VBA.ChrW(8319) & "
' sænket 0 " & VBA.ChrW(8320) & " sænket 1 " & VBA.ChrW(8321) & " hævet 2 " & VBA.ChrW(8322) & "_

    Select Case i
    Case 0
        tip = Sprog.A(325) '"Genvejen   Alt + M   indsætter nyt matematikfelt"
    Case 1
        tip = Sprog.A(326) '"Genvejen   Altgr + enter   beregner"
    Case 2
        tip = Sprog.A(327) '"Genvejen   Alt + L    løser ligning"
    Case 3
        tip = Sprog.A(328) '"Genvejen   Alt + i   åbner indstillinger"
    Case 4
        tip = Sprog.A(329) '"Genvejen   Alt + r   henter forrige resultater"
    Case 5
        tip = Sprog.A(330) '"Genvejen   Alt + E   slår enheder til/fra"
    Case 6
        tip = Sprog.A(331) & "   x_1   ->   x" & VBA.ChrW(8321)
    Case 7
        tip = Sprog.A(332) & VBA.ChrW(955)
    Case 8
        tip = Sprog.A(333) '"Genvejen  Alt + P  viser grafen for en forskrift"
    Case 9
        tip = Sprog.A(334) '"Du kan selv gemme ændringer i formelsamlingerne og figurdokumentet."
    Case 10
        tip = Sprog.A(335) '"Excelark kan både indsættes indlejret i dokumentet eller åbnes i Excel"
    Case 11
        tip = Sprog.A(336) '"Hurtig og bedre retning af opgaver? Prøv WordMark - www.eduap.com"
    Case 12
        tip = "(a+b)" & VBA.ChrW(178) & " = a" & VBA.ChrW(178) & " + b" & VBA.ChrW(178) & " + 2ab"
    Case 13
        tip = "(a+b)(a-b) = a" & VBA.ChrW(178) & " - b" & VBA.ChrW(178)
    Case 14
        tip = "(a-b)" & VBA.ChrW(178) & " = a" & VBA.ChrW(178) & " + b" & VBA.ChrW(178) & " - 2ab"
    Case 15
        tip = "(a" & VBA.ChrW(183) & "b)" & VBA.ChrW(7510) & " = a" & VBA.ChrW(7510) & VBA.ChrW(183) & "b" & VBA.ChrW(7510)
    Case 16
        tip = Sprog.A(337) '"Du har ialt foretaget " & AntalB & " beregninger med WordMat"
    Case 17
        tip = "log(a" & VBA.ChrW(7495) & ") = b" & VBA.ChrW(183) & "log(a)"
    Case 18
        tip = "log(a/b) = log(a) - log(b)"
    Case 19
        tip = "log(a" & VBA.ChrW(183) & "b) = log(a) + log(b)"
    Case 20
        tip = "\int      ->      " & VBA.ChrW(8747)
    Case 21
        tip = Sprog.A(338) '"Under definitioner kan du definere fysiske konstanter og tabelværdier"
    Case 22
        tip = Sprog.A(339) '"Word kører hurtigere i kladdevisning, specielt for ligninger."
    Case 23
        tip = Sprog.A(340) '"Indsæt dine egne beskrivelser i Exceldiagrammer vha. Indsæt/tekstboks"
    Case 24
        tip = "(a/b)" & VBA.ChrW(7510) & " = a" & VBA.ChrW(7510) & "/b" & VBA.ChrW(7510)
    Case 25
        tip = "a/b + c/d = (ad+bc)/bd"
    Case 26
        tip = Sprog.A(341) '"Genvejen   Alt + N   skifter mellem auto,eksakt,num"
    Case 27
        tip = Sprog.A(342) '"En funktion med bogstav-konstanter får automatisk skydere i GeoGebra"
    Case 28
        tip = "\inc  ->  delta"
    Case 29
        tip = ""
    Case 30
        tip = ""
    Case Else
        tip = ""
    End Select
    
    
'    GetRandomTip = tip(i)
    GetRandomTip = tip
    
End Function
Sub ShowTip()
    MsgBox GetRandomTip
End Sub
Sub ToggleUnits()
    Dim ufq As UserFormQuick
    
    If MaximaUnits Then
        Set ufq = New UserFormQuick
        ufq.Label_text.Caption = Sprog.A(166) 'unit off
        DoEvents
        ufq.Show vbModeless
        TurnUnitsOff
    Else
        MaximaUnits = True
        DoEvents
        PrepareMaximaNoSplash
#If Mac Then
#Else
        If MaxProc Is Nothing Then Exit Sub
#End If
chosunit:
            OutUnits = InputBox(Sprog.A(167), Sprog.A(168), OutUnits)
            If InStr(OutUnits, "/") > 0 Or InStr(OutUnits, "*") > 0 Or InStr(OutUnits, "^") > 0 Then
                MsgBox2 Sprog.A(343), vbOKOnly, Sprog.Error
                GoTo chosunit
            End If
            On Error Resume Next
            TurnUnitsOn
'            TurnUnitsOn
'            MaxProc.OutUnits = omax.ConvertUnits(OutUnits)
'            MaxProc.Units = 1
'            MaxProc.CloseProcess
'            MaxProc.StartMaximaProcess
        End If
    
'    UserFormQuick.Hide
'    Unload ufq

End Sub

Sub TurnUnitsOn()
' det er nødvendigt at slette definitioner først da de ellers nemt får load(unit) til at fejle
    
On Error Resume Next
    MaximaUnits = True
    Application.OMathAutoCorrect.Functions("min").Delete  ' ellers kan min ikke bruges som enhed
    
    Exit Sub ' resten er ikke nødv v. 1.23
    '******************************
#If Mac Then
#Else
    Exit Sub ' overtaget af maxprocunit
#End If

Dim text As String
    
    MaxProc.Units = 1
    text = omax.KillDef
    If Len(text) > 0 Then
         text = Left(text, Len(text) - 1) 'fjern sidste komma
         text = "errcatch(kill(" & text & "))"
         omax.KillDef = ""
    Else
        text = "" ' mærkeligt men len(text)=0 er ikke nødv ""
    End If
    
'    MaxProc.ExecuteMaximaCommand text, 0
'    MaxProc.OutUnits = omax.ConvertUnits(OutUnits)
'     MaxProc.TurnUnitsOn text, ""
'     MaxProc.TurnUnitsOn "", ""

'Dim text As String
   text = "[" & text & "load(WordMatUnitAddon)"
'    text = "[" & text & "keepfloat:false,usersetunits:[N,J,W,Pa,C,V,F,Ohm,T,H,K],load(unit)"
    If OutUnits <> "" Then
        text = text & ",setunits(" & omax.ConvertUnits(OutUnits) & ")"
    End If
    text = text & "]$"
    
    MaxProc.ExecuteMaximaCommand text, 0

'            MaxProc.TurnUnitsOn
End Sub
Sub TurnUnitsOff()
        MaximaUnits = False
#If Mac Then
        If Not MaxProc Is Nothing Then
            MaxProc.Units = 0
'            MaxProc.CloseProcess ' skal ikke køres efter v. 1.23
'            MaxProc.StartMaximaProcess
        End If
#End If
        On Error Resume Next
        Application.OMathAutoCorrect.Functions.Add "min"

End Sub
Sub UpdateUnits()
    Dim text As String
    text = "setunits(" & omax.ConvertUnits(OutUnits) & ")$"
    
    MaxProc.ExecuteMaximaCommand text, 0

End Sub
Sub ToggleNum()
    Dim ufq As New UserFormExactNum
    If MaximaExact = 0 Then
        ufq.SetExact
        DoEvents
        MaximaExact = 1
        ufq.Show vbModeless
    ElseIf MaximaExact = 1 Then
        ufq.SetNum
        DoEvents
        MaximaExact = 2
        ufq.Show vbModeless
    Else
'        ufq.Label_text.Caption = "Auto"
        ufq.SetAuto
        DoEvents
        MaximaExact = 0
        ufq.Show vbModeless
    End If
    
    On Error Resume Next
    WoMatRibbon.Invalidate
    
End Sub
Sub CheckForUpdateOld()
    Dim result As VbMsgBoxResult
'    On Error GoTo Fejl
#If Mac Then
    MsgBox "Automatic update is not (yet) available on Mac" & vbCrLf & "Current version is: " & AppVersion & vbCrLf & vbCrLf & "Remember the version no. above. You will now be send to the download page where you can check for a newer version -  eduap.com"
    OpenLink "https://www.eduap.com/da/download-wordmat/"
#Else
    Dim nyversion As String, News As String
    PrepareMaxima
    nyversion = MaxProc.CheckForUpdate()
    If nyversion = "" Then
        MsgBox Sprog.A(112), vbOKOnly, Sprog.Error
        Exit Sub
    End If

    If nyversion = AppVersion Then
        MsgBox Sprog.A(344) & " " & AppNavn, vbOKOnly, Sprog.OK
    Else
        News = MaxProc.GetVersionNews()
        result = MsgBox(Sprog.A(21) & News & vbCrLf & vbCrLf & Sprog.A(22), vbYesNo, Sprog.A(23))
        If result = vbYes Then
            OpenLink "https://www.eduap.com/da/download-wordmat/"
        End If
    End If

#End If
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Sub CheckForUpdate()
'#If Mac Then
'    CheckForUpdateF False
'#Else
    CheckForUpdateWindows False
'#End If
End Sub
'Sub CheckForUpdateF(Optional Silent As Boolean = False)
'' Denne skulle virke på mac og windows via speciel webrequest klasse, men det giver fejl i dictionary klassen på mac.
'    ' Create a WebClient for executing requests
'    ' and set a base url that all requests will be appended to
'    Dim p As Long, p2 As Long, p3 As Long, s As String, v As String, News As String
'    Dim MapsClient As New WebClient
''   On Error GoTo fejl
''#If Mac Then
''    MsgBox "Automatic update is not (yet) available on Mac" & vbCrLf & "Current version is: " & AppVersion & vbCrLf & vbCrLf & "Remember the version no. above. You will now be send to the download page where you can check for a newer version -  eduap.com"
''    OpenLink "http://eduap.com/wordmat/"
''#Else
'    Dim result As VbMsgBoxResult
'    MapsClient.BaseUrl = "https://www.eduap.com/wordmat-version-history/"
'
'    ' Use GetJSON helper to execute simple request and work with response
'    Dim Resource As String
'    Dim Response As WebResponse
'    Dim Request As New WebRequest
'    '    Request.Resource = "index.html"
'
'    Request.Format = WebFormat.PlainText
'
'    Request.Method = WebMethod.HttpGet
'    '    Request.ResponseFormat = PlainText
'    '    Request.Method = WebMethod.HttpGet
'    '    Request.Method = WebMethod.HttpPost
''    Resource = "" '
'    '    "directions/json?" & _
'    '        "origin=" & Origin & _
'    '        "&destination=" & Destination & _
'    '        "&sensor=false"
'
'    '    Set Response = MapsClient.GetJson(Resource)
'    Set Response = MapsClient.Execute(Request)
'    ' => GET https://maps.../api/directions/json?origin=...&destination=...&sensor=false
'    '    MsgBox Response.StatusCode & " - " & Response.StatusDescription
'
'    If Response.StatusCode = WebStatusCode.OK Or Response.StatusCode = 301 Then ' af ukendte årsager kommer der 301 fejl på mac, men det virker
'        '        MsgBox Response.Content
''        p = InStr(s, "Version history")
'
''        v = Trim(Mid(s, p + 16, 4))
'        '        MsgBox Response.Content
'        s = Response.Content
'        p = InStr(s, "<body")
'        p = InStr(p, s, "Version ")
'        If p <= 0 Then GoTo fejl
'        v = Trim(Mid(s, p + 8, 4))
'        p2 = InStr(p + 10, s, "Version " & AppVersion)
'        If p2 <= 0 Then p2 = InStr(p + 10, s, "Version")
'        If p2 <= 0 Then p2 = p + 50
''        p3 = InStr(p, s, "<p>")
'        News = Mid(s, p, p2 - p)
''        News = Replace(News, "- ", vbCrLf & "- ")
'        News = Replace(News, "&#8211;", vbCr & " -") ' bindestreg
'        News = Replace(News, "Version ", vbCrLf & "Version ") ' bindestreg
'        News = Replace(News, "<br />", "")
'        News = Replace(News, "<strong>", "")
'        News = Replace(News, "</strong>", "")
'        News = Replace(News, "<p>", "")
'        News = Replace(News, "</p>", "")
'        If v = AppVersion Then
'            If Not Silent Then
'                MsgBox Sprog.A(344) & " " & AppNavn, vbOKOnly, Sprog.OK
'            End If
'        Else
'            result = MsgBox(Sprog.A(21) & News & vbCrLf & Sprog.A(22), vbYesNo, Sprog.A(23))
'            If result = vbYes Then
'                OpenLink "https://www.eduap.com/da/download-wordmat/"
'            End If
'        End If
'    Else
'        GoTo slut
'    End If
'
'    If Response.StatusCode = WebStatusCode.OK Or Response.StatusCode = 301 Then
'    End If
'
''#End If
'    GoTo slut
'fejl:
'    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
'slut:
'End Sub
Sub CheckForUpdateWindows(Optional RunSilent As Boolean = False)
    ' selvom den hedder windows er det nu også mac
    On Error GoTo fejl
    Dim NewVersion As String, p As Integer, News As String, s As String
    Dim UpdateNow As Boolean, PartnerShip As Boolean
'    Dim UFvent As UserFormWaitForMaxima
    
   
    If GetInternetConnectedState = False Then
        If Not RunSilent Then MsgBox "Ingen internetforbindelse", vbOKOnly, "Fejl"
        Exit Sub
    End If
    
    If RunSilent Then
        If (Month(Date) = 5 Or Month(Date) = 6) Then GoTo slut ' ikke automatisk opdatere i maj og juni
        If IsDate(LastUpdateCheck) Then
            If DateDiff("d", LastUpdateCheck, Date) < 7 Then GoTo slut ' hvis der er checket indenfor de sidste 7 dage så afslut
        End If
    End If
    LastUpdateCheck = Date ' denne skal være her, og ikke i slutningen, for hvis der sker en fejl i opdateringen, skal den kun komme én gang
    
    '    s = GetHTML("https://www.eduap.com/wordmat-version-history/")
#If Mac Then
    s = RunScript("CheckUpdate", vbNullString)
    If InStr(s, "404 Not Found") > 0 Then s = vbNullString
#Else
    s = GetHTML("http://screinfo.eduap.com/wordmatversion.txt")
#End If
    If Len(s) = 0 Then
        If Not RunSilent Then
            MsgBox2 "Serveren kan ikke kontaktes", vbOKOnly, "Fejl"
        End If
        GoTo slut
    End If
    NewVersion = s
    p = InStr(NewVersion, vbLf)
    If p > 0 Then
        News = right(NewVersion, Len(NewVersion) - p)
        NewVersion = Trim(Left(NewVersion, p - 1))
    Else ' mac
        p = InStr(NewVersion, vbCr)
        If p > 0 Then
            News = right(NewVersion, Len(NewVersion) - p)
            NewVersion = Trim(Left(NewVersion, p - 1))
        End If
    End If
   
    If Len(NewVersion) = 0 Or Len(NewVersion) > 15 Then
        If Not RunSilent Then
            MsgBox2 "Serveren kan ikke kontaktes", vbOKOnly, "Fejl"
        End If
        GoTo slut
    End If
    
    '        p = InStr(s, "<body")
    '        p = InStr(p, s, "Version ")
    '        If p <= 0 Then GoTo fejl
    '        v = Trim(Mid(s, p + 8, 4))
    '        p2 = InStr(p + 10, s, "Version " & AppVersion)
    '        If p2 <= 0 Then p2 = InStr(p + 10, s, "Version")
    '        If p2 <= 0 Then p2 = p + 50
    '        News = Mid(s, p, p2 - p)
    '        News = Replace(News, "&#8211;", vbCr & " -") ' bindestreg
    '        News = Replace(News, "Version ", vbCrLf & "Version ") ' bindestreg
    '        News = Replace(News, "<br />", "")
    '        News = Replace(News, "<strong>", "")
    '        News = Replace(News, "</strong>", "")
    '        News = Replace(News, "<p>", "")
    '        News = Replace(News, "</p>", "")
    '    If Len(v) = 0 Then
    '        If Not RunSilent Then
    '            MsgBox "Serveren kan ikke kontaktes", vbOKOnly, "Fejl"
    '            GoTo slut
    '        End If
    '    End If
    '    If AppVersion <> v Then
    '        '      If UFreminder.Visible = True Then UFreminder.Top = 100
    '        result = MsgBox(Sprog.A(21) & News & vbCrLf & Sprog.A(22), vbYesNo, Sprog.A(23))
    '        If result = vbYes Then
    '            OpenLink "https://eduap.com/da/download-wordmat/"
    '        End If
    '    Else
    '        If Not RunSilent Then
    '            MsgBox "Du har allerede den nyeste version installeret", vbOKOnly, "Ingen opdatering"
    '        End If
    '    End If
   
    If IsNumeric(AppVersion) And IsNumeric(NewVersion) Then
        If val(AppVersion) < val(NewVersion) Then UpdateNow = True
    Else
        If AppVersion <> NewVersion Then UpdateNow = True
    End If
    On Error Resume Next
    PartnerShip = QActivePartnership()
    On Error GoTo fejl
    
    If UpdateNow Then
        If PartnerShip Then
'            Set UFvent = New UserFormWaitForMaxima
            If MsgBox2(Sprog.A(21) & News & vbCrLf & vbCrLf & "Klik OK for at starte opdateringen.", vbOKCancel, Sprog.A(23)) = vbOK Then
                On Error Resume Next
                Documents.Save NoPrompt:=True, OriginalFormat:=wdOriginalDocumentFormat
'                UFvent.Label_tip.Caption = "Downloader WordMat " & NewVersion
'                UFvent.Label_progress.Caption = "**"
'                UFvent.Show
                On Error GoTo Install2
                Application.Run macroname:="PUpdateWordMat"
                On Error GoTo fejl
            End If
        Else
Install2:
            On Error GoTo fejl
            MsgBox2 Sprog.A(21) & News & vbCrLf & Sprog.A(22) & vbCrLf & vbCrLf & "", vbOKOnly, Sprog.A(23)
            '        If MsgBox(Sprog.A(21) & News & vbCrLf & Sprog.A(22) & vbCrLf & vbCrLf & "", vbYesNo, Sprog.A(23)) = vbYes Then
            If Sprog.SprogNr = 1 Then
                OpenLink "https://www.eduap.com/da/wordmat/"
            Else
                OpenLink "https://www.eduap.com/wordmat/"
            End If
        End If
    Else
        If Not RunSilent Then
            MsgBox2 "Du har allerede den nyeste version af WordMat installeret: v." & AppVersion, vbOKOnly, "Ingen opdatering"
        End If
    End If
   
   
   
    GoTo slut
fejl:
    '   MsgBox "Fejl " & Err.Number & " (" & Err.Description & ") i procedure CheckForUpdate, linje " & Erl & ".", vbOKOnly Or vbCritical Or vbSystemModal, "Fejl"
    If Not RunSilent Then
        MsgBox "Current version is: " & AppVersion & vbCrLf & vbCrLf & "Remember the version no. above. You will now be send to the download page where you can check for a newer version -  www.eduap.com"
        OpenLink "https://www.eduap.com/da/download-wordmat/"
        '        MsgBox "Der skete en fejl i forbindelse at checke for ny version. Det kan skyldes en fejl med internetforbindelsen eller en fejl med serveren. Prøv igen senere, eller check selv på eduap.com om der er kommet en ny version. Den nuværende version er " & AppVersion, vbOKOnly Or vbCritical Or vbSystemModal, "Fejl"
    End If
slut:
    On Error Resume Next
'    Unload UFvent

End Sub

Sub CheckForUpdateSilent()
' maxproc skal være oprettet
    On Error GoTo fejl

'#If Mac Then
'    CheckForUpdateF True
'#Else
    CheckForUpdateWindows True
'#End If
GoTo slut
fejl:
'    MsgBox "Der kunne ikke oprettes forbindelse til serveren", vbOKOnly, "Fejl"
slut:
End Sub
Function GetHTML(URL As String) As String
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", URL & "?cb=" & Timer() * 100, False  ' timer sikrer at det ikke er cached version
        .Send
        GetHTML = .ResponseText
    End With
End Function

Public Function GetInternetConnectedState() As Boolean
#If Mac Then
   GetInternetConnectedState = True
#Else
    Dim r As Long
    r = InternetGetConnectedState(0&, 0&)
    If r = 0 Then
        GetInternetConnectedState = False
    Else
        If r <= 4 Then
            GetInternetConnectedState = True
        Else
            GetInternetConnectedState = False
        End If
    End If
#End If
End Function

Function ConvertNumberToString(ByVal n As Double) As String
    Dim ns As String
    Dim i As Integer
    For i = 1 To MaximaCifre
        ns = ns & "#"
    Next
    If n = 0 Then
        ConvertNumberToString = "0"
        Exit Function
    End If
    If MaximaVidNotation Or Abs(n) > 10 ^ 6 Or Abs(n) < 10 ^ -6 Then
'#If Mac Then
'    ConvertNumberToString = n
'    ConvertNumberToString = Replace(ConvertNumberToString, "e", "E")
'    ConvertNumberToString = Replace(ConvertNumberToString, "+0", "+")
'    ConvertNumberToString = Replace(ConvertNumberToString, "-0", "-")
'#Else
            ConvertNumberToString = VBA.Format(n, "0.0" & ns & "E-0")
'#End If
    Else
'        ConvertNumberToString = Format(n, "General Number")
#If Mac Then
        ConvertNumberToString = VBA.Format(n, "#################0.0####################;-#################0.0####################")
#Else
        ConvertNumberToString = VBA.Format(n, "#################0.0####################")
#End If
        If Len(ConvertNumberToString) > 1 Then
            If right(ConvertNumberToString, 2) = ".0" Then
                ConvertNumberToString = Left(ConvertNumberToString, Len(ConvertNumberToString) - 2)
            ElseIf right(ConvertNumberToString, 2) = ",0" Then
                ConvertNumberToString = Left(ConvertNumberToString, Len(ConvertNumberToString) - 2)
            End If
        End If
    End If
    If DecSeparator = "," Then
        ConvertNumberToString = Replace(ConvertNumberToString, ".", ",")
    Else
        ConvertNumberToString = Replace(ConvertNumberToString, ",", ".")
    End If
'    ConvertNumberToString = Replace(Replace(n, ",", "."), "E", VBA.ChrW(183) & "10^(")
    ConvertNumberToString = Replace(ConvertNumberToString, "E", VBA.ChrW(183) & "10^(")
    If InStr(ConvertNumberToString, "10^(") Then
        ConvertNumberToString = ConvertNumberToString & ") "
    End If
    
slut:
End Function
Function ConvertNumberToStringBC(n As Double, Optional bc As Integer) As String
' konverter tal til streng med angivet antal betydende cifre. Hvis ingen angives anvendes maximacifre
    If bc > 0 Then
        ConvertNumberToStringBC = ConvertNumberToString(betcif(n, bc))
    Else
        ConvertNumberToStringBC = ConvertNumberToString(betcif(n, MaximaCifre))
    End If
End Function

Function ConvertStringToNumber(ns As String) As Double
Dim nd As Double
On Error Resume Next
    nd = CDbl(ns)
    ns = Replace(ns, ",", ".")
    ns = Replace(ns, "*10^", "E")
    nd = CDbl(ns)
    nd = val(ns)
    If Err.Number > 0 Then
       Err.Clear
        ConvertStringToNumber = Null
    Else
        ConvertStringToNumber = nd
    End If
End Function

Function ConvertNumberToMaxima(n As String) As String
' tager højde for E, men ikke helt entydigt.

    n = Replace(n, ",", ".")
    
    If InStr(n, "E+") > 0 Or InStr(n, "E-") > 0 Then
    n = Replace(n, "E-0", "*10^(-")
    n = Replace(n, "E-", "*10^(")
    n = Replace(n, "E+0", "*10^(")
    n = Replace(n, "E+", "*10^(")
    n = n & ")"
'    n = Replace(n, "E", "*10^(") & ")"
    End If
    n = omax.CodeForMaxima(n)
    ConvertNumberToMaxima = n
End Function

Sub LandScapePage()
' indsætter landscape side og alm side efter
    ActiveDocument.Range(start:=Selection.start, End:=Selection.start).InsertBreak Type:=wdSectionBreakNextPage
    Selection.start = Selection.start + 1
    With ActiveDocument.Range(start:=Selection.start, End:=Selection.start).PageSetup
        .Orientation = wdOrientLandscape
        .SectionStart = wdSectionNewPage
    End With
'    If Selection.Range.Bookmarks.Exists("\EndOfDoc") Then
    ActiveDocument.Range(start:=Selection.start, End:=Selection.start).InsertBreak Type:=wdSectionBreakNextPage
    Selection.start = Selection.start + 1
    With ActiveDocument.Range(start:=Selection.start, End:=ActiveDocument.Content.End).PageSetup
        .Orientation = wdOrientPortrait
        .SectionStart = wdSectionNewPage
    End With
'    End If

End Sub

Sub ForceError()
    Dim A As Integer
    
    A = 0 / 0
End Sub

Public Sub ClearClipBoard()
' giver desværre sjældne problemer på nogle computere
' specielt hvis der er definitioner i dokumentet så den fyres to gange
On Error GoTo slut
    Dim oData   As New DataObject 'object to use the clipboard
     
    oData.SetText text:=Empty 'Clear
    oData.PutInClipboard 'take in the clipboard to empty it
    Set oData = Nothing
slut:
End Sub

Sub GoToEndOfMath()
Dim mc As OMaths
Dim i As Integer
    Selection.Collapse wdCollapseEnd
    Set mc = Selection.OMaths
    If mc.Count > 0 Then
On Error Resume Next
    mc(mc.Count).ParentOMath.Range.Select
On Error GoTo slut
    mc(mc.Count).Range.Select  ' virker med word 2010, parentomath giver tilgengæld problemer. Hmm problem med valgt del af udtryk og reducer
    Else
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
    If r.text = VBA.ChrW(11) Then ' hvis der er shift-enter i slutningen erstattes med alm. retur
        r.text = VBA.ChrW(13)
    End If
End Sub

Function NotZero(i As Integer) As Integer
' hvis negativ returner nul
    If i < 0 Then
        NotZero = 0
    Else
        NotZero = i
    End If
End Function

Sub TabelToList()
    Dim dd As New DocData
    Dim OM As Range
    On Error GoTo fejl
    If Selection.Range.Tables.Count = 0 Then
        If Sprog.SprogNr = 1 Then
            MsgBox "Du skal markere en tabel først", vbOKOnly, "Fejl"
        Else
            MsgBox "Select a table first", vbOKOnly, "Error"
        End If
        GoTo slut
    End If
    PrepareMaxima
    dd.ReadSelectionS

    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord

    GoToInsertPoint
    'Selection.TypeParagraph
    Set OM = Selection.OMaths.Add(Selection.Range)
    Selection.TypeText dd.GetListFormS(CInt(Not (MaximaSeparator)))
    OM.OMaths(1).BuildUp
    Selection.TypeParagraph
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Sub ListToTabel()
Dim dd As New DocData
Dim Tabel As Table
Dim i As Integer, j As Integer
On Error GoTo fejl
PrepareMaxima
    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord
dd.ReadSelection
If dd.nrows = 0 Or dd.ncolumns = 0 Then
    If Sprog.SprogNr = 1 Then
        MsgBox "Du skal markere en liste først fx [1;2;3]", vbOKOnly, "Fejl"
    Else
        MsgBox "Select a list first. Example: [1;2;3]", vbOKOnly, "Error"
    End If
    GoTo slut
End If
GoToInsertPoint
Selection.TypeParagraph
'Selection.Tables.Add Selection.Range, dd.nrows, dd.ncolumns
        Set Tabel = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=dd.nrows, NumColumns:=dd.ncolumns _
        , DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed)

'        ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=dd.nrows, NumColumns:=dd.ncolumns _
'        , DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
'        wdAutoFitFixed
        With Selection.Tables(1)
'        If .Style <> "Tabel - Gitter" Then
'            .Style = "Tabel - Gitter"
'        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
        For i = 1 To Tabel.Columns.Count
            .Columns(i).Width = 65
        Next
        End With

'Set tabel = Selection.Tables.Add(Selection.Range, dd.nrows, dd.ncolumns)

For i = 1 To dd.nrows
    For j = 1 To dd.ncolumns
        Tabel.Cell(i, j).Range.text = dd.TabelsCelle(i, j)
    Next
Next

GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
    Oundo.EndCustomRecord
End Sub
Sub GoToInsertPoint()
' finder næste punkt efter selection hvor der kan indsættes nog
' dvs. går efter mathbokse og tabeller

Selection.Collapse wdCollapseEnd
If Selection.OMaths.Count > 0 Then
    omax.GoToEndOfSelectedMaths
End If
If Selection.Tables.Count > 0 Then
    Selection.Tables(Selection.Tables.Count).Select
    Selection.Collapse wdCollapseEnd
End If

End Sub

Sub ToggleDebug()
    DebugWM = Not DebugWM
End Sub
Sub SwitchLanguage()
    LanguageSetting = InputBox("angiv sprognr", "Sprog", "1")
    Sprog.CheckSetting
    RefreshRibbon
End Sub

Sub GenerateAutoCorrect()
' genererer matematisk autokorrektur
    Application.OMathAutoCorrect.UseOutsideOMath = True
    
    Application.OMathAutoCorrect.Entries.Add "\bi", VBA.ChrW(8660) ' biimplikationspile
    Application.OMathAutoCorrect.Entries.Add "\imp", VBA.ChrW(8658) ' implikationspil højre
End Sub

Sub RestartWordMat()
' genstart Maxima og genopretter doc1
RestartMaxima
LukTempDoc
'tempDoc.Close (False)
'Set tempDoc = Nothing
OpretTempdoc
End Sub

Function Get2DVector(text As String) As String
    Dim ea As New ExpressionAnalyser
'    Dim c As Collection
    Dim m As CMatrix
    ea.text = text
    
'    c = ea.GetAllMatrices()
    For Each m In ea.GetAllMatrices()
        
        Get2DVector = Get2DVector & "vector(" & m
    Next
    
End Function

Sub InsertNumberedEquation(Optional AskRef As Boolean = False)
    Dim t As Table, F As Field, ccut As Boolean
    Dim placement As Integer
    On Error GoTo fejl
    Application.ScreenUpdating = False

    If Selection.Tables.Count > 0 Then
        MsgBox "Cant insert numbered equation in table", vbOKOnly, Sprog.Error
        Exit Sub
    End If

    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord

    If Selection.OMaths.Count > 0 Then
        If Not Selection.OMaths(1).Range.text = vbNullString Then
            Selection.OMaths(1).Range.Cut
            ccut = True
            'der kan nogen gange være en rest af et matematikfelt
            If Selection.OMaths.Count > 0 Then
                If Selection.OMaths(1).Range.text = vbNullString Then
                    Selection.OMaths(1).Range.Delete
                Else
                    Selection.TypeParagraph
                End If
            End If
        End If
        If Selection.Tables.Count > 0 Then
            Selection.TypeParagraph
        End If
    End If

    Selection.Collapse wdCollapseEnd
    Set t = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=1, NumColumns:=3, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed)
#If Mac Then
#Else
    With t
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
    End With
#End If
    t.PreferredWidthType = wdPreferredWidthPercent

    t.Columns(1).PreferredWidth = 7
    t.Columns(2).PreferredWidth = 84
    t.Columns(3).PreferredWidth = 7

    t.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    t.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    t.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    t.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    t.Borders(wdBorderVertical).LineStyle = wdLineStyleNone

    'indsæt nummer
    If EqNumPlacement Then
        placement = 1
    Else
        placement = 3
    End If
    t.Cell(1, placement).Range.Select
    Selection.Collapse wdCollapseStart
    If Not EqNumType Then
        Set F = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, text:="LISTNUM ""WMeq"" ""NumberDefault"" \L 4")
        F.Update
        '        f.Code.Fields.ToggleShowCodes
    Else
        Selection.TypeText "("
        '        Set f = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, text:="SEQ chapter \c")
        Set F = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, text:="SEQ WMeq1 \c")
        '        Set f = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, text:="STYLEREF ""Overskrift 1""")
        '        Set f = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, text:="SECTION")
        F.Update
        '        f.Code.Fields.ToggleShowCodes
        Selection.TypeText "."
        '        Set f = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, text:="SEQ figure \s1") ' starter automatisk forfra ved ny overskrift 1
        Set F = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, text:="SEQ WMeq2 ")
        F.Update
        '        f.Code.Fields.ToggleShowCodes
        Selection.TypeText ")"
    End If

    If AskRef Then
        Dim EqName As String
        t.Cell(1, 3).Range.Fields(1).Select
        UserFormEnterEquationRef.Show
        EqName = UserFormEnterEquationRef.EquationName    'Replace(InputBox(Sprog.A(5), Sprog.A(4), "Eq"), " ", "")
        If EqName <> vbNullString Then
            With ActiveDocument.Bookmarks
                .Add Range:=Selection.Range, Name:=EqName
                .DefaultSorting = wdSortByName
                .ShowHidden = False
            End With
        End If
    End If

    ' indsæt mat-felt
    t.Cell(1, 2).Range.Select
    Selection.Collapse wdCollapseStart
    If ccut Then
        DoEvents
        Selection.Paste
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Else
        Selection.OMaths.Add Range:=Selection.Range
    End If

    t.Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
    t.Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    t.Cell(1, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
    t.Cell(1, 1).VerticalAlignment = wdCellAlignVerticalCenter
    t.Cell(1, 2).VerticalAlignment = wdCellAlignVerticalCenter
    t.Cell(1, 3).VerticalAlignment = wdCellAlignVerticalCenter

    ActiveDocument.Fields.Update

    Oundo.EndCustomRecord

    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub InsertEquationRef()
    Dim b As String
    On Error GoTo fejl
    UserFormEquationReference.Show
    b = UserFormEquationReference.EqName
    
    If b <> vbNullString Then
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
        Selection.TypeText Sprog.Equation & " "
#If Mac Then
        Selection.InsertCrossReference referencetype:=wdRefTypeBookmark, ReferenceKind:= _
            wdContentText, ReferenceItem:=b, InsertAsHyperlink:=False, _
            IncludePosition:=False
#Else
        Selection.InsertCrossReference referencetype:=wdRefTypeBookmark, ReferenceKind:= _
            wdContentText, ReferenceItem:=b, InsertAsHyperlink:=False, _
            IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
#End If

        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.Fields.ToggleShowCodes
        Selection.Collapse wdCollapseEnd
    
        Oundo.EndCustomRecord
    
    End If
    
    ActiveDocument.Fields.Update
    
    '    Selection.InsertCrossReference referencetype:="Bogmærke", ReferenceKind:= _
    '        wdContentText, ReferenceItem:="lign1", InsertAsHyperlink:=False, _
    '        IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "

    '    Selection.MoveLeft Unit:=wdCharacter, count:=1
    '    Selection.Fields.ToggleShowCodes
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub SetEquationNumber()
On Error GoTo fejl
    Application.ScreenUpdating = False
    Dim F As Field, f2 As Field, n As String, p As Integer, Arr As Variant
    
    If Selection.Fields.Count = 0 Then
        MsgBox Sprog.A(345), vbOKOnly, Sprog.Error
        Exit Sub
    End If
    
    Set F = Selection.Fields(1)
    If Selection.Fields.Count = 1 And InStr(F.Code.text, "LISTNUM") > 0 Then
        n = InputBox(Sprog.A(346), Sprog.A(6), "1")
        p = InStr(F.Code.text, "\S")
        If p > 0 Then
            F.Code.text = Left(F.Code.text, p - 1)
        End If
        F.Code.text = F.Code.text & "\S" & n
        F.Update
    ElseIf Selection.Fields.Count = 1 Or Selection.Fields.Count = 2 And InStr(F.Code.text, "WMeq") > 0 Then
        If Selection.Fields.Count = 2 Then
            Set f2 = Selection.Fields(2)
            n = InputBox(Sprog.A(346), Sprog.A(6), F.result & "." & f2.result)
            Arr = Split(n, ".")
            If UBound(Arr) > 0 Then
                SetFieldNo F, CStr(Arr(0))
                SetFieldNo f2, CStr(Arr(1))
            Else
                SetFieldNo F, CStr(Arr(0))
            End If
        Else
            n = InputBox(Sprog.A(346), Sprog.A(6), F.result)
            SetFieldNo F, n
        End If
        
    End If
    
    ActiveDocument.Fields.Update
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub SetFieldNo(F As Field, n As String)
    Dim p As Integer, p2 As Integer
On Error GoTo fejl
    p = InStr(F.Code.text, "\r")
    p2 = InStr(F.Code.text, "\c")
    If p2 > 0 And p2 < p Then p = p2
    If p > 0 Then
        F.Code.text = Left(F.Code.text, p - 1)
    End If
    F.Code.text = F.Code.text & "\r" & n & " \c"
    F.Update
    ActiveDocument.Fields.Update
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub InsertEquationHeadingNo()
    Dim result As Long
On Error GoTo fejl
    result = MsgBox(Sprog.A(348), vbYesNoCancel, Sprog.A(8))
    If result = vbCancel Then Exit Sub
    If result = vbYes Then
      Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, text:="SEQ WMeq1"
    Else
      Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, text:="SEQ WMeq1 \h"
    End If
      Selection.Collapse wdCollapseEnd
      Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, text:="SEQ WMeq2 \r0 \h"

    ActiveDocument.Fields.Update
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub UpdateEquationNumbers()
On Error GoTo fejl

    ActiveDocument.Fields.Update
    
    GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub CreateSprogArrays()
' to create more arrays to fill in sprog

Dim startn As Integer, no As Integer, startDepth As Integer, endDepth As Integer
Dim i As Integer, j As Integer, s As String
'Dim obj As New DataObject

startn = 181
no = 19
startDepth = 0
endDepth = 5

For i = startn To startn + no
    For j = startDepth To endDepth
        s = s & "SA(" & i & ", " & j & ") = """ & vbCrLf
    Next
'    For j = endDepth + 1 To endDepth + 5
'        s = s & "'SA(" & i & ", " & j & ") = """ & vbCrLf
'    Next
    s = s & vbCrLf
Next
Dim p As Long
p = Selection.start
Selection.Range.InsertAfter s
'Selection.start = p
'Selection.Cut
'obj.PutInClipboard

'Set obj = Nothing
End Sub

Sub SaveBackup()
    On Error GoTo fejl
    Dim Path As String
    Dim UFbackup As UserFormBackup
    Dim UfWait As UserFormWaitForMaxima
    Const lCancelled_c As Long = 0
    Dim tempDoc2 As Document
    
    
    If BackupType = 2 Or BackupAnswer = 2 Then
        Exit Sub
    ElseIf BackupType = 0 And BackupAnswer = 0 Then
        Set UFbackup = New UserFormBackup
        UFbackup.Show
        '        If MsgBox(Sprog.A(179), vbYesNo, "Backup") = vbNo Then
        If UFbackup.Backup = False Then
            BackupAnswer = 2
            Exit Sub
        Else
            BackupAnswer = 1
        End If
    End If
        
    If Timer - SaveTime < BackupTime * 60 Then Exit Sub
    SaveTime = Timer
    If ActiveDocument.Path = "" Then
        MsgBox Sprog.A(679)
        Exit Sub
    End If
    Set UfWait = New UserFormWaitForMaxima
    UfWait.Show vbModeless
    UfWait.Label_tip.Caption = "Saving backup" ' to " & VbCrLfMac & "documents\WordMat-Backup"
    UfWait.Label_progress.Caption = "*"
    DoEvents
   
    
    '    Application.ScreenUpdating = False
    If ActiveDocument.Saved = False Then ActiveDocument.Save
    UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
    DoEvents
    BackupNo = BackupNo + 1
    If BackupNo > BackupMaxNo Then BackupNo = 1
#If Mac Then
    Path = GetTempDir & "WordMat-Backup/"
#Else
    Path = GetDocumentsDir & "\WordMat-Backup\"
#End If
    '    If Dir(path, vbDirectory) = "" Then MkDir path
    If Not FileExists(Path) Then MkDir Path
    UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
    DoEvents
    Path = Path & "WordMatBackup" & BackupNo & ".docx"
    If VBA.LenB(Path) = lCancelled_c Then Exit Sub
    
#If Mac Then
    Set tempDoc2 = Application.Documents.Add(Template:=ActiveDocument.FullName, visible:=False)
    UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
    DoEvents
    tempDoc2.ActiveWindow.Left = 2000
    tempDoc2.SaveAs Path
    UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
    DoEvents
    tempDoc2.Close
#Else
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ActiveDocument.Save
    fso.CopyFile ActiveDocument.FullName, Path
    Set fso = Nothing
    UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
    DoEvents
    UfWait.Label_progress.Caption = UfWait.Label_progress.Caption & "*"
    DoEvents
#End If

    GoTo slut
fejl:
    MsgBox Sprog.A(178), vbOKOnly, Sprog.A(208)
slut:
    On Error Resume Next
    If Not UfWait Is Nothing Then Unload UfWait
    Application.ScreenUpdating = True
End Sub

Sub OpenLatexTemplate()
On Error GoTo fejl
    Documents.Add Template:=GetWordMatDir() & "WordDocs/LatexWordTemplate.dotx"
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub DeleteNormalDotm()
    Dim UserDir As String

MsgBox Sprog.A(681), vbOKOnly, ""
#If Mac Then
    Dim p As Integer
    UserDir = MacScript("return POSIX path of (path to home folder)")
    p = InStr(UserDir, "/Containers")
    UserDir = Left(UserDir, p) & "Group Containers/UBF8T346G9.Office/User Content.localized/Templates.localized"
    RunScript "OpenFinder", UserDir
#Else
    UserDir = Environ$("username")
    MsgBox "Open and delete this normal.dotm in this folder" & vbCrLf & "C:\Users\" & UserDir & "\AppData\Roaming\Microsoft\Templates"
'    shell "explorer.exe " & "C:\Users\" & UserDir & "\AppData\Roaming\Microsoft\Templates", vbNormalFocus ' Bitdefender problems
#End If
End Sub
Sub DeleteKeyboardShortcutsInNormalDotm()
' Sletter genveje til WordMat makroer der ved en fejl skulle være blevet gemt i normal.dotm
    Dim GemT As Template
    Dim KB As KeyBinding
'    On Error Resume Next
    
    Set GemT = CustomizationContext
            
    CustomizationContext = NormalTemplate

#If Mac Then ' ved en elev var KB.command helt tom, så på mac anvendes kb.keystring, selvom det er lidt mere usikkert.
    For Each KB In KeyBindings
        If Len(KB.KeyString) = 8 And Left(KB.KeyString, 7) = "Option+" Then
            KB.Clear
        End If
    Next
#Else
    For Each KB In KeyBindings
        If LCase(Left(KB.Command, 8)) = "wordmat." Then
            KB.Clear
        End If
    Next
#End If
    NormalTemplate.Save
' man kan ikke gemme WordMat.dotm som global skabelon når den er gemt for alle brugere
'    For i = 1 To Application.Documents.Count
'        arr(i) = Application.Documents(i).Saved
'        Application.Documents(i).Saved = True
'    Next
'    DoEvents
'    Documents.Save noprompt:=True, OriginalFormat:=wdOriginalDocumentFormat
'    DoEvents
'    For i = 1 To Application.Documents.Count
'        Application.Documents(i).Saved = arr(i)
'    Next
'    NormalTemplate.Save

    CustomizationContext = GemT
End Sub

Function ReadTextfileToString(FilNavn As String) As String
#If Mac Then
   Dim filnr As Integer
   filnr = FreeFile()
   Open FilNavn For Input As filnr   ' Open file
   ReadTextfileToString = Input$(LOF(1), 1)
   Close #filnr
   
#Else
   Dim fsT As Object
   'On Error GoTo fejl

   Set fsT = CreateObject("ADODB.Stream")
   fsT.Type = 2 'Specify stream type - we want To save text/string data.
   fsT.Charset = "iso-8859-1" 'Specify charset For the source text data. (Alternate: utf-8)
   fsT.Open 'Open the stream
   fsT.LoadFromFile FilNavn
   ReadTextfileToString = fsT.ReadText()
   fsT.Close
   Set fsT = Nothing
#End If

   GoTo slut
fejl:
   MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error '"Der skete en fejl i forsøget på at gemme LaTex-filen"
slut:

End Function

Sub WriteTextfileToString(FilNavn As String, WriteText As String)
#If Mac Then
   Dim filnr As Integer
   filnr = FreeFile()
   Open FilNavn For Output As filnr   ' Open file for output.
   
   Print #filnr, WriteText  ' print skriver uden " "
   Close #filnr    ' Close file.
#Else
   Dim fsT As Object
   'On Error GoTo fejl

   If FilNavn = "" Then GoTo slut
   If WriteText = "" Then
      If Dir(FilNavn) <> "" Then Kill FilNavn
         GoTo slut
   End If
   Set fsT = CreateObject("ADODB.Stream")
   fsT.Type = 2 'Specify stream type - we want To save text/string data.
   fsT.Charset = "iso-8859-1" 'Specify charset For the source text data. utf-8
   fsT.Open 'Open the stream And write binary data To the object
   fsT.WriteText WriteText
   fsT.SaveToFile FilNavn, 2 'Save binary data To disk
   fsT.Close
   Set fsT = Nothing
#End If


   GoTo slut
fejl:
   MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error '"Der skete en fejl i forsøget på at gemme LaTexfilen"
slut:

End Sub

Public Function Local_Document_Path(ByRef Doc As Document, Optional bPathOnly As Boolean = True) As String
'returns local path or nothing if local path not found. Converts a onedrive path to local path
#If Mac Then
   Local_Document_Path = Doc.Path
#Else
Dim i As Long, X As Long
Dim OneDrivePath As String
Dim ShortName As String
Dim testWbkPath As String
Dim OneDrivePathFound As Boolean

'Check if it looks like a OneDrive location
If InStr(1, Doc.FullName, "https://", vbTextCompare) > 0 Then

    'loop through three OneDrive options
    For i = 1 To 3
        'Replace forward slashes with back slashes
        ShortName = Replace(Doc.FullName, "/", "\")

        'Remove the first four backslashes
        For X = 1 To 4
            ShortName = RemoveTopFolderFromPath(ShortName)
        Next
        'Choose the version of Onedrive
        OneDrivePath = Environ(Choose(i, "OneDrive", "OneDriveCommercial", "OneDriveConsumer"))
        If Len(OneDrivePath) > 0 Then
            'Loop to see if the tentative LocalWorkbookName is the name of a file that actually exists, if so return the name
            Do While ShortName Like "*\*"
                testWbkPath = OneDrivePath & "\" & ShortName
                If Not (Dir(testWbkPath)) = vbNullString Then
                    OneDrivePathFound = True
                    Exit Do
                End If
                'remove top folder in path
                ShortName = RemoveTopFolderFromPath(ShortName)
            Loop
        End If
        If OneDrivePathFound Then Exit For
    Next i
Else
    If bPathOnly Then
        Local_Document_Path = RemoveFileNameFromPath(Doc.FullName)
    Else
        Local_Document_Path = Doc.FullName
    End If
End If
If OneDrivePathFound Then
        If bPathOnly Then
        Local_Document_Path = RemoveFileNameFromPath(testWbkPath)
    Else
        Local_Document_Path = testWbkPath
    End If
End If
#End If
End Function

Function RemoveTopFolderFromPath(ByVal ShortName As String) As String
   RemoveTopFolderFromPath = Mid(ShortName, InStr(ShortName, "\") + 1)
End Function

Function RemoveFileNameFromPath(ByVal ShortName As String) As String
   RemoveFileNameFromPath = Mid(ShortName, 1, Len(ShortName) - InStr(StrReverse(ShortName), "\"))
End Function

Function ExtractTag(s As String, StartTag As String, EndTag As String) As String
   Dim p As Long, p2 As Long
   
   p = InStr(s, StartTag)
   If p <= 0 Then
      ExtractTag = ""
      Exit Function
   End If
   p2 = InStr(p + Len(StartTag), s, EndTag)
   If p2 <= 0 Then
      ExtractTag = ""
      Exit Function
   End If
   
   ExtractTag = Mid(s, p + Len(StartTag), p2 - p - Len(StartTag))
   
End Function

Sub SetMathAutoCorrect()
' unfortunately cant be run from autoexec.
    If MaximaGangeTegn = VBA.ChrW(183) Then
        Call Application.OMathAutoCorrect.Entries.Add(Name:="*", Value:=VBA.ChrW(183))
    ElseIf MaximaGangeTegn = VBA.ChrW(215) Then
        Call Application.OMathAutoCorrect.Entries.Add(Name:="*", Value:=VBA.ChrW(215))
    Else
        On Error Resume Next
        Call Application.OMathAutoCorrect.Entries("*").Delete
    End If
End Sub

Function ConvertNumber(ByVal n As String) As String
' sørger for at streng har maximaindstilling med separatorer

If DecSeparator = "," Then
'    n = Replace(n, ",", ";")
    ConvertNumber = Replace(n, ".", ",")
Else
    ConvertNumber = Replace(n, ",", ".")
'    n = Replace(n, ";", ",")
End If

End Function
Function GetWordMatDir() As String
#If Mac Then
    GetWordMatDir = "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/"
#Else
    GetWordMatDir = GetProgramFilesDir() & "\WordMat\"
#End If
End Function

Sub NewEquation()
    Dim r As Range
    On Error GoTo fejl
    On Error Resume Next
    
    If DoubleTapM = 1 Then
        If Timer() - TapTime < 0.8 Then
            Application.Run macroname:="WMPShowFormler"
        End If
        TapTime = Timer()
    End If
    
    If Selection.OMaths.Count = 0 Then
        Set r = Selection.OMaths.Add(Selection.Range)
    ElseIf Selection.Tables.Count = 0 Then
        If Selection.OMaths(1).Range.text = vbNullString Then
            Set r = Selection.OMaths.Add(Selection.Range)
        ElseIf DoubleTapM = 2 Then
            If Not Selection.Range.ListFormat.ListValue = 0 Then
                Selection.Range.ListFormat.RemoveNumbers
            End If
            InsertNumberedEquation EqAskRef
        End If
    ElseIf Selection.Tables(1).Columns.Count = 3 And Selection.Tables(1).Cell(1, 3).Range.Fields.Count > 0 Then
        Selection.Tables(1).Cell(1, 2).Range.OMaths(1).Range.Cut
        Selection.Tables(1).Select
'        Selection.MoveEnd unit:=wdCharacter, count:=2
        Selection.Tables(1).Delete
        Selection.Paste
        Selection.TypeParagraph
        Selection.MoveLeft Unit:=wdCharacter, Count:=2
    End If
GoTo slut
fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Function FormatDefinitions(DefS As String) As String
' Tager en streng som kommer fra omax.definitions og laver den så pæn som mulig til visning i en textbox
' Bruges til visning af gældende definitioner på flere Forms
    DefS = " " & omax.ConvertToAscii(DefS)
    DefS = Replace(DefS, "$", VbCrLfMac & VbCrLfMac & " ")
    DefS = Replace(DefS, ":=", " = ")
    DefS = Replace(DefS, ":", " = ")
        
    If Not Radians Then DefS = Replace(DefS, "%pi/180*", "")
        
    DefS = Replace(DefS, " ## ", MaximaGangeTegn)
    DefS = Replace(DefS, "^^", "^")
    DefS = Replace(DefS, "SymVecta", ChrW(&H20D7))
    DefS = Replace(DefS, "matrix", vbNullString)
            
        
    DefS = Replace(DefS, "*", MaximaGangeTegn)
        
    DefS = Replace(DefS, "%pi", ChrW(&H3C0))
    DefS = Replace(DefS, "%i", "i")
    DefS = Replace(DefS, "log(", "ln(")
    DefS = Replace(DefS, "log10(", "log(")
    DefS = Replace(DefS, "^(x)", ChrW(&H2E3))
    DefS = Replace(DefS, "^(2)", ChrW(&HB2))
    DefS = Replace(DefS, "^(3)", ChrW(&HB3))
    DefS = Replace(DefS, "^(4)", ChrW(&H2074))
    DefS = Replace(DefS, "^(5)", ChrW(&H2075))
    DefS = Replace(DefS, "^(6)", ChrW(&H2076))
    DefS = Replace(DefS, "^(7)", ChrW(&H2077))
    DefS = Replace(DefS, "^(8)", ChrW(&H2078))
    DefS = Replace(DefS, "^(9)", ChrW(&H2079))
    DefS = Replace(DefS, "^(-1)", ChrW(&H207B) & ChrW(&HB9))
    DefS = Replace(DefS, "^(-2)", ChrW(&H207B) & ChrW(&HB2))
    DefS = Replace(DefS, "^(-3)", ChrW(&H207B) & ChrW(&HB3))
        
    DefS = Replace(DefS, "_0(", ChrW(&H2080) & "(")
    DefS = Replace(DefS, "_1(", ChrW(&H2081) & "(")
    DefS = Replace(DefS, "_2(", ChrW(&H2082) & "(")
    DefS = Replace(DefS, "_3(", ChrW(&H2083) & "(")
    DefS = Replace(DefS, "_4(", ChrW(&H2084) & "(")
    DefS = Replace(DefS, "_5(", ChrW(&H2085) & "(")
    DefS = Replace(DefS, "_6(", ChrW(&H2086) & "(")
    DefS = Replace(DefS, "_7(", ChrW(&H2087) & "(")
    DefS = Replace(DefS, "_8(", ChrW(&H2088) & "(")
    DefS = Replace(DefS, "_9(", ChrW(&H2089) & "(")
    DefS = Replace(DefS, "_a(", ChrW(&H2090) & "(")
    DefS = Replace(DefS, "_x(", ChrW(&H2093) & "(")
    DefS = Replace(DefS, "_n(", ChrW(&H2099) & "(")
        
    DefS = Replace(DefS, "minf", "-" & ChrW(&H221E))
    DefS = Replace(DefS, "inf", ChrW(&H221E))
        
    DefS = Replace(DefS, "sqrt(", ChrW(&H221A) & "(")
    DefS = Replace(DefS, "NIntegrate(", ChrW(&H222B) & "(")
    DefS = Replace(DefS, "Integrate(", ChrW(&H222B) & "(")
    DefS = Replace(DefS, "integrate(", ChrW(&H222B) & "(")
    DefS = Replace(DefS, "<=", VBA.ChrW(8804))
    DefS = Replace(DefS, ">=", VBA.ChrW(8805))
    DefS = Replace(DefS, "ae", "æ")
    DefS = Replace(DefS, "oe", "ø")
    DefS = Replace(DefS, "aa", "å")
    DefS = Replace(DefS, "AE", "Æ")
    DefS = Replace(DefS, "OE", "Ø")
    DefS = Replace(DefS, "AA", "Å")
        
    'græske bogstaver
    DefS = Replace(DefS, "gamma", VBA.ChrW(915))    ' stort gammategn
    DefS = Replace(DefS, "Delta", VBA.ChrW(916))
    DefS = Replace(DefS, "delta", VBA.ChrW(948))
    DefS = Replace(DefS, "alpha", VBA.ChrW(945))
    DefS = Replace(DefS, "beta", VBA.ChrW(946))
    DefS = Replace(DefS, "gammaLB", VBA.ChrW(947))
    DefS = Replace(DefS, "theta", VBA.ChrW(952))
    DefS = Replace(DefS, "Theta", VBA.ChrW(920))
    DefS = Replace(DefS, "lambda", VBA.ChrW(955))
    DefS = Replace(DefS, "Lambda", VBA.ChrW(923))
    DefS = Replace(DefS, "mu", VBA.ChrW(956))
    DefS = Replace(DefS, "rho", VBA.ChrW(961))
    DefS = Replace(DefS, "sigma", VBA.ChrW(963))
    DefS = Replace(DefS, "Sigma", VBA.ChrW(931))
    DefS = Replace(DefS, "varphi", VBA.ChrW(966))
    DefS = Replace(DefS, "phi", VBA.ChrW(981))
    DefS = Replace(DefS, "Phi", VBA.ChrW(934))
    DefS = Replace(DefS, "varepsilon", VBA.ChrW(949))
    DefS = Replace(DefS, "epsilon", VBA.ChrW(1013))
    DefS = Replace(DefS, "psi", VBA.ChrW(968))
    DefS = Replace(DefS, "Psi", VBA.ChrW(936))
    DefS = Replace(DefS, "Xi", VBA.ChrW(926))
    DefS = Replace(DefS, "xi", VBA.ChrW(958))
    DefS = Replace(DefS, "Chi", VBA.ChrW(935))
    DefS = Replace(DefS, "chi", VBA.ChrW(967))
    DefS = Replace(DefS, "Pi", VBA.ChrW(928))
    DefS = Replace(DefS, "tau", VBA.ChrW(964))
    DefS = Replace(DefS, "greek-nu", VBA.ChrW(957))
    DefS = Replace(DefS, "kappa", VBA.ChrW(954))
    DefS = Replace(DefS, "eta", VBA.ChrW(951))
    DefS = Replace(DefS, "zeta", VBA.ChrW(950))
    DefS = Replace(DefS, "omega", VBA.ChrW(969))    ' lille omega
    
    DefS = Replace(DefS, "((x))", "(x)")
        
    
    If DecSeparator = "," Then
        '        DefS = Replace(DefS, ",", ";")
        DefS = Replace(DefS, ".", ",")
    End If
        
    FormatDefinitions = DefS
End Function

Function MsgBox2(prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKCancel, Optional Title As String) As VbMsgBoxResult
' erstatning for indbygget msgbox. Der bruger samme stil som resten af Userforms. Den tilpasser sig i størrelse.
' Buttons understøttes: vbYesNo, vbOKonly, vbOKCancel
' MsgBox2 "Dette er en lille test", vbOKOnly, "Hello"

    Dim UFMsgBox As New UserFormMsgBox
    
    UFMsgBox.MsgBoxStyle = Buttons
    UFMsgBox.Title = Title
    UFMsgBox.prompt = prompt
    
    UFMsgBox.Show
    
    MsgBox2 = UFMsgBox.MsgBoxResult
    
    Unload UFMsgBox
End Function

Sub TestMe()
 MsgBox2 "Dette er en lille test", vbOKOnly, "Hello"
 MsgBox2 "Dette er en længere test" & vbCrLf & "Der skal være flere og længere linjer" & vbCrLf & "hej." & vbCrLf & "hej." & vbCrLf & "hej." & vbCrLf & "hej.", vbOKCancel, "Hello"
 MsgBox2 "Dette er en bred test" & vbCrLf & "Der skal være en meget lang linje med mange forskellige tegn, så boksen bliver bred. Mon den kan blive så bred som denne linje?" & vbCrLf & "hej." & vbCrLf & "hej." & vbCrLf & "hej." & vbCrLf & "hej.", vbOKCancel, "Hello"
End Sub

Sub TestError()
    On Error Resume Next
    Err.Raise 1, , "dsds"
End Sub

