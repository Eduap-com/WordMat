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

Sub SolveCantSaveProblem()
    ' Forsøger at
    ' finder to gentagne softreturns og reducerer til 1
    Application.ScreenUpdating = False
    Dim resultat As VbMsgBoxResult
    resultat = MsgBox("Word 2007 har en fejl der gør det umuligt at gemme dokumentet under specielle omstændigheder. Problemet opstår ved en kombination af ligninger og shift-enter, men kun i specielle tilfælde." & vbCrLf & vbCrLf & " Hvis du ikke kan gemme dit dokument kan denne funktion måske finde fejlen i dette dokument og rette det" & vbCrLf & "Du kan altid gemme dokumentet i Word 2003 format og så senere konvertere tilbage til 2007 format" & vbCrLf & "Tryk OK for at rette fejlen.", vbOKCancel, "Hjælp jeg kan ikke gemme")
    If resultat = vbOK Then
            
        ActiveDocument.OMaths.Linearize
        ActiveDocument.OMaths.BuildUp
        ' skal udføres to gange ellers kan der godt stadig være dobbelt hvis der kommer flere i træk
        '        Call Selection.Range.Find.Execute(VBA.ChrW(11) & VBA.ChrW(11), , , , , , , , , VBA.ChrW(11) & " " & VBA.ChrW(11), wdReplaceAll)
        '        Call Selection.Range.Find.Execute(VBA.ChrW(11) & VBA.ChrW(11), , , , , , , , , VBA.ChrW(11) & " " & VBA.ChrW(11), wdReplaceAll)
    End If

End Sub
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
#If Mac Then
    Call tempDoc
#Else
Dim d As Document
If tempDoc Is Nothing Then
For Each d In Application.Documents
    If d.BuiltInDocumentProperties("Title") = "MMtempDoc" Then
        Set tempDoc = d
        Exit For
    End If
Next
End If

If tempDoc Is Nothing Then
    Set tempDoc = Documents.Add(, , , False)
    tempDoc.BuiltInDocumentProperties("Title") = "MMtempDoc"
End If

If Not tempDoc.BuiltInDocumentProperties("Title") = "MMtempDoc" Then
    tempDoc.Close SaveChanges:=wdDoNotSaveChanges
    Set tempDoc = Documents.Add(, , , False)
    tempDoc.BuiltInDocumentProperties("Title") = "MMtempDoc"
End If
#End If
End Sub
Sub ShowTempDoc()
    MsgBox tempDoc.Range.Text
End Sub
Sub LukTempDoc()
On Error GoTo slut
#If Mac Then
    If Not m_tempDoc Is Nothing Then
        If Word.Application.IsObjectValid(m_tempDoc) Then
            m_tempDoc.Close (False)
        End If
    End If
    Set m_tempDoc = Nothing
#Else
    tempDoc.Close
'    tempDoc.ActiveWindow
    Set tempDoc = Nothing ' added v. 1.11
#End If
slut:
End Sub

#If Mac Then
Function tempDoc() As Document
    'Mac: User may have closed the document, so the variable tempDoc is now a function
    Dim farRight As Integer
    On Error Resume Next
'    farRight = ScreenWidth - 1 ' just inside the screen hides most
    If Not m_tempDoc Is Nothing Then
        If Word.Application.IsObjectValid(m_tempDoc) Then
            If m_tempDoc.ActiveWindow.Left <> farRight Then m_tempDoc.ActiveWindow.Left = farRight
            Set tempDoc = m_tempDoc
            Exit Function
        Else
            Set m_tempDoc = Nothing
        End If
    End If
    Dim activeDoc As Document
    Set activeDoc = ActiveDocument
    
' men kun hvis ikke eksisterer allerede
Dim d As Document
If m_tempDoc Is Nothing Then
For Each d In Application.Documents
'    If d.BuiltInDocumentProperties("Title") = "MMtempDoc" Then
    If d.ActiveWindow.Caption = "WordMatTempDoc" Then
        Set m_tempDoc = d
        Exit For
    End If
Next
End If

If m_tempDoc Is Nothing Then
    Set m_tempDoc = Documents.Add(, , , False)
    m_tempDoc.ActiveWindow.Left = farRight
'    m_tempDoc.BuiltInDocumentProperties("Title") = "MMtempDoc" ' på mac gav denne problemer. Der blev skiftet fokus til tempdoc nogle sekunder senere. Måske fordi den er meget langsom
    'Mac: Visible=False?
    m_tempDoc.ActiveWindow.Caption = "WordMatTempDoc"
    
    m_tempDoc.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = Sprog.A(680) '"Do NOT edit this document or close or it. WordMat needs it for calculations. Anything you enter here will be deleted."
    'Note: Update 14.2.5 for Office 2011 allows document to be placed outside screen
    'm_tempDoc.ActiveWindow.WindowState = wdWindowStateMinimize
    m_tempDoc.Saved = True
End If

' fjernet 26/1-17
'If Not m_tempDoc.BuiltInDocumentProperties("Title") = "MMtempDoc" Then
'    m_tempDoc.Close SaveChanges:=wdDoNotSaveChanges
'    m_tempDoc.ActiveWindow.Left = farRight
'    Set m_tempDoc = Documents.Add(, , , False)
'    m_tempDoc.BuiltInDocumentProperties("Title") = "MMtempDoc"
'
'    'Mac: Visible=False?
'    m_tempDoc.ActiveWindow.Caption = "WordMatTempDoc"
'    'Note: Update 14.2.5 for Office 2011 allows document to be placed outside screen
'    'm_tempDoc.ActiveWindow.WindowState = wdWindowStateMinimize
'    m_tempDoc.Saved = True
'End If
    'Mac:
    Set tempDoc = m_tempDoc
    If Not activeDoc Is Nothing Then activeDoc.Activate
End Function
Sub SetTempDocSaved()
    On Error Resume Next
    m_tempDoc.Saved = True
End Sub
#End If

Sub ActivateTask(navn As String)
AppActivate navn
Exit Sub

Dim task1 As Task
Dim tasksave As Task
For Each task1 In Tasks
    If InStr(task1.Name, navn) > 0 Then
        Set tasksave = task1
        Exit For
    End If
Next
'Call AppActivate("Word", True)
Dim i As Integer
On Error GoTo start
start:
Err.Clear
i = i + 1
Wait (0.1)
If i > 2 Then GoTo slut
tasksave.Activate

slut:
End Sub

Sub Wait(pausetime As Variant)
'pausetime in milliseconds
Dim start
    start = Timer    ' Set start time.
    Do While Timer < start + pausetime
        DoEvents    ' Yield to other processes.
    Loop

End Sub

Function MakeMMathCompatible(ut As String) As String
    ut = Replace(ut, ",", ".")
    ut = Replace(ut, "E", VBA.ChrW(183) & "10^ ")
    MakeMMathCompatible = ut
End Function

Sub testwait()
    Wait (3000)
End Sub

Sub ChangeAutoHyphen()
'
    Options.AutoFormatAsYouTypeReplaceFarEastDashes = False
    Options.AutoFormatAsYouTypeReplaceSymbols = False
End Sub
Public Sub GenerateKeyboardShortcuts()
    Dim Wd As WdKey
    CustomizationContext = ActiveDocument.AttachedTemplate
On Error Resume Next
'#If Mac Then
'    Wd = wdKeyControl
'#Else
    Wd = wdKeyAlt
'#End If
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyG, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="Gange"
        
If Sprog.SprogNr = 1 Then
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyB, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="beregn"
Else
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyC, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="beregn"
End If

#If Mac Then
#Else
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyReturn, Wd, wdKeyControl), KeyCategory:= _
        wdKeyCategoryCommand, Command:="beregn"
#End If

If Sprog.SprogNr = 1 Then
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyL, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="MaximaSolve"
Else
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyE, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="MaximaSolve"
End If
    
If Sprog.SprogNr = 1 Then
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyS, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="InsertSletDef"
Else
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="InsertSletDef"
End If
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyD, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="InsertDefiner"
    
If Sprog.SprogNr = 1 Then
'#If Mac Then ' alt+i bruges til numerisk tegn på mac, så hellere ikke genvej til indstillinger
'#Else
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyJ, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="MaximaSettings"
'#End If
Else
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyO, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="MaximaSettings"
End If
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyP, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="StandardPlot"
        
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyM, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="NewEquation"
        
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyR, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="ForrigeResultat"
        
If Sprog.SprogNr = 1 Then
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyE, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="ToggleUnits"
Else
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyU, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="ToggleUnits"
End If
    
If Sprog.SprogNr = 1 Then
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyO, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="Omskriv"
Else
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyS, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="Omskriv"
End If
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyN, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="ToggleNum"
        
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyT, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="ToggleLatex"
        
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyQ, Wd), KeyCategory:= _
        wdKeyCategoryCommand, Command:="SaveDocToLatexPdf()"
        

End Sub


Function GetProgramFilesDir() As String
' bruges ikke af maxima mere da det er dll-filen der står for det nu.
' bruges af de Worddokumenter mm. der skal findes
'MsgBox GetProgFilesPath
On Error GoTo Fejl
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
     GetProgramFilesDir = RegKeyRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\ProgramFilesDir (x86)")
 End If
 If Dir(GetProgramFilesDir & "\WordMat", vbDirectory) = "" Then
     GetProgramFilesDir = Environ("ProgramFiles")
 End If
 ProgramFilesDir = GetProgramFilesDir
 End If

#End If

GoTo slut
Fejl:
    MsgBox Sprog.A(110), vbOKOnly, Sprog.Error
slut:
'MsgBox GetProgramFilesDir
End Function
Function GetDocumentsDir() As String
On Error GoTo Fejl
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
Fejl:
    MsgBox Sprog.A(110), vbOKOnly, Sprog.Error
slut:
'MsgBox GetProgramFilesDir
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
Sub OpenLink(Link As String, Optional Script As Boolean = False)
On Error Resume Next

#If Mac Then
    If Script Then
        RunScript "OpenLink", Link
    Else
        ActiveDocument.FollowHyperlink Address:=Link, NewWindow:=True
    End If
#Else
    If Script Then
        Shell """" & GetProgramFilesDir & "\Microsoft\Edge\Application\msedge.exe"" """ & Link & """", vbNormalFocus
    Else
        ActiveDocument.FollowHyperlink Address:=Link, NewWindow:=True
    End If
#End If
Fejl:
End Sub

 Sub TestDll()
'Dim mp As New MaximaProcessClass
'Dim mp As New MathMenu.MaximaProcessClass
   
'    mp.ExecuteMaximaCommand "2+3;", 1
'    MsgBox mp.LastMaximaOutput

End Sub
#If Mac Then
#Else
Sub TestDll2()
' dll skal ligge i samme mappe som programmet (Word)
' navnet skal være navnet på klassen og metoder skal være com-visible
' Med denne metode kan man dog ikke bruge intellisense, men den er måske nemmere at distribuere
' man kan måske registrere på udviklingsmaskinen og så ændre til object når distribueres

If MaxProc Is Nothing Then
    Set MaxProc = GetMaxProc() ' CreateObject("MaximaProcessClass")
End If
    MaxProc.ExecuteMaximaCommand "2+3;", 1
    MsgBox MaxProc.MaximaOutput

Fejl:

End Sub
#End If
Sub InsertSletDef()
    Dim tdefs As String
    Dim gemfontsize As Integer
    Dim gemitalic As Boolean
    Dim gemfontcolor As Integer
    Dim gemsb As Integer
    Dim gemsa As Integer
    Dim mo As Range
#If Mac Then
#Else
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
#End If
            
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
#If Mac Then
#Else
        Oundo.EndCustomRecord
#End If
End Sub


Sub InsertDefiner()
    On Error GoTo Fejl

    Application.ScreenUpdating = False
    Selection.InsertAfter (Sprog.A(62) & ": ")
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths.BuildUp
'    Selection.OMaths(1).BuildUp
    Selection.Collapse wdCollapseEnd
    
    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub ForrigeResultat()

    Dim ra As Range
    Dim sr As Range
    Dim r As Range
    Dim s As String
    Dim start As Integer
    Dim sslut As Integer
    Dim posligmed As Integer
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
    On Error GoTo Fejl
    start = Selection.Range.start
    sslut = Selection.Range.End
    Set ra = ActiveDocument.Range
    ra.End = sslut + 1
    matfeltno = ra.OMaths.Count
    Do
        If ResFeltIndex >= matfeltno - 1 Then
            If ActiveDocument.Range.OMaths(matfeltno).Range.Text = Selection.Range.Text Then
                Selection.Text = ""
                Selection.OMaths.Add Range:=Selection.Range
            Else
                Selection.Text = ""
            End If
            GoTo Fejl
        End If
'        ActiveDocument.Range.OMaths(matfeltno - 1 - ResFeltIndex).Range.Select
        Set r = ActiveDocument.Range.OMaths(matfeltno - 1 - ResFeltIndex).Range
        If Len(r.Text) = 0 Then
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
    If Selection.Range.Text = "Skriv ligningen her." Then
        ResPos1 = ResPos1 - 1 ' hvis tom i forvejen er selection af eller anden grund 1 tegn for meget
    End If
    s = Replace(s, VBA.ChrW(8289), "") ' funktionstegn  sin(x) bliver ellers til si*n(x). også problem med andre funktioner
    Selection.Text = s
    
'    Dim ml As Integer
'    ml = Len(ActiveDocument.Range.OMaths(matfeltno).Range.text)
'    Selection.TypeText s
'    ResPos2 = Selection.start
'    ActiveDocument.Range.OMaths(ra.OMaths.Count).BuildUp
'    ResPos2 = ResPos1 + Len(ActiveDocument.Range.OMaths(matfeltno).Range.text) - ml
GoTo slut
Fejl:
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

Function KlipTilLigmed(Text As String, ByVal indeks As Integer) As String
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
    posligmed = InStr(Text, "=")
    possumtegn = InStr(Text, VBA.ChrW(8721))
    posca = InStr(Text, VBA.ChrW(8776))
    poseller = InStr(Text, VBA.ChrW(8744))
    
    Pos = Len(Text)
'    pos = posligmed
    If posligmed > 0 And posligmed < Pos Then Pos = posligmed
    If posca > 0 And posca < Pos Then Pos = posca
    If poseller > 0 And poseller < Pos Then Pos = poseller
    
    If possumtegn > 0 And possumtegn < Pos Then ' hvis sumtegn er der =tegn som del deraf
        Pos = 0
    End If
    If Pos = Len(Text) Then Pos = 0
    If Pos > 0 Then
        Arr(i) = Left(Text, Pos - 1)
        Text = right(Text, Len(Text) - Pos)
        i = i + 1
    Else
        Arr(i) = Text
    End If
    Loop While Pos > 0
    
    If indeks = i Then ResIndex = -1  ' global variabel markerer at der ikke er flere til venstre
    If i = 0 Then
        KlipTilLigmed = Text
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
    
    ReadEquationFast = sr.OMaths(1).Range.Text
    
    Selection.OMaths(1).ConvertToMathText
    Selection.OMaths(1).Range.Select
    Selection.OMaths.BuildUp
    Selection.Collapse (wdCollapseEnd)

    sr.Select
    
End Function

Sub testdef()
Dim ea As New ExpressionAnalyser
Dim i As Integer
ea.Text = "f(x)=x^2;a=3;b=a;c=[1;4;7];f(x;y)=x*y"
Do
    MsgBox ea.GetNextListItem(10)
    i = i + 1
Loop While i < 10
End Sub
Sub OpenFormulae(Filnavn As String)
On Error GoTo Fejl
#If Mac Then
    Documents.Open "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/WordDocs/" & Filnavn
#Else
    OpenWordFile "" & Filnavn
#End If
GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Sub OpenWordFile(Filnavn As String)
' OpenWordFile ("Figurer.docx")

Dim filnavn1 As String
#If Mac Then
    Filnavn = Replace(Filnavn, "\", "/")
    filnavn1 = GetWordMatDir() & "WordDocs/" & Filnavn
    Documents.Open filnavn1
#Else
Dim filnavn2 As String
Dim appdir As String
Dim fs
On Error GoTo Fejl
Set fs = CreateObject("Scripting.FileSystemObject")
appdir = Environ("AppData")
filnavn1 = appdir & "\WordMat\WordDocs\" & Filnavn
filnavn2 = GetProgramFilesDir & "\WordMat\WordDocs\" & Filnavn

If Dir(filnavn1) = "" And Dir(filnavn2) <> "" Then
    If Dir(appdir & "\WordMat\WordDocs\", vbDirectory) = "" Then MkDir appdir & "\WordDocs\WordMat"
    fs.CopyFile filnavn2, appdir & "\WordMat\WordDocs\"
End If

If Dir(filnavn1) <> "" Then
    Documents.Open FileName:=filnavn1
ElseIf Dir(filnavn2) <> "" Then
    Documents.Open FileName:=filnavn2, ReadOnly:=True
Else
    MsgBox Sprog.A(111) & Filnavn, vbOKOnly, Sprog.Error
End If
#End If

GoTo slut
Fejl:
    MsgBox Sprog.A(111) & Filnavn, vbOKOnly, Sprog.Error
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
                MsgBox Sprog.A(343), vbOKOnly, Sprog.Error
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
    
#If Mac Then
#Else
    Exit Sub ' overtaget af maxprocunit
#End If

Dim Text As String
    
    MaxProc.Units = 1
    Text = omax.KillDef
    If Len(Text) > 0 Then
         Text = Left(Text, Len(Text) - 1) 'fjern sidste komma
         Text = "kill(" & Text & ")"
         omax.KillDef = ""
    Else
        Text = "" ' mærkeligt men len(text)=0 er ikke nødv ""
    End If
    
    
    
'    MaxProc.ExecuteMaximaCommand text, 0
'    MaxProc.OutUnits = omax.ConvertUnits(OutUnits)
'     MaxProc.TurnUnitsOn text, ""
'     MaxProc.TurnUnitsOn "", ""


'Dim text As String
   Text = "[" & Text & "load(WordMatUnitAddon)"
'    text = "[" & text & "keepfloat:false,usersetunits:[N,J,W,Pa,C,V,F,Ohm,T,H,K],load(unit)"
    If OutUnits <> "" Then
        Text = Text & ",setunits(" & omax.ConvertUnits(OutUnits) & ")"
    End If
    Text = Text & "]$"
    
    MaxProc.ExecuteMaximaCommand Text, 0

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
    Dim Text As String
    Text = "setunits(" & omax.ConvertUnits(OutUnits) & ")$"
    
    MaxProc.ExecuteMaximaCommand Text, 0

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
    On Error GoTo Fejl
#If Mac Then
    MsgBox "Automatic update is not (yet) available on Mac" & vbCrLf & "Current version is: " & AppVersion & vbCrLf & vbCrLf & "Remember the version no. above. You will now be send to the download page where you can check for a newer version -  www.eduap.com/WordMat/Download.aspx"
    OpenLink "http://www.eduap.com/WordMat/Download.aspx"
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
            OpenLink "http://eduap.com/da/download-wordmat/" ' "http://www.eduap.com/wordmat/download.aspx"
        End If
    End If

#End If
GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Sub CheckForUpdate()
#If Mac Then
    CheckForUpdateF False
#Else
    CheckForUpdateWindows False
#End If
End Sub
Sub CheckForUpdateF(Optional Silent As Boolean = False)
    ' Create a WebClient for executing requests
    ' and set a base url that all requests will be appended to
    Dim p As Long, p2 As Long, p3 As Long, s As String, v As String, News As String
    Dim MapsClient As New WebClient
'   On Error GoTo fejl
'#If Mac Then
'    MsgBox "Automatic update is not (yet) available on Mac" & vbCrLf & "Current version is: " & AppVersion & vbCrLf & vbCrLf & "Remember the version no. above. You will now be send to the download page where you can check for a newer version -  eduap.com"
'    OpenLink "http://eduap.com/download-wordmat/"
'#Else
    Dim result As VbMsgBoxResult
    MapsClient.BaseUrl = "http://www.eduap.com/wordmat-version-history/"

    ' Use GetJSON helper to execute simple request and work with response
    Dim Resource As String
    Dim Response As WebResponse
    Dim Request As New WebRequest
    '    Request.Resource = "index.html"
    
    Request.Format = WebFormat.PlainText

    Request.Method = WebMethod.HttpGet
    '    Request.ResponseFormat = PlainText
    '    Request.Method = WebMethod.HttpGet
    '    Request.Method = WebMethod.HttpPost
'    Resource = "" '
    '    "directions/json?" & _
    '        "origin=" & Origin & _
    '        "&destination=" & Destination & _
    '        "&sensor=false"
    
    '    Set Response = MapsClient.GetJson(Resource)
    Set Response = MapsClient.Execute(Request)
    ' => GET https://maps.../api/directions/json?origin=...&destination=...&sensor=false
    '    MsgBox Response.StatusCode & " - " & Response.StatusDescription

    If Response.StatusCode = WebStatusCode.OK Or Response.StatusCode = 301 Then ' af ukendte årsager kommer der 301 fejl på mac, men det virker
        '        MsgBox Response.Content
'        p = InStr(s, "Version history")
        
'        v = Trim(Mid(s, p + 16, 4))
        '        MsgBox Response.Content
        s = Response.Content
        p = InStr(s, "<body")
        p = InStr(p, s, "Version ")
        If p <= 0 Then GoTo Fejl
        v = Trim(Mid(s, p + 8, 4))
        p2 = InStr(p + 10, s, "Version " & AppVersion)
        If p2 <= 0 Then p2 = InStr(p + 10, s, "Version")
        If p2 <= 0 Then p2 = p + 50
'        p3 = InStr(p, s, "<p>")
        News = Mid(s, p, p2 - p)
'        News = Replace(News, "- ", vbCrLf & "- ")
        News = Replace(News, "&#8211;", vbCr & " -") ' bindestreg
        News = Replace(News, "Version ", vbCrLf & "Version ") ' bindestreg
        News = Replace(News, "<br />", "")
        News = Replace(News, "<strong>", "")
        News = Replace(News, "</strong>", "")
        News = Replace(News, "<p>", "")
        News = Replace(News, "</p>", "")
        If v = AppVersion Then
            If Not Silent Then
                MsgBox Sprog.A(344) & " " & AppNavn, vbOKOnly, Sprog.OK
            End If
        Else
            result = MsgBox(Sprog.A(21) & News & vbCrLf & Sprog.A(22), vbYesNo, Sprog.A(23))
            If result = vbYes Then
                OpenLink "http://eduap.com/download-wordmat/"
            End If
        End If
    Else
        GoTo slut
    End If
    
    If Response.StatusCode = WebStatusCode.OK Or Response.StatusCode = 301 Then
    End If
    
'#End If
    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Sub CheckForUpdateWindows(Optional RunSilent As Boolean = False)
    On Error GoTo Fejl
    Dim NewVersion As String, p As Integer, p2 As Integer, News As String, s As String, v As String
    Dim Filnavn As String, FilDir As String, FilPath As String, result As VbMsgBoxResult
    
    If GetInternetConnectedState = False Then
        If Not RunSilent Then MsgBox "Ingen internetforbindelse", vbOKOnly, "Fejl"
        Exit Sub
    End If
   
    s = GetHTML("http://www.eduap.com/wordmat-version-history/")
    If Len(s) = 0 Then
        If Not RunSilent Then
            MsgBox "Serveren kan ikke kontaktes", vbOKOnly, "Fejl"
            GoTo slut
        End If
    End If
    
            p = InStr(s, "<body")
        p = InStr(p, s, "Version ")
        If p <= 0 Then GoTo Fejl
        v = Trim(Mid(s, p + 8, 4))
        p2 = InStr(p + 10, s, "Version " & AppVersion)
        If p2 <= 0 Then p2 = InStr(p + 10, s, "Version")
        If p2 <= 0 Then p2 = p + 50
'        p3 = InStr(p, s, "<p>")
        News = Mid(s, p, p2 - p)
'        News = Replace(News, "- ", vbCrLf & "- ")
        News = Replace(News, "&#8211;", vbCr & " -") ' bindestreg
        News = Replace(News, "Version ", vbCrLf & "Version ") ' bindestreg
        News = Replace(News, "<br />", "")
        News = Replace(News, "<strong>", "")
        News = Replace(News, "</strong>", "")
        News = Replace(News, "<p>", "")
        News = Replace(News, "</p>", "")

    If Len(v) = 0 Then
        If Not RunSilent Then
            MsgBox "Serveren kan ikke kontaktes", vbOKOnly, "Fejl"
            GoTo slut
        End If
    End If

   
    If AppVersion <> v Then
        '      If UFreminder.Visible = True Then UFreminder.Top = 100
        result = MsgBox(Sprog.A(21) & News & vbCrLf & Sprog.A(22), vbYesNo, Sprog.A(23))
        If result = vbYes Then
            OpenLink "http://eduap.com/download-wordmat/"
        End If
    Else
        If Not RunSilent Then
            MsgBox "Du har allerede den nyeste version installeret", vbOKOnly, "Ingen opdatering"
        End If
    End If
   
    GoTo slut
Fejl:
    '   MsgBox "Fejl " & Err.Number & " (" & Err.Description & ") i procedure CheckForUpdate, linje " & Erl & ".", vbOKOnly Or vbCritical Or vbSystemModal, "Fejl"
    If Not RunSilent Then
        MsgBox "Der skete en fejl i forbindelse at checke for ny version. Det kan skyldes en fejl med internetforbindelsen eller en fejl med serveren. Prøv igen senere, eller check selv på eduap.com om der er kommet en ny version. Den nuværende version er " & AppVersion, vbOKOnly Or vbCritical Or vbSystemModal, "Fejl"
    End If
slut:

End Sub
Sub CheckForUpdateSilentOld()
' maxproc skal være oprettet
#If Mac Then
#Else
    Dim nyversion As String, News As String
    Dim result As VbMsgBoxResult
    On Error GoTo Fejl
    nyversion = MaxProc.CheckForUpdate()
    If nyversion = "" Then
        Exit Sub
    End If

    If nyversion <> AppVersion Then
        News = MaxProc.GetVersionNews()
        result = MsgBox(Sprog.A(21) & News & vbCrLf & vbCrLf & Sprog.A(22), vbYesNo, Sprog.A(23))
        If result = vbYes Then
            OpenLink ("http://www.eduap.com/wordmat/download.aspx")
        End If
    End If


GoTo slut
Fejl:
'    MsgBox "Der kunne ikke oprettes forbindelse til serveren", vbOKOnly, "Fejl"
slut:
#End If
End Sub
Sub CheckForUpdateSilent()
' maxproc skal være oprettet
    On Error GoTo Fejl
#If Mac Then
    CheckForUpdateF True
#Else
    CheckForUpdateWindows True
#End If
GoTo slut
Fejl:
'    MsgBox "Der kunne ikke oprettes forbindelse til serveren", vbOKOnly, "Fejl"
slut:
End Sub
Function GetHTML(Url As String) As String
    Dim html As String
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", Url & "?cb=" & Timer() * 100, False  ' timer sikrer at det ikke er cached version
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
Dim ea As New ExpressionAnalyser

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

Function TrimR(ByVal Text As String, c As String)
' fjerner c fra højre side af text
Dim s As String
If Text = "" Then GoTo slut
Do While right(Text, 1) = c
    Text = Left(Text, Len(Text) - 1)
Loop
TrimR = Text
slut:
End Function
Function TrimL(ByVal Text As String, c As String)
' fjerner c fra venstre side af text
Dim s As String
If Text = "" Then GoTo slut
Do While Left(Text, 1) = c
    Text = right(Text, Len(Text) - 1)
Loop
TrimL = Text
slut:
End Function

Function TrimB(ByVal Text As String, c As String)
' fjerner c fra Begge sider af text

TrimB = TrimL(Text, c)
TrimB = TrimR(TrimB, c)
slut:
End Function
Function TrimRenter(ByVal Text As String)
' removes crlf at right end
    TrimRenter = TrimR(TrimR(Text, vbLf), vbCr)
End Function
Sub ForceError()
    Dim A As Integer
    
    A = 0 / 0
End Sub

Public Sub ClearClipBoard()
' giver desværre sjældne problemer på nogle computere
' specielt hvis der er definitioner i dokumentet så den fyres to gange
On Error GoTo slut
    Dim oData   As New DataObject 'object to use the clipboard
     
    oData.SetText Text:=Empty 'Clear
    oData.PutInClipboard 'take in the clipboard to empty it
    Set oData = Nothing
slut:
End Sub

Sub testt()
MsgBox Application.International(wdProductLanguageID)

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
    If r.Text = VBA.ChrW(11) Then ' hvis der er shift-enter i slutningen erstattes med alm. retur
        r.Text = VBA.ChrW(13)
    End If
End Sub

Sub KonverterTilLaTex()

    PrepareMaxima
'    omax.ReadSelection
    Dim uflatex As New UserFormLatex
    uflatex.Show
    
End Sub
Sub ToggleLatex()
Dim mtext As String
Dim r As Range
On Error GoTo slut
#If Mac Then
#Else
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
#End If
    If Selection.OMaths.Count > 0 Then
        PrepareMaxima
        omax.ReadSelection
        Selection.OMaths(1).Range.Text = ""
        Selection.InsertAfter LatexStart & omax.ConvertToLatex(omax.Kommando) & LatexSlut
    Else
        PrepareMaxima
        
        mtext = omax.ConvertLatexToWord(RemoveLatexOmslut(Selection.Range.Text))
        Selection.Range.Delete
        Selection.Collapse wdCollapseEnd
        Set r = Selection.OMaths.Add(Selection.Range)
        Selection.TypeText mtext
        r.OMaths(1).BuildUp
        Selection.TypeParagraph
    End If
#If Mac Then
#Else
        Oundo.EndCustomRecord
#End If

slut:
End Sub
Function RemoveLatexOmslut(Text As String)

    Text = TrimB(Text, "$")
    Text = TrimL(Text, "\[")
    Text = TrimR(Text, "\]")
    RemoveLatexOmslut = Text
End Function
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
Dim om As Range
On Error GoTo Fejl
PrepareMaxima
dd.ReadSelectionS

GoToInsertPoint
'Selection.TypeParagraph
Set om = Selection.OMaths.Add(Selection.Range)
Selection.TypeText dd.GetListFormS(CInt(Not (MaximaSeparator)))
om.OMaths(1).BuildUp
Selection.TypeParagraph
GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Sub ListToTabel()
Dim dd As New DocData
Dim om As Range
Dim Tabel As Table
Dim i As Integer, j As Integer
On Error GoTo Fejl
PrepareMaxima
dd.ReadSelection

GoToInsertPoint
Selection.TypeParagraph
'Selection.Tables.Add Selection.Range, dd.nrows, dd.ncolumns
        Set Tabel = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=dd.nrows, NumColumns:=dd.ncolumns _
        , DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed)

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
        Tabel.Cell(i, j).Range.Text = dd.TabelsCelle(i, j)
    Next
Next

GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
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

Sub testpf()
MsgBox Environ("%programfiles%")
MsgBox Environ("programfiles")
End Sub

Sub RestartWordMat()
' genstart Maxima og genopretter doc1
RestartMaxima
LukTempDoc
'tempDoc.Close (False)
'Set tempDoc = Nothing
OpretTempdoc
End Sub
Sub TestVector()
PrepareMaxima
    omax.ReadSelection
    MsgBox Get2DVector(omax.Kommando)
End Sub
Function Get2DVector(Text As String) As String
    Dim ea As New ExpressionAnalyser
'    Dim c As Collection
    Dim M As CMatrix
    ea.Text = Text
    
'    c = ea.GetAllMatrices()
    For Each M In ea.GetAllMatrices()
        
        Get2DVector = Get2DVector & "vector(" & M
    Next
    
End Function

Sub InsertNumberedEquation(Optional AskRef As Boolean = False)
    Dim t As Table, f As Field, ccut As Boolean, i As Long
    Dim placement As Integer
    On Error GoTo Fejl
    Application.ScreenUpdating = False


    If Selection.Tables.Count > 0 Then
        MsgBox "Cant insert numbered equation in table", vbOKOnly, Sprog.Error
        Exit Sub
    End If

#If Mac Then
#Else
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
#End If

    If Selection.OMaths.Count > 0 Then
        If Not Selection.OMaths(1).Range.Text = vbNullString Then
            Selection.OMaths(1).Range.Cut
            ccut = True
            'der kan nogen gange være en rest af et matematikfelt
            If Selection.OMaths.Count > 0 Then
                If Selection.OMaths(1).Range.Text = vbNullString Then
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
      DoEvents
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
    t.Columns(1).PreferredWidthType = _
    wdPreferredWidthPercent

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
        Set f = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, Text:="LISTNUM ""WMeq"" ""NumberDefault"" \L 4")
        f.Update
        '        f.Code.Fields.ToggleShowCodes
    Else
        Selection.TypeText "("
        '        Set f = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, text:="SEQ chapter \c")
        Set f = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, Text:="SEQ WMeq1 \c")
        '        Set f = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, text:="STYLEREF ""Overskrift 1""")
        '        Set f = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, text:="SECTION")
        f.Update
        '        f.Code.Fields.ToggleShowCodes
        Selection.TypeText "."
        '        Set f = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, text:="SEQ figure \s1") ' starter automatisk forfra ved ny overskrift 1
        Set f = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, Text:="SEQ WMeq2 ")
        f.Update
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

#If Mac Then
#Else
        Oundo.EndCustomRecord
#End If

    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub InsertEquationRef()
Dim b As String
    On Error GoTo Fejl
    UserFormEquationReference.Show
    b = UserFormEquationReference.EqName
    
    If b <> vbNullString Then
#If Mac Then
#Else
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
#End If
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
    
#If Mac Then
#Else
        Oundo.EndCustomRecord
#End If
    
    End If
    
    ActiveDocument.Fields.Update
    
'    Selection.InsertCrossReference referencetype:="Bogmærke", ReferenceKind:= _
'        wdContentText, ReferenceItem:="lign1", InsertAsHyperlink:=False, _
'        IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "

'    Selection.MoveLeft Unit:=wdCharacter, count:=1
'    Selection.Fields.ToggleShowCodes
    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub SetEquationNumber()
On Error GoTo Fejl
    Application.ScreenUpdating = False
    Dim f As Field, f2 As Field, t As String, n As String, i As Integer, p As Integer, Arr As Variant
    
    If Selection.Fields.Count = 0 Then
        MsgBox Sprog.A(345), vbOKOnly, Sprog.Error
        Exit Sub
    End If
    
    Set f = Selection.Fields(1)
    If Selection.Fields.Count = 1 And InStr(f.Code.Text, "LISTNUM") > 0 Then
        n = InputBox(Sprog.A(346), Sprog.A(6), "1")
        p = InStr(f.Code.Text, "\S")
        If p > 0 Then
            f.Code.Text = Left(f.Code.Text, p - 1)
        End If
        f.Code.Text = f.Code.Text & "\S" & n
        f.Update
    ElseIf Selection.Fields.Count = 1 Or Selection.Fields.Count = 2 And InStr(f.Code.Text, "WMeq") > 0 Then
        If Selection.Fields.Count = 2 Then
            Set f2 = Selection.Fields(2)
            n = InputBox(Sprog.A(346), Sprog.A(6), f.result & "." & f2.result)
            Arr = Split(n, ".")
            If UBound(Arr) > 0 Then
                SetFieldNo f, CStr(Arr(0))
                SetFieldNo f2, CStr(Arr(1))
            Else
                SetFieldNo f, CStr(Arr(0))
            End If
        Else
            n = InputBox(Sprog.A(346), Sprog.A(6), f.result)
            SetFieldNo f, n
        End If
        
    End If
    
    ActiveDocument.Fields.Update
    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub SetFieldNo(f As Field, n As String)
    Dim p As Integer, p2 As Integer
On Error GoTo Fejl
    p = InStr(f.Code.Text, "\r")
    p2 = InStr(f.Code.Text, "\c")
    If p2 > 0 And p2 < p Then p = p2
    If p > 0 Then
        f.Code.Text = Left(f.Code.Text, p - 1)
    End If
    f.Code.Text = f.Code.Text & "\r" & n & " \c"
    f.Update
    ActiveDocument.Fields.Update
    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub InsertEquationHeadingNo()
    Dim result As Long
On Error GoTo Fejl
    result = MsgBox(Sprog.A(348), vbYesNoCancel, Sprog.A(8))
    If result = vbCancel Then Exit Sub
    If result = vbYes Then
      Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, Text:="SEQ WMeq1"
    Else
      Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, Text:="SEQ WMeq1 \h"
    End If
      Selection.Collapse wdCollapseEnd
      Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, Text:="SEQ WMeq2 \r0 \h"

    ActiveDocument.Fields.Update
    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub UpdateEquationNumbers()
On Error GoTo Fejl

    ActiveDocument.Fields.Update
    
    GoTo slut
Fejl:
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
    On Error GoTo Fejl
    Dim path As String
    Dim UFbackup As UserFormBackup
    Dim UFwait As UserFormWaitForMaxima
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
    If ActiveDocument.path = "" Then
        MsgBox Sprog.A(679)
        Exit Sub
    End If
    Set UFwait = New UserFormWaitForMaxima
    UFwait.Show vbModeless
    UFwait.Label_tip.Caption = "Saving backup" ' to " & VbCrLfMac & "documents\WordMat-Backup"
    UFwait.Label_progress.Caption = "*"
    DoEvents
   
    
'    Application.ScreenUpdating = False
    If ActiveDocument.Saved = False Then ActiveDocument.Save
    UFwait.Label_progress.Caption = UFwait.Label_progress.Caption & "*"
    DoEvents
    BackupNo = BackupNo + 1
    If BackupNo > BackupMaxNo Then BackupNo = 1
#If Mac Then
    path = GetTempDir & "WordMat-Backup/"
#Else
    path = GetDocumentsDir & "\WordMat-Backup\"
#End If
'    If Dir(path, vbDirectory) = "" Then MkDir path
    If Not FileExists(path) Then MkDir path
    UFwait.Label_progress.Caption = UFwait.Label_progress.Caption & "*"
    DoEvents
    path = path & "WordMatBackup" & BackupNo & ".docx"
    If VBA.LenB(path) = lCancelled_c Then Exit Sub
    
    Set tempDoc2 = Application.Documents.Add(Template:=ActiveDocument.FullName, visible:=False)
    UFwait.Label_progress.Caption = UFwait.Label_progress.Caption & "*"
    DoEvents
#If Mac Then
    tempDoc2.ActiveWindow.Left = 2000
    tempDoc2.SaveAs path
#Else
    tempDoc2.SaveAs2 path
#End If
    UFwait.Label_progress.Caption = UFwait.Label_progress.Caption & "*"
    DoEvents
    tempDoc2.Close

GoTo slut
Fejl:
    MsgBox Sprog.A(178), vbOKOnly, Sprog.A(208)
slut:
On Error Resume Next
    If Not UFwait Is Nothing Then Unload UFwait
    Application.ScreenUpdating = True
End Sub

Sub OpenLatexTemplate()
On Error GoTo Fejl
    Documents.Add Template:=GetWordMatDir() & "WordDocs/LatexWordTemplate.dotx"
GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub DeleteNormalDotm()
On Error Resume Next
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
    Shell "explorer.exe " & "C:\Users\" & UserDir & "\AppData\Roaming\Microsoft\Templates", vbNormalFocus
#End If
End Sub

Function ReadTextfileToString(Filnavn As String) As String
#If Mac Then
   Dim filnr As Integer
   filnr = FreeFile()
   Open Filnavn For Input As filnr   ' Open file
   ReadTextfileToString = Input$(LOF(1), 1)
   Close #filnr
   
#Else
   Dim fsT As Object
   'On Error GoTo fejl

   Set fsT = CreateObject("ADODB.Stream")
   fsT.Type = 2 'Specify stream type - we want To save text/string data.
   fsT.Charset = "iso-8859-1" 'Specify charset For the source text data. (Alternate: utf-8)
   fsT.Open 'Open the stream
   fsT.LoadFromFile Filnavn
   ReadTextfileToString = fsT.ReadText()
   fsT.Close
   Set fsT = Nothing
#End If

   GoTo slut
Fejl:
   MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error '"Der skete en fejl i forsøget på at gemme LaTex-filen"
slut:

End Function

Sub WriteTextfileToString(Filnavn As String, WriteText As String)
#If Mac Then
   Dim filnr As Integer
   filnr = FreeFile()
   Open Filnavn For Output As filnr   ' Open file for output.
   
   Print #filnr, WriteText  ' print skriver uden " "
   Close #filnr    ' Close file.
#Else
   Dim fsT As Object
   'On Error GoTo fejl

   If Filnavn = "" Then GoTo slut
   If WriteText = "" Then
      If Dir(Filnavn) <> "" Then Kill Filnavn
         GoTo slut
   End If
   Set fsT = CreateObject("ADODB.Stream")
   fsT.Type = 2 'Specify stream type - we want To save text/string data.
   fsT.Charset = "iso-8859-1" 'Specify charset For the source text data. utf-8
   fsT.Open 'Open the stream And write binary data To the object
   fsT.WriteText WriteText
   fsT.SaveToFile Filnavn, 2 'Save binary data To disk
   fsT.Close
   Set fsT = Nothing
#End If


   GoTo slut
Fejl:
   MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error '"Der skete en fejl i forsøget på at gemme LaTexfilen"
slut:

End Sub

Public Function Local_Document_Path(ByRef Doc As Document, Optional bPathOnly As Boolean = True) As String
'returns local path or nothing if local path not found. Converts a onedrive path to local path
#If Mac Then
   Local_Document_Path = Doc.path
#Else
Dim i As Long, x As Long
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
        For x = 1 To 4
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
   Dim p As Long, p2 As Long, p3 As Long
   
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
    On Error GoTo Fejl
    If Selection.OMaths.Count = 0 Then
        Set r = Selection.OMaths.Add(Selection.Range)
    ElseIf Selection.Tables.Count = 0 Then
        If Selection.OMaths(1).Range.Text = vbNullString Then
            Set r = Selection.OMaths.Add(Selection.Range)
        Else
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
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
