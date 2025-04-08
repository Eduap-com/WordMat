Attribute VB_Name = "Diverse"
Option Explicit
Public TimeText As String
Public cxl As CExcel
Public ProgramFilesDir
Public DocumentsDir
Private UserDir As String
Private tmpdir As String
#If Mac Then
#Else
Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
#End If

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

Function fileExists(FullFileName As String) As Boolean
' returns TRUE if the file or folder exists
    On Error GoTo Err
    fileExists = False
    fileExists = Len(Dir(FullFileName)) > 0 Or Len(Dir(FullFileName, vbDirectory)) > 0
    Exit Function
Err:
End Function

Function GetTempDir() As String
#If Mac Then
    If UserDir = vbNullString Then UserDir = MacScript("return POSIX path of (path to home folder)")
    tmpdir = UserDir & "WordMat/"
    On Error Resume Next
    MkDir (tmpdir)
    GetTempDir = tmpdir
#Else
    GetTempDir = Environ("TEMP") & "\"
#End If
End Function

Sub ChangeAutoHyphen()
    Options.AutoFormatAsYouTypeReplaceFarEastDashes = False
    Options.AutoFormatAsYouTypeReplaceSymbols = False
End Sub

Sub ShowCustomizationContext()
'    MsgBox CustomizationContext & vbCrLf & ActiveDocument.AttachedTemplate
    MsgBox Templates(4)
End Sub
Function GetWordMatTemplate(Optional NormalDotmOK As Boolean = False) As Template
' If the current document is called wordmat*.dotm then it is returned as a template
' Otherwise all global templates are searched to see if there is one called wordmat*.dotm
    If Len(ActiveDocument.AttachedTemplate) > 10 Then
        If LCase(Left(ActiveDocument.AttachedTemplate, 7)) = "wordmat" And LCase(right(ActiveDocument.AttachedTemplate, 5)) = ".dotm" Then
            Set GetWordMatTemplate = ActiveDocument.AttachedTemplate
            Exit Function
        End If
    End If
    If NormalDotmOK Then
        Set GetWordMatTemplate = NormalTemplate
    End If

' It is not possible to modify wordmat.dotm if the file is not opened directly. It cannot be saved.
'    For Each WT In Application.Templates
'        If LCase(Left(WT, 7)) = "wordmat" And LCase(right(WT, 5)) = ".dotm" Then
'            Set GetWordMatTemplate = WT
'            Exit Function
'        End If
'    Next
End Function

Function GetProgramFilesDir() As String
' is not used by maxima anymore as the dll file is responsible for it now.
' is used by the Word documents etc. that need to be found
'MsgBox GetProgFilesPath
    On Error GoTo Fejl
#If Mac Then
    GetProgramFilesDir = "/Applications/"
#Else
    If ProgramFilesDir <> "" Then
        GetProgramFilesDir = ProgramFilesDir
    Else
        GetProgramFilesDir = RegKeyRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\ProgramFilesDir")
        If Dir(GetProgramFilesDir & "\WordMat", vbDirectory) = "" Then
            GetProgramFilesDir = RegKeyRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\ProgramW6432Dir")
        End If
        If Dir(GetProgramFilesDir & "\WordMat", vbDirectory) = "" Then
            GetProgramFilesDir = Environ("ProgramFiles")
        End If
        If Dir(GetProgramFilesDir & "\WordMat", vbDirectory) = "" Then
            GetProgramFilesDir = RegKeyRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\ProgramFilesDir (x86)")
        End If
        ProgramFilesDir = GetProgramFilesDir
    End If
#End If

    GoTo slut
Fejl:
    MsgBox Sprog.A(110), vbOKOnly, Sprog.Error
slut:
End Function

Function GetDocumentsDir() As String
    On Error GoTo Fejl
    If DocumentsDir <> "" Then
        GetDocumentsDir = DocumentsDir
    Else
#If Mac Then
        Dim p As Integer
        GetDocumentsDir = MacScript("return POSIX path of (path to documents folder) as string")
        p = InStr(GetDocumentsDir, "/Library")
        GetDocumentsDir = Left(GetDocumentsDir, p) & "Documents"
#Else
        GetDocumentsDir = RegKeyRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal")
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
End Function

Function GetDownloadsFolder() As String
#If Mac Then
    GetDownloadsFolder = RunScript("GetDownloadsFolder", vbNullString)
#Else
    GetDownloadsFolder = RegKeyRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\{374DE290-123F-4565-9164-39C4925E467B}")
    GetDownloadsFolder = Replace(GetDownloadsFolder, "%USERPROFILE%", Environ$("USERPROFILE"))
#End If
End Function
Function RegKeyRead(i_RegKey As String) As Variant
'example syntax: "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\ProgramFilesDir"
#If Mac Then
    RegKeyRead = GetSetting("com.wordmat", "defaults", i_RegKey)
#Else
    On Error Resume Next
    If MaxProc Is Nothing Then
        Dim myWS As Object
        Set myWS = CreateObject("WScript.Shell")
        RegKeyRead = myWS.RegRead(i_RegKey)
    Else
        RegKeyRead = MaxProc.RegKeyRead(i_RegKey)
    End If
#End If
End Function

Function RegKeyExists(i_RegKey As String) As Boolean
#If Mac Then
    RegKeyExists = True
    If GetSetting("com.wordmat", "defaults", i_RegKey) = "" Then RegKeyExists = False
#Else
    If MaxProc Is Nothing Then
        Dim myWS As Object
        On Error GoTo ErrorHandler
        Set myWS = CreateObject("WScript.Shell")
        myWS.RegRead i_RegKey
        RegKeyExists = True
        Exit Function
ErrorHandler:  'key was not found
        RegKeyExists = False
    Else
        RegKeyExists = MaxProc.RegKeyExists(i_RegKey)
    End If
#End If
End Function

Sub RegKeySave(ByVal i_RegKey As String, ByVal i_Value As String, Optional ByVal i_Type As String = "REG_SZ")
    '
#If Mac Then
    SaveSetting "com.wordmat", "defaults", i_RegKey, i_Value
#Else
    If MaxProc Is Nothing Then
        Dim myWS As Object
        On Error Resume Next
        Set myWS = CreateObject("WScript.Shell")
        myWS.RegWrite i_RegKey, i_Value, i_Type
    Else
        MaxProc.RegKeySave i_RegKey, i_Value 'i_value can be string or integer. can be saved to REG_SZ or REG_DWORD. If key does not exist REG_SZ type can be created, not DWORD
    End If
#End If
End Sub

Function RegKeyDelete(i_RegKey As String) As Boolean
#If Mac Then
    DeleteSetting "com.wordmat", "defaults", i_RegKey
#Else
    If MaxProc Is Nothing Then
        Dim myWS As Object
        On Error GoTo ErrorHandler
        Set myWS = CreateObject("WScript.Shell")
        On Error Resume Next
        myWS.RegDelete i_RegKey
        RegKeyDelete = True
        Exit Function
ErrorHandler:
        RegKeyDelete = False
    Else
        MaxProc.RegKeyDelete i_RegKey
    End If
#End If
End Function
Sub OpenLink(Link As String, Optional Script As Boolean = False)
' note: Script is always true on mac to prevent warning
On Error Resume Next

#If Mac Then
    Script = True
    If Script Then
        RunScript "OpenLink", Link
    Else
        ActiveDocument.FollowHyperlink Address:=Link, NewWindow:=True
    End If
#Else
' ActiveDocument.FollowHyperlink removes parameters such as ?command=... Therefore it may be necessary to use script
    If Script Then
        MaxProc.RunFile GetProgramFilesDir & "\Microsoft\Edge\Application\msedge.exe", """" & Link & """"
    Else
        ActiveDocument.FollowHyperlink Address:=Link, NewWindow:=True ' If the link doesn't work, nothing happens.
    End If
#End If
Fejl:
End Sub

Sub InsertSletDef()
    Dim gemfontsize As Integer
    Dim gemitalic As Boolean
    Dim gemfontcolor As Integer
    Dim gemsb As Integer
    Dim gemsa As Integer
    Dim Oundo As UndoRecord
    On Error GoTo slut
    
    If Selection.Tables.Count > 0 Then
        MsgBox2 Sprog.A(475), vbOKOnly, Sprog.Error
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord
    
    gemfontsize = Selection.Font.Size
    gemitalic = Selection.Font.Italic
    gemfontcolor = Selection.Font.ColorIndex
    gemsb = Selection.ParagraphFormat.SpaceBefore
    gemsa = Selection.ParagraphFormat.SpaceAfter
            
    If Selection.OMaths.Count > 0 Then
        Selection.OMaths(1).Range.Select
        Selection.Collapse wdCollapseEnd
        Selection.MoveRight wdCharacter, 1
        Selection.TypeParagraph
    Else
        Selection.Paragraphs(1).Range.Select
        Selection.Collapse wdCollapseEnd
        If Selection.OMaths.Count > 0 Then
            Selection.MoveLeft wdCharacter, 1
            If Selection.OMaths.Count > 0 Then
                Selection.MoveLeft wdCharacter, 1
            End If
        End If
    End If

    Selection.OMaths.Add Range:=Selection.Range
    DoEvents
    On Error Resume Next
    Selection.OMaths(1).Range.Font.Size = 8
    Selection.OMaths(1).Range.Font.ColorIndex = wdGray50
    On Error GoTo slut
    Selection.TypeText Sprog.A(69) & ":"
    Selection.Collapse (wdCollapseEnd)
    Selection.TypeParagraph
    Selection.Font.Bold = False
        
    If Selection.OMaths.Count = 0 Then
        Selection.Font.Size = gemfontsize
        Selection.Font.Italic = gemitalic
        Selection.Font.ColorIndex = gemfontcolor
        With Selection.ParagraphFormat
            .SpaceBefore = gemsb
            '        .SpaceBeforeAuto = False
            .SpaceAfter = gemsa
            '        .SpaceAfterAuto = False
        End With
    End If
slut:
    Oundo.EndCustomRecord
End Sub

Sub InsertDefiner()
    On Error GoTo Fejl

    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord

    Application.ScreenUpdating = False
    If Selection.OMaths.Count > 0 Then
        If Selection.OMaths(1).Type = wdOMathInline Then
            Selection.OMaths(1).Range.Select
            If Selection.OMaths(1).Range.Text = "Type equation here." Or Selection.OMaths(1).Range.Text = "Skriv ligningen her." Then
            Else
                Selection.Collapse wdCollapseStart
                Selection.MoveRight wdCharacter, 1
            End If
            Selection.TypeText Sprog.A(62) & ": "
        Else
            Selection.OMaths(1).Range.Select
            Selection.Collapse wdCollapseStart
            If Selection.OMaths(1).Range.Text = "Type equation here." Or Selection.OMaths(1).Range.Text = "Skriv ligningen her." Then
                Selection.MoveRight wdCharacter, 1
            End If
            Selection.TypeText Sprog.A(62) & ": "
        End If
    Else
        Selection.OMaths.Add Selection.Range
        Selection.TypeText Sprog.A(62) & ": "
    End If
    Selection.Collapse wdCollapseEnd
        
    GoTo slut
Fejl:
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
    If ResPos1 = Selection.Range.start Then ' if repeated click
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
        s = omax.ReadEquation2(r)
'        s = omax.ReadEquation(r)
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
        ResPos1 = ResPos1 - 1 ' if already empty, selection is for some reason 1 character too many
    End If
    s = Replace(s, VBA.ChrW(8289), "") ' function sign sin(x) otherwise becomes si*n(x). also problem with other functions
    Selection.Text = s
    
GoTo slut
Fejl:
    ResIndex = 0
    ResFeltIndex = 0
    ResPos2 = 0
    ResPos1 = 0
slut:
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub

Function KlipTilLigmed(Text As String, ByVal indeks As Integer) As String
' returns the last part of the text to the first position counted from the end for = or approximately equal
' = in the sum sign is ignored

    Dim posligmed As Integer
    Dim possumtegn As Integer
    Dim posca As Integer
    Dim poseller As Integer
    Dim Pos As Integer
    Dim Arr(20) As String
    Dim i As Integer
    
    Do ' go back to nearest equal sign
        posligmed = InStr(Text, "=")
        possumtegn = InStr(Text, VBA.ChrW(8721))
        posca = InStr(Text, VBA.ChrW(8776))
        poseller = InStr(Text, VBA.ChrW(8744))
        
        Pos = Len(Text)
    '    pos = posligmed
        If posligmed > 0 And posligmed < Pos Then Pos = posligmed
        If posca > 0 And posca < Pos Then Pos = posca
        If poseller > 0 And poseller < Pos Then Pos = poseller
        
        If possumtegn > 0 And possumtegn < Pos Then ' if there is a sum sign, there is an = sign as part of it
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
    
    If indeks = i Then ResIndex = -1  ' global variable marks that there are no more to the left
    If i = 0 Then
        KlipTilLigmed = Text
        ResIndex = -1
    Else
        KlipTilLigmed = Arr(i - indeks)
    End If
    
    ' remove returns and spaces etc.
'    s = Replace(s, vbCrLf, "")
    KlipTilLigmed = Replace(KlipTilLigmed, vbCr, "")
    KlipTilLigmed = Replace(KlipTilLigmed, VBA.ChrW(11), "")
'    s = Replace(s, vbLf, "")
    KlipTilLigmed = Replace(KlipTilLigmed, VBA.ChrW(8744), "") 'eller tegn
'    KlipTilLigmed = Replace(KlipTilLigmed, " ", "")
    KlipTilLigmed = Trim(KlipTilLigmed)
    
    If InStr(KlipTilLigmed, "/") > 0 Then KlipTilLigmed = "  " & KlipTilLigmed
    
End Function

Sub OpenFormulae(FilNavn As String)
On Error GoTo Fejl
#If Mac Then
    Documents.Open "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/WordDocs/" & FilNavn
#Else
    OpenWordFile "" & FilNavn
#End If
GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Sub OpenWordFile(FilNavn As String)
'Example: OpenWordFile ("Figurer.docx")

    Dim filnavn1 As String
#If Mac Then
    FilNavn = Replace(FilNavn, "\", "/")
    filnavn1 = GetWordMatDir() & "WordDocs/" & FilNavn
    Documents.Open filnavn1
#Else
    Dim filnavn2 As String
    Dim appdir As String
    On Error GoTo Fejl
    appdir = Environ("AppData")
    filnavn1 = appdir & "\WordMat\WordDocs\" & FilNavn

    If Dir(filnavn1) = vbNullString Then
        filnavn2 = GetProgramFilesDir & "\WordMat\WordDocs\" & FilNavn
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
Fejl:
    MsgBox Sprog.A(111) & FilNavn, vbOKOnly, Sprog.Error
slut:
End Sub

Function GetRandomTip()
    Dim i As Integer
    Dim n As Integer
    Dim mindste As Integer
    Dim tip As String
    n = 29 ' no. of tips
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
    i = Int(Rnd(1) * (n - mindste) + mindste) ' random number n 0-(n-1)

    Select Case i
    Case 0
        tip = Sprog.A(325)
    Case 1
        tip = Sprog.A(326)
    Case 2
        tip = Sprog.A(327)
    Case 3
        tip = Sprog.A(328)
    Case 4
        tip = Sprog.A(329)
    Case 5
        tip = Sprog.A(330)
    Case 6
        tip = Sprog.A(331) & "   x_1   ->   x" & VBA.ChrW(8321)
    Case 7
        tip = Sprog.A(332) & VBA.ChrW(955)
    Case 8
        tip = Sprog.A(333)
    Case 9
        tip = Sprog.A(334)
    Case 10
        tip = Sprog.A(335)
    Case 11
        tip = Sprog.A(336)
    Case 12
        tip = "(a+b)" & VBA.ChrW(178) & " = a" & VBA.ChrW(178) & " + b" & VBA.ChrW(178) & " + 2ab"
    Case 13
        tip = "(a+b)(a-b) = a" & VBA.ChrW(178) & " - b" & VBA.ChrW(178)
    Case 14
        tip = "(a-b)" & VBA.ChrW(178) & " = a" & VBA.ChrW(178) & " + b" & VBA.ChrW(178) & " - 2ab"
    Case 15
        tip = "(a" & VBA.ChrW(183) & "b)" & VBA.ChrW(7510) & " = a" & VBA.ChrW(7510) & VBA.ChrW(183) & "b" & VBA.ChrW(7510)
    Case 16
        tip = Sprog.A(337) & AntalB
    Case 17
        tip = "log(a" & VBA.ChrW(7495) & ") = b" & VBA.ChrW(183) & "log(a)"
    Case 18
        tip = "log(a/b) = log(a) - log(b)"
    Case 19
        tip = "log(a" & VBA.ChrW(183) & "b) = log(a) + log(b)"
    Case 20
        tip = "\int      ->      " & VBA.ChrW(8747)
    Case 21
        tip = Sprog.A(338)
    Case 22
        tip = Sprog.A(339)
    Case 23
        tip = Sprog.A(340)
    Case 24
        tip = "(a/b)" & VBA.ChrW(7510) & " = a" & VBA.ChrW(7510) & "/b" & VBA.ChrW(7510)
    Case 25
        tip = "a/b + c/d = (ad+bc)/bd"
    Case 26
        tip = Sprog.A(341)
    Case 27
        tip = Sprog.A(342)
    Case 28
        tip = "\inc  ->  delta"
    Case 29
        tip = ""
    Case 30
        tip = ""
    Case Else
        tip = ""
    End Select
        
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
        PrepareMaxima False
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
    End If
    RefreshRibbon
End Sub

Sub TurnUnitsOn()
    On Error Resume Next
    MaximaUnits = True
    Application.OMathAutoCorrect.Functions("min").Delete  ' otherwise mine cannot be used as a unit
End Sub
Sub TurnUnitsOff()
    On Error Resume Next
    MaximaUnits = False
    Application.OMathAutoCorrect.Functions.Add "min"
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

Sub CheckForUpdate()
    CheckForUpdatePar False
End Sub
Sub CheckForUpdatePar(Optional RunSilent As Boolean = False)
    On Error GoTo Fejl
    Dim NewVersion As String, p As Integer, News As String, s As String
    Dim UpdateNow As Boolean, PartnerShip As Boolean
   
    If GetInternetConnectedState = False Then
        If Not RunSilent Then MsgBox Sprog.A(63), vbOKOnly, Sprog.Error
        Exit Sub
    End If
    
    If RunSilent Then
        If (Month(Date) = 5 Or Month(Date) = 6) Then GoTo slut ' do not automatically update in May and June
        If IsDate(LastUpdateCheck) Then
            If DateDiff("d", LastUpdateCheck, Date) < 7 Then GoTo slut ' if checked within the last 7 days then exit
        End If
    End If
    LastUpdateCheck = Date ' this should be here, and not at the end, because if an error occurs in the update, it should only come once
    
    On Error Resume Next
    PartnerShip = QActivePartnership()
    On Error GoTo Fejl
    
#If Mac Then
    If PartnerShip Then
        s = RunScript("GetHTML", "https://www.eduap.com/download/info/wordmatversionP.txt")
    Else
        s = RunScript("GetHTML", "https://www.eduap.com/download/info/wordmatversion.txt")
    End If
    If InStr(s, "404 Not Found") > 0 Then s = vbNullString
#Else
    If PartnerShip Then
        s = GetHTML("https://www.eduap.com/download/info/wordmatversionP.txt")
    Else
        s = GetHTML("https://www.eduap.com/download/info/wordmatversionP.txt")
    End If
#End If
    If Len(s) = 0 Then
        If Not RunSilent Then
            MsgBox2 Sprog.A(112), vbOKOnly, Sprog.Error
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
            MsgBox2 Sprog.A(112), vbOKOnly, Sprog.Error
        End If
        GoTo slut
    End If
   
    If IsNumeric(AppVersion) And IsNumeric(NewVersion) Then
        If val(AppVersion) < val(NewVersion) Then UpdateNow = True
    Else
        If AppVersion <> NewVersion Then UpdateNow = True
    End If
    
    If UpdateNow Then
        If PartnerShip Then
            If MsgBox2(Sprog.A(21) & News & vbCrLf & vbCrLf & Sprog.A(64), vbOKCancel, Sprog.A(23)) = vbOK Then
                On Error Resume Next
                Documents.Save NoPrompt:=True, OriginalFormat:=wdOriginalDocumentFormat
                On Error GoTo Install2
                Application.Run macroname:="PUpdateWordMat"
                On Error GoTo Fejl
            End If
        Else
Install2:
            On Error GoTo Fejl
            MsgBox2 Sprog.A(21) & News & vbCrLf & Sprog.A(22) & vbCrLf & vbCrLf & "", vbOKOnly, Sprog.A(23)
            If Sprog.SprogNr = 1 Then
                OpenLink "https://www.eduap.com/da/wordmat/"
            Else
                OpenLink "https://www.eduap.com/wordmat/"
            End If
        End If
    Else
        If Not RunSilent Then
            MsgBox2 Sprog.A(344) & " " & AppNavn & " v." & AppVersion, vbOKOnly, "No Update"
        End If
    End If
   
    GoTo slut
Fejl:
    If Not RunSilent Then
        If MsgBox2(Sprog.A(581) & AppVersion, vbOKCancel, Sprog.Error) = vbOK Then
            If Sprog.SprogNr = 1 Then
                OpenLink "https://www.eduap.com/da/wordmat/"
            Else
                OpenLink "https://www.eduap.com/wordmat/"
            End If
        End If
    End If
slut:

End Sub

Sub CheckForUpdateSilent()
    CheckForUpdatePar True
End Sub
Function GetHTML(URL As String) As String
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", URL & "?cb=" & timer() * 100, False  ' timer ensures that it is not a cached version
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
    If MaximaDecOutType = 3 Or Abs(n) > 10 ^ 6 Or Abs(n) < 10 ^ -6 Then
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
' convert number to string with specified number of significant digits. If none is specified, maximum digits are used
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
' takes E into account, but not entirely uniquely.

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
' inserts landscape page and regular page after
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
' unfortunately causes rare problems on some computers
' especially if there are definitions in the document so it is fired twice
On Error GoTo slut
    Dim oData   As New DataObject 'object to use the clipboard
     
    oData.SetText Text:=Empty 'Clear
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
        mc(mc.Count).Range.Select  ' works with word 2010, parentomath gives problems though. Hmm problem with selected part of expression and reducer
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
    If r.Text = VBA.ChrW(11) Then ' if there is shift-enter at the end, replace with regular return
        r.Text = VBA.ChrW(13)
    End If
End Sub

Function NotZero(i As Integer) As Integer
' if negative return zero
    If i < 0 Then
        NotZero = 0
    Else
        NotZero = i
    End If
End Function

Sub TabelToList()
    Dim dd As New DocData
    Dim OM As Range
    On Error GoTo Fejl
    If Selection.Range.Tables.Count = 0 Then
        MsgBox Sprog.A(871), vbOKOnly, Sprog.Error
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
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub
Sub ListToTabel()
    Dim dd As New DocData
    Dim Tabel As Table
    Dim i As Integer, j As Integer
    On Error GoTo Fejl
    
    PrepareMaxima
    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord
    dd.ReadSelection
    If dd.nrows = 0 Or dd.ncolumns = 0 Then
        MsgBox Sprog.A(901), vbOKOnly, Sprog.Error
        GoTo slut
    End If
    GoToInsertPoint
    Selection.TypeParagraph
    Set Tabel = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=dd.nrows, NumColumns:=dd.ncolumns, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed)

    With Selection.Tables(1)
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

    For i = 1 To dd.nrows
        For j = 1 To dd.ncolumns
            Tabel.Cell(i, j).Range.Text = dd.TabelsCelle(i, j)
        Next
    Next

    GoTo slut
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
    Oundo.EndCustomRecord
End Sub
Sub GoToInsertPoint()
' finds the next point after selection where more can be inserted
' i.e. goes after math boxes and tables

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

Sub GenerateAutoCorrect()
' generates math autocorrect. Not used, but has potential
    Application.OMathAutoCorrect.UseOutsideOMath = True
    
    Application.OMathAutoCorrect.Entries.Add "\bi", VBA.ChrW(8660) ' biimplicative arrows
    Application.OMathAutoCorrect.Entries.Add "\imp", VBA.ChrW(8658) ' implikationarrow right
End Sub

Sub RestartWordMat()
    RestartMaxima
End Sub

Sub InsertNumberedEquation(Optional AskRef As Boolean = False)
    Dim t As Table, F As Field, ccut As Boolean
    Dim placement As Integer
    On Error GoTo Fejl
    
    If AskRef Then
        Dim EqName As String
        UserFormEnterEquationRef.Show
        EqName = UserFormEnterEquationRef.EquationName    'Replace(InputBox(Sprog.A(5), Sprog.A(4), "Eq"), " ", "")
        If EqName = vbNullString Then GoTo slut
    End If
    
    Application.ScreenUpdating = False

    If Selection.Tables.Count > 0 Then
        MsgBox2 Sprog.A(872), vbOKOnly, Sprog.Error
        Exit Sub
    End If

    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord

    If Selection.OMaths.Count > 0 Then
        If Not Selection.OMaths(1).Range.Text = vbNullString Then
            Selection.OMaths(1).Range.Cut
            ccut = True
            'there may sometimes be a remainder of a mathematical field
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

    'Insert number
    If EqNumPlacement Then
        placement = 1
    Else
        placement = 3
    End If
    t.Cell(1, placement).Range.Select
    Selection.Collapse wdCollapseStart
    If Not EqNumType Then
        Set F = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, Text:="LISTNUM ""WMeq"" ""NumberDefault"" \L 4")
        F.Update
        '        f.Code.Fields.ToggleShowCodes
    Else
        Selection.TypeText "("
        '        Set f = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, text:="SEQ chapter \c")
        Set F = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, Text:="SEQ WMeq1 \c")
        '        Set f = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, text:="STYLEREF ""Overskrift 1""")
        '        Set f = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, text:="SECTION")
        F.Update
        '        f.Code.Fields.ToggleShowCodes
        Selection.TypeText "."
        '        Set f = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, text:="SEQ figure \s1") ' starter automatisk forfra ved ny overskrift 1
        Set F = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, Text:="SEQ WMeq2 ")
        F.Update
        '        f.Code.Fields.ToggleShowCodes
        Selection.TypeText ")"
    End If

    If AskRef Then
        If EqName <> vbNullString Then
            t.Cell(1, 3).Range.Fields(1).Select
            With ActiveDocument.Bookmarks
                .Add Range:=Selection.Range, Name:=EqName
                .DefaultSorting = wdSortByName
                .ShowHidden = False
            End With
        End If
    End If

    ' insert math field
    t.Cell(1, 2).Range.Select
    Selection.Collapse wdCollapseStart
    If ccut Then
        DoEvents
        Selection.Paste
        Selection.MoveLeft unit:=wdCharacter, Count:=1
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
Fejl:
    MsgBox2 Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub InsertEquationRef()
    Dim b As String
    On Error GoTo Fejl
    UserFormEquationReference.Show
    b = UserFormEquationReference.EqName
    
    If b <> vbNullString Then
        Dim Oundo As UndoRecord
        Set Oundo = Application.UndoRecord
        Oundo.StartCustomRecord
        Selection.TypeText Sprog.A(833) & " "
#If Mac Then
        Selection.InsertCrossReference referencetype:=wdRefTypeBookmark, ReferenceKind:= _
            wdContentText, ReferenceItem:=b, InsertAsHyperlink:=False, _
            IncludePosition:=False
#Else
        Selection.InsertCrossReference referencetype:=wdRefTypeBookmark, ReferenceKind:= _
            wdContentText, ReferenceItem:=b, InsertAsHyperlink:=False, _
            IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
#End If

        Selection.MoveLeft unit:=wdCharacter, Count:=1
        Selection.Fields.ToggleShowCodes
        Selection.Collapse wdCollapseEnd
    
        Oundo.EndCustomRecord
    End If
    
    ActiveDocument.Fields.Update
    
    GoTo slut
Fejl:
    MsgBox2 Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub SetEquationNumber()
On Error GoTo Fejl
    Application.ScreenUpdating = False
    Dim F As Field, f2 As Field, n As String, p As Integer, Arr As Variant
    
    If Selection.Fields.Count = 0 Then
        MsgBox Sprog.A(345), vbOKOnly, Sprog.Error
        Exit Sub
    End If
    
    Set F = Selection.Fields(1)
    If Selection.Fields.Count = 1 And InStr(F.Code.Text, "LISTNUM") > 0 Then
        n = InputBox(Sprog.A(346), Sprog.A(6), "1")
        p = InStr(F.Code.Text, "\S")
        If p > 0 Then
            F.Code.Text = Left(F.Code.Text, p - 1)
        End If
        F.Code.Text = F.Code.Text & "\S" & n
        F.Update
    ElseIf Selection.Fields.Count = 1 Or Selection.Fields.Count = 2 And InStr(F.Code.Text, "WMeq") > 0 Then
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
Fejl:
    MsgBox Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Sub SetFieldNo(F As Field, n As String)
    Dim p As Integer, p2 As Integer
    On Error GoTo Fejl
    p = InStr(F.Code.Text, "\r")
    p2 = InStr(F.Code.Text, "\c")
    If p2 > 0 And p2 < p Then p = p2
    If p > 0 Then
        F.Code.Text = Left(F.Code.Text, p - 1)
    End If
    F.Code.Text = F.Code.Text & "\r" & n & " \c"
    F.Update
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

Sub OpenLatexTemplate()
On Error GoTo Fejl
    Documents.Add Template:=GetWordMatDir() & "WordDocs/LatexWordTemplate.dotx"
GoTo slut
Fejl:
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
#End If
End Sub
Sub DeleteKeyboardShortcutsInNormalDotm()
' Deletes shortcuts to WordMat macros that were accidentally saved in normal.dotm
    Dim GemT As Template
    Dim KB As KeyBinding
    On Error Resume Next
    
    Set GemT = CustomizationContext
            
    CustomizationContext = NormalTemplate

#If Mac Then ' For one student, KB.command was completely empty, so on Mac, kb.keystring is used, although it is a little more insecure.
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
' you cannot save WordMat.dotm as a global template when it is saved for all users
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
' cant be run from automacros
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
' ensures that string has setting with separators
If DecSeparator = "," Then
'    n = Replace(n, ",", ";")
    ConvertNumber = Replace(n, ".", ",")
Else
    ConvertNumber = Replace(n, ",", ".")
'    n = Replace(n, ";", ",")
End If

End Function

Function GetWordMatDir(Optional SubDir As String) As String
' if a subdir to WordMat folder is stated. That particular folder will be looked for
' GetWordMatDir(
#If Mac Then
    GetWordMatDir = "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/"
#Else
    If SubDir <> vbNullString Then
        SubDir = Trim(SubDir)
        If right(SubDir, 1) <> "\" Then SubDir = SubDir & "\"
    End If
    If InstallLocation = "All" Then
        GetWordMatDir = GetProgramFilesDir() & "\WordMat\"
        If Dir(GetWordMatDir & SubDir, vbDirectory) = vbNullString Then
            GetWordMatDir = Environ("AppData") & "\WordMat\"
            If Dir(GetWordMatDir & SubDir, vbDirectory) = vbNullString Then
                MsgBox "WordMat folder could not be found", vbOKOnly, Sprog.Error
            End If
        End If
    Else
        GetWordMatDir = Environ("AppData") & "\WordMat\"
        If Dir(GetWordMatDir & SubDir, vbDirectory) = vbNullString Then
            GetWordMatDir = GetProgramFilesDir() & "\WordMat\"
            If Dir(GetWordMatDir & SubDir, vbDirectory) = vbNullString Then
                MsgBox "WordMat folder could not be found", vbOKOnly, Sprog.Error
            End If
        End If
    End If
#End If
End Function

Sub NewEquation()
    Dim r As Range
    On Error GoTo Fejl
    On Error Resume Next
    
    If Selection.OMaths.Count = 0 Then
        Set r = Selection.OMaths.Add(Selection.Range)
    ElseIf Selection.Tables.Count = 0 Then
        If Selection.OMaths(1).Range.Text = vbNullString Then
            Set r = Selection.OMaths.Add(Selection.Range)
        End If
    ElseIf Selection.Tables(1).Columns.Count = 3 And Selection.Tables(1).Cell(1, 3).Range.Fields.Count > 0 Then
        Selection.Tables(1).Cell(1, 2).Range.OMaths(1).Range.Cut
        Selection.Tables(1).Select
'        Selection.MoveEnd unit:=wdCharacter, count:=2
        Selection.Tables(1).Delete
        Selection.Paste
        Selection.TypeParagraph
        Selection.MoveLeft unit:=wdCharacter, Count:=2
    End If
GoTo slut
Fejl:
    MsgBox2 Sprog.ErrorGeneral, vbOKOnly, Sprog.Error
slut:
End Sub

Function FormatDefinitions(DefS As String) As String
' Takes a string from omax.defintions and makes it as pretty as possible for showing in a textbox
' used for showing present definitions on several forms
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
    DefS = Replace(DefS, "ae", ChrW(230))
    DefS = Replace(DefS, "oe", ChrW(248))
    DefS = Replace(DefS, "aa", ChrW(229))
    DefS = Replace(DefS, "AE", ChrW(198))
    DefS = Replace(DefS, "OE", ChrW(216))
    DefS = Replace(DefS, "AA", ChrW(197))
        
    'greek letters
    DefS = Replace(DefS, "gamma", VBA.ChrW(915))    ' big gamma
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
    DefS = Replace(DefS, "omega", VBA.ChrW(969))    ' small omega
    
    DefS = Replace(DefS, "((x))", "(x)")
        
    If DecSeparator = "," Then
        DefS = Replace(DefS, ".", ",")
    End If
        
    FormatDefinitions = DefS
End Function

Function MsgBox2(prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKCancel, Optional Title As String) As VbMsgBoxResult
' Replacement for msgbox. This matches the UI of the other userforms. It can adapt in size.
' Buttons supported: vbYesNo, vbOKonly, vbOKCancel
' Example: MsgBox2 "This is a test", vbOKOnly, "Hello"

    Dim UFMsgBox As New UserFormMsgBox
    
    UFMsgBox.MsgBoxStyle = Buttons
    UFMsgBox.Title = Title
    UFMsgBox.prompt = prompt
    
    UFMsgBox.Show
    
    MsgBox2 = UFMsgBox.MsgBoxResult
    
    Unload UFMsgBox
End Function

Sub TestMe()
 MsgBox2 "This a small test", vbOKOnly, "Hello"
 MsgBox2 "This is a longer test" & vbCrLf & "More and longer lines" & vbCrLf & "Hello" & vbCrLf & "Hello" & vbCrLf & "Hello" & vbCrLf & "Hello", vbOKCancel, "Hello"
 MsgBox2 "This is a wide test" & vbCrLf & "Has to have a long line with different characters and some different stuff" & vbCrLf & "hello" & vbCrLf & "hello" & vbCrLf & "hello" & vbCrLf & "hello", vbOKCancel, "Hello"
End Sub

Sub TestError()
    On Error Resume Next
    Err.Raise 1, , "test"
End Sub

Sub TestSprog()
    Dim tid As Double, n As Integer
    tid = timer
    n = Sprog.SprogNr
    MsgBox timer - tid
End Sub
