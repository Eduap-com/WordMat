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
        If LCase$(Left$(ActiveDocument.AttachedTemplate, 7)) = "wordmat" And LCase$(Right$(ActiveDocument.AttachedTemplate, 5)) = ".dotm" Then
            Set GetWordMatTemplate = ActiveDocument.AttachedTemplate
            Exit Function
        End If
    End If
    If NormalDotmOK Then
        Set GetWordMatTemplate = NormalTemplate
    End If

' It is not possible to modify wordmat.dotm if the file is not opened directly. It cannot be saved.
'    For Each WT In Application.Templates
'        If lcase$(left$(WT, 7)) = "wordmat" And lcase$(right$(WT, 5)) = ".dotm" Then
'            Set GetWordMatTemplate = WT
'            Exit Function
'        End If
'    Next
End Function

Function GetProgramFilesDir() As String
' is not used by maxima anymore as the dll file is responsible for it now.
' is used by the Word documents etc. that need to be found
'MsgBox GetProgFilesPath
    On Error GoTo fejl
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
fejl:
    MsgBox TT.A(110), vbOKOnly, TT.Error
slut:
End Function

Function GetDocumentsDir() As String
    On Error GoTo fejl
    If DocumentsDir <> "" Then
        GetDocumentsDir = DocumentsDir
    Else
#If Mac Then
        Dim p As Integer
        GetDocumentsDir = MacScript("return POSIX path of (path to documents folder) as string")
        p = InStr(GetDocumentsDir, "/Library")
        GetDocumentsDir = Left$(GetDocumentsDir, p) & "Documents"
#Else
        GetDocumentsDir = RegKeyRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal")
        If Dir(GetDocumentsDir, vbDirectory) = "" Then
            GetDocumentsDir = "c:\"
        End If
#End If
        DocumentsDir = GetDocumentsDir
    End If
 
    GoTo slut
fejl:
    MsgBox TT.A(110), vbOKOnly, TT.Error
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
        If Dir("C:\Program Files\Google\Chrome\Application\chrome.exe") <> vbNullString Then
            MaxProc.RunFile "C:\Program Files\Google\Chrome\Application\chrome.exe", """" & Link & """"
        Else
            MaxProc.RunFile GetProgramFilesDir & "\Microsoft\Edge\Application\msedge.exe", """" & Link & """"
        End If
    Else
        ActiveDocument.FollowHyperlink Address:=Link, NewWindow:=True ' If the link doesn't work, nothing happens.
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
    On Error GoTo slut
    
    If Selection.Tables.Count > 0 Then
        MsgBox2 TT.A(475), vbOKOnly, TT.Error
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
    Selection.TypeText TT.A(69) & ":"
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
    On Error GoTo fejl

    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord

    Application.ScreenUpdating = False
    If Selection.OMaths.Count > 0 Then
        If Selection.OMaths(1).Type = wdOMathInline Then
            Selection.OMaths(1).Range.Select
            If Selection.OMaths(1).Range.text = "Type equation here." Or Selection.OMaths(1).Range.text = "Skriv ligningen her." Then
            Else
                Selection.Collapse wdCollapseStart
                Selection.MoveRight wdCharacter, 1
            End If
            Selection.TypeText TT.A(62) & ": "
        Else
            Selection.OMaths(1).Range.Select
            Selection.Collapse wdCollapseStart
            If Selection.OMaths(1).Range.text = "Type equation here." Or Selection.OMaths(1).Range.text = "Skriv ligningen her." Then
                Selection.MoveRight wdCharacter, 1
            End If
            Selection.TypeText TT.A(62) & ": "
        End If
    Else
        Selection.OMaths.Add Selection.Range
        Selection.TypeText TT.A(62) & ": "
    End If
    Selection.Collapse wdCollapseEnd
        
    GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
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
        s = omax.ReadEquation2(r)
'        s = omax.ReadEquation(r)
        hopover = False
        If InStr(VBA.LCase$(s), "defin") > 0 Then
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
            ElseIf s = VBA.ChrW$(8661) Then
                ResFeltIndex = ResFeltIndex + 1
                ResIndex = 0
                hopover = True
            End If
        End If
Loop While hopover
    
    sr.Select
    ResPos1 = Selection.Range.start
    If Selection.Range.text = "Skriv ligningen her." Then
        ResPos1 = ResPos1 - 1 ' if already empty, selection is for some reason 1 character too many
    End If
    s = Replace(s, VBA.ChrW$(8289), "") ' function sign sin(x) otherwise becomes si*n(x). also problem with other functions
    Selection.text = s
    
GoTo slut
fejl:
    ResIndex = 0
    ResFeltIndex = 0
    ResPos2 = 0
    ResPos1 = 0
slut:
    ActiveWindow.VerticalPercentScrolled = scrollpos
End Sub

Function KlipTilLigmed(text As String, ByVal indeks As Integer) As String
' returns the last part of the text to the first position counted from the end for = or approximately equal
' = in the sum sign is ignored

    Dim posligmed As Integer
    Dim possumtegn As Integer
    Dim posca As Integer
    Dim poseller As Integer
    Dim pos As Integer
    Dim arr(20) As String
    Dim i As Integer
    
    Do ' go back to nearest equal sign
        posligmed = InStr(text, "=")
        possumtegn = InStr(text, VBA.ChrW$(8721))
        posca = InStr(text, VBA.ChrW$(8776))
        poseller = InStr(text, VBA.ChrW$(8744))
        
        pos = Len(text)
    '    pos = posligmed
        If posligmed > 0 And posligmed < pos Then pos = posligmed
        If posca > 0 And posca < pos Then pos = posca
        If poseller > 0 And poseller < pos Then pos = poseller
        
        If possumtegn > 0 And possumtegn < pos Then ' if there is a sum sign, there is an = sign as part of it
            pos = 0
        End If
        If pos = Len(text) Then pos = 0
        If pos > 0 Then
            arr(i) = Left$(text, pos - 1)
            text = Right$(text, Len(text) - pos)
            i = i + 1
        Else
            arr(i) = text
        End If
    Loop While pos > 0
    
    If indeks = i Then ResIndex = -1  ' global variable marks that there are no more to the left
    If i = 0 Then
        KlipTilLigmed = text
        ResIndex = -1
    Else
        KlipTilLigmed = arr(i - indeks)
    End If
    
    ' remove returns and spaces etc.
'    s = Replace(s, vbCrLf, "")
    KlipTilLigmed = Replace(KlipTilLigmed, vbCr, "")
    KlipTilLigmed = Replace(KlipTilLigmed, VBA.ChrW$(11), "")
'    s = Replace(s, vbLf, "")
    KlipTilLigmed = Replace(KlipTilLigmed, VBA.ChrW$(8744), "") 'eller tegn
'    KlipTilLigmed = Replace(KlipTilLigmed, " ", "")
    KlipTilLigmed = Trim$(KlipTilLigmed)
    
    If InStr(KlipTilLigmed, "/") > 0 Then KlipTilLigmed = "  " & KlipTilLigmed
    
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
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
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
    On Error GoTo fejl
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
        MsgBox TT.A(111) & FilNavn, vbOKOnly, TT.Error
    End If
#End If

    GoTo slut
fejl:
    MsgBox TT.A(111) & FilNavn, vbOKOnly, TT.Error
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
        tip = TT.A(325)
    Case 1
        tip = TT.A(326)
    Case 2
        tip = TT.A(327)
    Case 3
        tip = TT.A(328)
    Case 4
        tip = TT.A(329)
    Case 5
        tip = TT.A(329)
    Case 6
        tip = TT.A(331) & "   x_1   ->   x" & VBA.ChrW$(8321)
    Case 7
        tip = TT.A(332) & VBA.ChrW$(955)
    Case 8
        tip = TT.A(333)
    Case 9
        tip = TT.A(334)
    Case 10
        tip = TT.A(335)
    Case 11
        tip = TT.A(336)
    Case 12
        tip = "(a+b)" & VBA.ChrW$(178) & " = a" & VBA.ChrW$(178) & " + b" & VBA.ChrW$(178) & " + 2ab"
    Case 13
        tip = "(a+b)(a-b) = a" & VBA.ChrW$(178) & " - b" & VBA.ChrW$(178)
    Case 14
        tip = "(a-b)" & VBA.ChrW$(178) & " = a" & VBA.ChrW$(178) & " + b" & VBA.ChrW$(178) & " - 2ab"
    Case 15
        tip = "(a" & VBA.ChrW$(183) & "b)" & VBA.ChrW$(7510) & " = a" & VBA.ChrW$(7510) & VBA.ChrW$(183) & "b" & VBA.ChrW$(7510)
    Case 16
        tip = TT.A(337) & AntalB
    Case 17
        tip = "log(a" & VBA.ChrW$(7495) & ") = b" & VBA.ChrW$(183) & "log(a)"
    Case 18
        tip = "log(a/b) = log(a) - log(b)"
    Case 19
        tip = "log(a" & VBA.ChrW$(183) & "b) = log(a) + log(b)"
    Case 20
        tip = "\int      ->      " & VBA.ChrW$(8747)
    Case 21
        tip = TT.A(338)
    Case 22
        tip = TT.A(339)
    Case 23
        tip = TT.A(340)
    Case 24
        tip = "(a/b)" & VBA.ChrW$(7510) & " = a" & VBA.ChrW$(7510) & "/b" & VBA.ChrW$(7510)
    Case 25
        tip = "a/b + c/d = (ad+bc)/bd"
    Case 26
        tip = TT.A(341)
    Case 27
        tip = TT.A(342)
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
        ufq.Label_text.Caption = TT.A(166) 'unit off
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
        OutUnits = InputBox(TT.A(167), TT.A(168), OutUnits)
        If InStr(OutUnits, "/") > 0 Or InStr(OutUnits, "*") > 0 Or InStr(OutUnits, "^") > 0 Then
            MsgBox2 TT.A(343), vbOKOnly, TT.Error
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
    On Error GoTo fejl
    Dim NewVersion As String, p As Integer, News As String, s As String
    Dim UpdateNow As Boolean, PartnerShip As Boolean
   
    If GetInternetConnectedState = False Then
        If Not RunSilent Then MsgBox TT.A(63), vbOKOnly, TT.Error
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
    On Error GoTo fejl
    
#If Mac Then
    If PartnerShip Then
        s = RunScript("GetHTML", "https://www.eduap.com/download/info/wordmatmacversionP.txt")
    Else
        s = RunScript("GetHTML", "https://www.eduap.com/download/info/wordmatmacversion.txt")
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
            MsgBox2 TT.A(112), vbOKOnly, TT.Error
        End If
        GoTo slut
    End If
    NewVersion = s
    p = InStr(NewVersion, vbLf)
    If p > 0 Then
        News = Right$(NewVersion, Len(NewVersion) - p)
        NewVersion = Trim$(Left$(NewVersion, p - 1))
    Else ' mac
        p = InStr(NewVersion, vbCr)
        If p > 0 Then
            News = Right$(NewVersion, Len(NewVersion) - p)
            NewVersion = Trim$(Left$(NewVersion, p - 1))
        End If
    End If
   
    If Len(NewVersion) = 0 Or Len(NewVersion) > 15 Then
        If Not RunSilent Then
            MsgBox2 TT.A(112), vbOKOnly, TT.Error
        End If
        GoTo slut
    End If
   
   Dim MajorVersion As Integer, MinorVersion As Integer, AppPatchVersion As Integer, arr() As String
   Dim NewMajorVersion As Integer, NewMinorVersion As Integer, NewPatchVersion As Integer, NewEkstraInfoVersion As Integer
   arr = Split(AppVersion & PatchVersion, ".")
   MajorVersion = CInt(arr(0))
   If UBound(arr) > 0 Then MinorVersion = CInt(arr(1))
   If UBound(arr) > 1 Then AppPatchVersion = CInt(arr(2))
   arr = Split(NewVersion, ".")
   NewMajorVersion = CInt(arr(0))
   If UBound(arr) > 0 Then NewMinorVersion = CInt(arr(1))
   If UBound(arr) > 1 Then NewPatchVersion = CInt(arr(2))
   
   If NewMajorVersion > MajorVersion Then
        UpdateNow = True
    ElseIf NewMajorVersion = MajorVersion And NewMinorVersion > MinorVersion Then
        UpdateNow = True
   ElseIf (Not RunSilent) And NewMajorVersion = MajorVersion And NewMinorVersion = MinorVersion And NewPatchVersion > AppPatchVersion Then ' if updatebutton was pressed, then also look for patch version
        UpdateNow = True
   End If
   
   ' deprecated
'    If IsNumeric(AppVersion) And IsNumeric(NewVersion) Then
'        If val(AppVersion) < val(NewVersion) Then UpdateNow = True
'    Else
'        If AppVersion <> NewVersion Then UpdateNow = True
'    End If
    
    If UpdateNow Then
        If PartnerShip Then
            If MsgBox2(TT.A(21) & News & vbCrLf & vbCrLf & TT.A(64), vbOKCancel, TT.A(23)) = vbOK Then
                On Error Resume Next
                Documents.Save NoPrompt:=True, OriginalFormat:=wdOriginalDocumentFormat
                On Error GoTo Install2
                Application.Run macroname:="PUpdateWordMat"
                On Error GoTo fejl
            End If
        Else
Install2:
            On Error GoTo fejl
            MsgBox2 TT.A(21) & News & vbCrLf & TT.A(22) & vbCrLf & vbCrLf & "", vbOKOnly, TT.A(23)
            If TT.LangNo = 1 Then
                OpenLink "https://www.eduap.com/da/wordmat/"
            Else
                OpenLink "https://www.eduap.com/wordmat/"
            End If
        End If
    Else
        If Not RunSilent Then
            MsgBox2 TT.A(344) & " " & AppNavn & " v." & AppVersion, vbOKOnly, "No Update"
        End If
    End If
   
    GoTo slut
fejl:
    If Not RunSilent Then
        If MsgBox2(TT.A(581) & AppVersion, vbOKCancel, TT.Error) = vbOK Then
            If TT.LangNo = 1 Then
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
        .Open "GET", URL & "?cb=" & Timer() * 100, False  ' timer ensures that it is not a cached version
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
            If Right$(ConvertNumberToString, 2) = ".0" Then
                ConvertNumberToString = Left$(ConvertNumberToString, Len(ConvertNumberToString) - 2)
            ElseIf Right$(ConvertNumberToString, 2) = ",0" Then
                ConvertNumberToString = Left$(ConvertNumberToString, Len(ConvertNumberToString) - 2)
            End If
        End If
    End If
    If DecSeparator = "," Then
        ConvertNumberToString = Replace(ConvertNumberToString, ".", ",")
    Else
        ConvertNumberToString = Replace(ConvertNumberToString, ",", ".")
    End If
'    ConvertNumberToString = Replace(Replace(n, ",", "."), "E", VBA.chrw$(183) & "10^(")
    ConvertNumberToString = Replace(ConvertNumberToString, "E", VBA.ChrW$(183) & "10^(")
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
    If r.text = VBA.ChrW$(11) Then ' if there is shift-enter at the end, replace with regular return
        r.text = VBA.ChrW$(13)
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
    On Error GoTo fejl
    If Selection.Range.Tables.Count = 0 Then
        MsgBox TT.A(871), vbOKOnly, TT.Error
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
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
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
        MsgBox TT.A(901), vbOKOnly, TT.Error
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
            Tabel.Cell(i, j).Range.text = dd.TabelsCelle(i, j)
        Next
    Next

    GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
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
    
    Application.OMathAutoCorrect.Entries.Add "\bi", VBA.ChrW$(8660) ' biimplicative arrows
    Application.OMathAutoCorrect.Entries.Add "\imp", VBA.ChrW$(8658) ' implikationarrow right
End Sub

Sub RestartWordMat()
    RestartMaxima
End Sub

Sub InsertNumberedEquation(Optional AskRef As Boolean = False)
    Dim t As Table, F As Field, ccut As Boolean
    Dim placement As Integer
    On Error GoTo fejl
    
    If AskRef Then
        Dim EqName As String
        UserFormEnterEquationRef.Show
        EqName = UserFormEnterEquationRef.EquationName    'Replace(InputBox(TT.A(5), TT.A(4), "Eq"), " ", "")
        If EqName = vbNullString Then GoTo slut
    End If
    
    Application.ScreenUpdating = False

    If Selection.Tables.Count > 0 Then
        MsgBox2 TT.A(872), vbOKOnly, TT.Error
        Exit Sub
    End If

    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord

    If Selection.OMaths.Count > 0 Then
        If Not Selection.OMaths(1).Range.text = vbNullString Then
            Selection.OMaths(1).Range.Cut
            ccut = True
            'there may sometimes be a remainder of a mathematical field
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

    'Insert number
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
fejl:
    MsgBox2 TT.ErrorGeneral, vbOKOnly, TT.Error
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
        Selection.TypeText TT.A(833) & " "
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
fejl:
    MsgBox2 TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
End Sub

Sub SetEquationNumber()
On Error GoTo fejl
    Application.ScreenUpdating = False
    Dim F As Field, f2 As Field, n As String, p As Integer, arr As Variant
    
    If Selection.Fields.Count = 0 Then
        MsgBox TT.A(345), vbOKOnly, TT.Error
        Exit Sub
    End If
    
    Set F = Selection.Fields(1)
    If Selection.Fields.Count = 1 And InStr(F.Code.text, "LISTNUM") > 0 Then
        n = InputBox(TT.A(346), TT.A(6), "1")
        p = InStr(F.Code.text, "\S")
        If p > 0 Then
            F.Code.text = Left$(F.Code.text, p - 1)
        End If
        F.Code.text = F.Code.text & "\S" & n
        F.Update
    ElseIf Selection.Fields.Count = 1 Or Selection.Fields.Count = 2 And InStr(F.Code.text, "WMeq") > 0 Then
        If Selection.Fields.Count = 2 Then
            Set f2 = Selection.Fields(2)
            n = InputBox(TT.A(346), TT.A(6), F.result & "." & f2.result)
            arr = Split(n, ".")
            If UBound(arr) > 0 Then
                SetFieldNo F, CStr(arr(0))
                SetFieldNo f2, CStr(arr(1))
            Else
                SetFieldNo F, CStr(arr(0))
            End If
        Else
            n = InputBox(TT.A(346), TT.A(6), F.result)
            SetFieldNo F, n
        End If
    End If
    
    ActiveDocument.Fields.Update
    GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
End Sub

Sub SetFieldNo(F As Field, n As String)
    Dim p As Integer, p2 As Integer
    On Error GoTo fejl
    p = InStr(F.Code.text, "\r")
    p2 = InStr(F.Code.text, "\c")
    If p2 > 0 And p2 < p Then p = p2
    If p > 0 Then
        F.Code.text = Left$(F.Code.text, p - 1)
    End If
    F.Code.text = F.Code.text & "\r" & n & " \c"
    F.Update
    ActiveDocument.Fields.Update
    GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
End Sub

Sub InsertEquationHeadingNo()
    Dim result As Long
    On Error GoTo fejl
    result = MsgBox(TT.A(348), vbYesNoCancel, TT.A(8))
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
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
End Sub

Sub UpdateEquationNumbers()
    On Error GoTo fejl
    ActiveDocument.Fields.Update
    GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
End Sub

Sub OpenLatexTemplate()
On Error GoTo fejl
    Documents.Add Template:=GetWordMatDir() & "WordDocs/LatexWordTemplate.dotx"
GoTo slut
fejl:
    MsgBox TT.ErrorGeneral, vbOKOnly, TT.Error
slut:
End Sub

Sub DeleteNormalDotm()
    Dim UserDir As String

    MsgBox TT.A(681), vbOKOnly, ""
#If Mac Then
    Dim p As Integer
    UserDir = MacScript("return POSIX path of (path to home folder)")
    p = InStr(UserDir, "/Containers")
    UserDir = Left$(UserDir, p) & "Group Containers/UBF8T346G9.Office/User Content.localized/Templates.localized"
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
        If Len(KB.KeyString) = 8 And Left$(KB.KeyString, 7) = "Option+" Then
            KB.Clear
        End If
    Next
#Else
    For Each KB In KeyBindings
        If LCase$(Left$(KB.Command, 8)) = "wordmat." Then
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
   Local_Document_Path = Doc.path
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
   RemoveTopFolderFromPath = Mid$(ShortName, InStr(ShortName, "\") + 1)
End Function

Function RemoveFileNameFromPath(ByVal ShortName As String) As String
   RemoveFileNameFromPath = Mid$(ShortName, 1, Len(ShortName) - InStr(StrReverse(ShortName), "\"))
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
   
   ExtractTag = Mid$(s, p + Len(StartTag), p2 - p - Len(StartTag))
End Function

Sub SetMathAutoCorrect()
' cant be run from automacros
    If MaximaGangeTegn = VBA.ChrW$(183) Then
        Call Application.OMathAutoCorrect.Entries.Add(Name:="*", Value:=VBA.ChrW$(183))
    ElseIf MaximaGangeTegn = VBA.ChrW$(215) Then
        Call Application.OMathAutoCorrect.Entries.Add(Name:="*", Value:=VBA.ChrW$(215))
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
        SubDir = Trim$(SubDir)
        If Right$(SubDir, 1) <> "\" Then SubDir = SubDir & "\"
    End If
    If InstallLocation = "All" Then
        GetWordMatDir = GetProgramFilesDir() & "\WordMat\"
        If Dir(GetWordMatDir & SubDir, vbDirectory) = vbNullString Then
            GetWordMatDir = Environ("AppData") & "\WordMat\"
            If Dir(GetWordMatDir & SubDir, vbDirectory) = vbNullString Then
                MsgBox "WordMat folder could not be found", vbOKOnly, TT.Error
            End If
        End If
    Else
        GetWordMatDir = Environ("AppData") & "\WordMat\"
        If Dir(GetWordMatDir & SubDir, vbDirectory) = vbNullString Then
            GetWordMatDir = GetProgramFilesDir() & "\WordMat\"
            If Dir(GetWordMatDir & SubDir, vbDirectory) = vbNullString Then
                MsgBox "WordMat folder could not be found", vbOKOnly, TT.Error
            End If
        End If
    End If
#End If
End Function

Sub NewEquation()
    Dim r As Range
    On Error GoTo fejl
    On Error Resume Next
    
    If Selection.OMaths.Count = 0 Then
        With Selection.Font
            .Bold = False
            .ColorIndex = wdAuto
            .Italic = False
            .Underline = False
        End With
        Set r = Selection.OMaths.Add(Selection.Range)
    ElseIf Selection.Tables.Count = 0 Then
        With Selection.Font
            .Bold = False
            .ColorIndex = wdAuto
            .Italic = False
            .Underline = False
        End With
        If Selection.OMaths(1).Range.text = vbNullString Then
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
fejl:
    MsgBox2 TT.ErrorGeneral, vbOKOnly, TT.Error
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
    DefS = Replace(DefS, "SymVecta", ChrW$(&H20D7))
    DefS = Replace(DefS, "matrix", vbNullString)
            
    DefS = Replace(DefS, "*", MaximaGangeTegn)
        
    DefS = Replace(DefS, "%pi", ChrW$(&H3C0))
    DefS = Replace(DefS, "%i", "i")
    DefS = Replace(DefS, "log(", "ln(")
    DefS = Replace(DefS, "log10(", "log(")
    DefS = Replace(DefS, "^(x)", ChrW$(&H2E3))
    DefS = Replace(DefS, "^(2)", ChrW$(&HB2))
    DefS = Replace(DefS, "^(3)", ChrW$(&HB3))
    DefS = Replace(DefS, "^(4)", ChrW$(&H2074))
    DefS = Replace(DefS, "^(5)", ChrW$(&H2075))
    DefS = Replace(DefS, "^(6)", ChrW$(&H2076))
    DefS = Replace(DefS, "^(7)", ChrW$(&H2077))
    DefS = Replace(DefS, "^(8)", ChrW$(&H2078))
    DefS = Replace(DefS, "^(9)", ChrW$(&H2079))
    DefS = Replace(DefS, "^(-1)", ChrW$(&H207B) & ChrW$(&HB9))
    DefS = Replace(DefS, "^(-2)", ChrW$(&H207B) & ChrW$(&HB2))
    DefS = Replace(DefS, "^(-3)", ChrW$(&H207B) & ChrW$(&HB3))
        
    DefS = Replace(DefS, "_0(", ChrW$(&H2080) & "(")
    DefS = Replace(DefS, "_1(", ChrW$(&H2081) & "(")
    DefS = Replace(DefS, "_2(", ChrW$(&H2082) & "(")
    DefS = Replace(DefS, "_3(", ChrW$(&H2083) & "(")
    DefS = Replace(DefS, "_4(", ChrW$(&H2084) & "(")
    DefS = Replace(DefS, "_5(", ChrW$(&H2085) & "(")
    DefS = Replace(DefS, "_6(", ChrW$(&H2086) & "(")
    DefS = Replace(DefS, "_7(", ChrW$(&H2087) & "(")
    DefS = Replace(DefS, "_8(", ChrW$(&H2088) & "(")
    DefS = Replace(DefS, "_9(", ChrW$(&H2089) & "(")
    DefS = Replace(DefS, "_a(", ChrW$(&H2090) & "(")
    DefS = Replace(DefS, "_x(", ChrW$(&H2093) & "(")
    DefS = Replace(DefS, "_n(", ChrW$(&H2099) & "(")
        
    DefS = Replace(DefS, "minf", "-" & ChrW$(&H221E))
    DefS = Replace(DefS, "inf", ChrW$(&H221E))
        
    DefS = Replace(DefS, "sqrt(", ChrW$(&H221A) & "(")
    DefS = Replace(DefS, "NIntegrate(", ChrW$(&H222B) & "(")
    DefS = Replace(DefS, "Integrate(", ChrW$(&H222B) & "(")
    DefS = Replace(DefS, "integrate(", ChrW$(&H222B) & "(")
    DefS = Replace(DefS, "<=", VBA.ChrW$(8804))
    DefS = Replace(DefS, ">=", VBA.ChrW$(8805))
    DefS = Replace(DefS, "ae", ChrW$(230))
    DefS = Replace(DefS, "oe", ChrW$(248))
    DefS = Replace(DefS, "aa", ChrW$(229))
    DefS = Replace(DefS, "AE", ChrW$(198))
    DefS = Replace(DefS, "OE", ChrW$(216))
    DefS = Replace(DefS, "AA", ChrW$(197))
        
    'greek letters
    DefS = Replace(DefS, "gamma", VBA.ChrW$(915))    ' big gamma
    DefS = Replace(DefS, "Delta", VBA.ChrW$(916))
    DefS = Replace(DefS, "delta", VBA.ChrW$(948))
    DefS = Replace(DefS, "alpha", VBA.ChrW$(945))
    DefS = Replace(DefS, "beta", VBA.ChrW$(946))
    DefS = Replace(DefS, "gammaLB", VBA.ChrW$(947))
    DefS = Replace(DefS, "theta", VBA.ChrW$(952))
    DefS = Replace(DefS, "Theta", VBA.ChrW$(920))
    DefS = Replace(DefS, "lambda", VBA.ChrW$(955))
    DefS = Replace(DefS, "Lambda", VBA.ChrW$(923))
    DefS = Replace(DefS, "mu", VBA.ChrW$(956))
    DefS = Replace(DefS, "rho", VBA.ChrW$(961))
    DefS = Replace(DefS, "sigma", VBA.ChrW$(963))
    DefS = Replace(DefS, "Sigma", VBA.ChrW$(931))
    DefS = Replace(DefS, "varphi", VBA.ChrW$(966))
    DefS = Replace(DefS, "phi", VBA.ChrW$(981))
    DefS = Replace(DefS, "Phi", VBA.ChrW$(934))
    DefS = Replace(DefS, "varepsilon", VBA.ChrW$(949))
    DefS = Replace(DefS, "epsilon", VBA.ChrW$(1013))
    DefS = Replace(DefS, "psi", VBA.ChrW$(968))
    DefS = Replace(DefS, "Psi", VBA.ChrW$(936))
    DefS = Replace(DefS, "Xi", VBA.ChrW$(926))
    DefS = Replace(DefS, "xi", VBA.ChrW$(958))
    DefS = Replace(DefS, "Chi", VBA.ChrW$(935))
    DefS = Replace(DefS, "chi", VBA.ChrW$(967))
    DefS = Replace(DefS, "Pi", VBA.ChrW$(928))
    DefS = Replace(DefS, "tau", VBA.ChrW$(964))
    DefS = Replace(DefS, "greek-nu", VBA.ChrW$(957))
    DefS = Replace(DefS, "kappa", VBA.ChrW$(954))
    DefS = Replace(DefS, "eta", VBA.ChrW$(951))
    DefS = Replace(DefS, "zeta", VBA.ChrW$(950))
    DefS = Replace(DefS, "omega", VBA.ChrW$(969))    ' small omega
    
    DefS = Replace(DefS, "((x))", "(x)")
    
    DefS = ReplaceFakeVarNamesBack(DefS)
    
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
    tid = Timer
    n = TT.LangNo
    MsgBox Timer - tid
End Sub

Public Sub ShowComputerID()
    Dim MyData As New DataObject, HID As String
    On Error GoTo fejl
    HID = GetHardwareUUID()
    MyData.SetText HID
    MyData.PutInClipboard
    If MsgBox("Hardware ID is used to identify your computer" & vbCrLf & HID & vbCrLf & vbCrLf & "The ID has been copied to the clipboard" & vbCrLf & "Do you wish to send this ID to Eduap to initiate payment of WordMat+", vbYesNo, "Hardware UUID") = vbYes Then
        Application.Run macroname:="InitiatePayment"
    End If
    GoTo slut
fejl:
    MsgBox "Error" & vbCrLf & Err.Description, vbOKOnly, "Error"
slut:
End Sub

Function StringToUTF8Bytes(s As String) As Byte()
' Convert a VBA string (UTF-16) to a UTF-8 byte array
' saving a string to a text file using print #fh,text   saves the file in ANSI encoding (Windows 1252 on Windows) which is single byte encoding
' This does not support special characters like greek letters. On windows this can be circumvented using adodb.stream, but it is not Mac compatible
' This function can be used with WriteUTF8File to write a string to a text file in UTF-8 encoding which is variable length (1-4 bytes pr. character)
' (encoding to several bytes is not straight forward, every byte has a special bit signature.)

    Dim bytes() As Byte
    Dim i As Long, ch As Long
    Dim pos As Long
    
    ' Allocate maximum space (worst case: 4 bytes per char)
    ReDim bytes(1 To Len(s) * 4)
    pos = 1
    
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        
        If ch < &H80 Then
            ' 1-byte ASCII
            bytes(pos) = ch
            pos = pos + 1
        ElseIf ch < &H800 Then
            ' 2-byte UTF-8
            bytes(pos) = &HC0 Or ((ch And &H7C0) \ &H40)
            bytes(pos + 1) = &H80 Or (ch And &H3F)
            pos = pos + 2
        ElseIf ch < &H10000 Then
            ' 3-byte UTF-8
            bytes(pos) = &HE0 Or ((ch And &HF000) \ &H1000)
            bytes(pos + 1) = &H80 Or ((ch And &HFC0) \ &H40)
            bytes(pos + 2) = &H80 Or (ch And &H3F)
            pos = pos + 3
        Else
            ' 4-byte UTF-8 (surrogate pair / supplementary plane)
            ' Not handled in this simple example
            bytes(pos) = &H3F ' ?
            pos = pos + 1
        End If
    Next i
    
    ' Trim the array
    ReDim Preserve bytes(1 To pos - 1)
    StringToUTF8Bytes = bytes
End Function

' Write a UTF-8 file
Sub WriteUTF8File(filePath As String, text As String)
' Writes a string to a text file in UTF-8 encoding. Which supports special characters on Windows and Mac. The normal print #fh, text   creates ANSI encoded files
' The trick is to use the function StringToUTF8Bytes to create a bytearray of the UTF-file, and then open file for binary write,

    Dim b() As Byte
    Dim fnum As Integer
    
    ' Convert string to UTF-8 byte array
    b = StringToUTF8Bytes(text)
    
    ' Write bytes to file
    fnum = FreeFile
    Open filePath For Binary As #fnum
    Put #fnum, , b
    Close #fnum
End Sub
Sub InsertGradtegn()
    Selection.TypeText ChrW(176)
End Sub
' Example usage
Sub TestUTF8Write()
    Dim s As String
    s = "Hello µ world €"
    
    WriteUTF8File "/Users/yourname/Desktop/test_utf8.txt", s
End Sub

Sub SpeedTest()
Dim t1 As Single, t2 As Single, i As Long, n As Long, s As String
    n = 1000000
    
    s = "Dette er en streng"
    
    t1 = Timer
    For i = 0 To n
'        s = Replace(s, "en", "et")
        s = vbNewLine
    Next
    t1 = Timer - t1
    
    t2 = Timer
    For i = 0 To n
'        If InStr(s, "en") <> 0 Then s = Replace(s, "en", "et")
        s = ChrW$(183)
    Next
    t2 = Timer - t2
    
    MsgBox "t1: " & t1 & vbNewLine & "t2: " & t2 & vbNewLine & "Factor t1/t2: " & t1 / t2, vbOKOnly, "Speedtest"
    
End Sub
