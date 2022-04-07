Attribute VB_Name = "MaximaSettingsModule"
Option Explicit
Public UFMSettings As UserFormMaximaSettings
Public SettingsRead As Boolean
Private mforklaring As Boolean
Private mkommando As Boolean
Private mExact As Integer
Private mgangetegn As String
Private mradians As Boolean
Private mcifre As Integer
Private mseparator As Boolean
Private mlistseparator As String
Private mdecseparator As String
Private mComplex As Boolean
Private mUnits As Boolean
Private mvidnot As Boolean
Private mlogout As Integer
Private mexcelembed As Boolean
Private malltrig As Boolean
Private moutunits As String
Private mautostart As Boolean
Private mantalb As LongPtr
Private mbigfloat As Boolean
Private mIndex As Boolean
Private mshowassum As Boolean
Private mpolaroutput As Boolean
Private mgraphapp As Integer
Private mlanguage As Integer
Private mlmset As Boolean ' l*oe*sninger som l*oe*sningsm*ae*ngde
Private mdasdiffchr As Boolean
Private mlatexstart As String
Private mlatexslut As String
Private mlatexunits As Boolean
Private mConvertTexWithMaxima As Boolean
Private meqnumplacement As Boolean
Private meqnumtype As Boolean
Private maskref As Boolean
Private mBackupType As Integer
Private mbackupno As Integer
Private mbackupmaxno As Integer
Private mbackuptime As Integer
Private mLatexPreamble As String
Private mLatexSectionNumbering As Boolean
Private mLatexDocumentclass As Integer
Private mLatexFontsize As String
Private mLatexWordMargins As Boolean
Private mLatexTitlePage As Integer
Private mLatexTOC As Integer
Private mCASengine As Integer



Public Sub ReadAllSettingsFromRegistry()
Dim setn As Integer
On Error Resume Next
    
    mforklaring = CBool(GetRegSetting("Forklaring"))
    mkommando = CBool(GetRegSetting("MaximaCommand"))
    mExact = GetRegSetting("Exact")
    mradians = CBool(GetRegSetting("Radians"))
    mcifre = GetRegSetting("SigFig")
    mComplex = CBool(GetRegSetting("Complex"))
    mlmset = CBool(GetRegSetting("SolveBoolOrSet"))
    mUnits = CBool(GetRegSetting("Units"))
    mvidnot = CBool(GetRegSetting("VidNot"))
    mlogout = GetRegSetting("LogOutput")
    mexcelembed = CBool(GetRegSetting("ExcelEmbed"))
    malltrig = CBool(GetRegSetting("AllTrig"))
    moutunits = GetRegSettingString("OutUnits")
    mautostart = CBool(GetRegSetting("AutoStart"))
    mbigfloat = CBool(GetRegSetting("BigFloat"))
    mantalb = GetRegSettingLong("AntalBeregninger")
    mIndex = CBool(GetRegSetting("Index"))
    mshowassum = CBool(GetRegSetting("ShowAssum"))
    mpolaroutput = CBool(GetRegSetting("PolarOutput"))
    mgraphapp = CInt(GetRegSettingLong("GraphApp"))
#If Mac Then
    If mgraphapp = 0 Then mgraphapp = 2 ' gnuplot er pt afskaffet, s*aa* der bruges webapp
#End If
    mlanguage = CInt(GetRegSettingLong("Language"))
    mdasdiffchr = CBool(GetRegSetting("dAsDiffChr"))
    mlatexstart = GetRegSettingString("LatexStart")
    mlatexslut = GetRegSettingString("LatexSlut")
    mlatexunits = CBool(GetRegSetting("LatexUnits"))
    mConvertTexWithMaxima = CBool(GetRegSetting("ConvertTexWithMaxima"))
    meqnumplacement = CBool(GetRegSetting("EqNumPlacement"))
    meqnumtype = CBool(GetRegSetting("EqNumType"))
    maskref = CBool(GetRegSetting("EqAskRef"))
    mBackupType = CInt(GetRegSettingLong("BackupType"))
    mbackupno = CInt(GetRegSettingLong("BackupNo"))
    mbackupmaxno = CInt(GetRegSettingLong("BackupMaxNo"))
    mbackuptime = CInt(GetRegSettingLong("BackupTime"))
    mLatexSectionNumbering = CBool(GetRegSetting("LatexSectionNumbering"))
    mLatexDocumentclass = CInt(GetRegSettingLong("LatexDocumentclass"))
    mLatexFontsize = GetRegSettingString("LatexFontsize")
    mLatexWordMargins = CBool(GetRegSetting("LatexWordMargins"))
    mLatexTitlePage = CInt(GetRegSettingLong("LatexTitlePage"))
    mLatexTOC = CInt(GetRegSettingLong("LatexToc"))
    mCASengine = CInt(GetRegSettingLong("CASengine"))
    
    mseparator = CBool(GetRegSetting("Separator"))
    If mseparator Then
        mdecseparator = "."
        mlistseparator = ","
    Else
        mdecseparator = ","
        mlistseparator = ";"
    End If

    setn = GetRegSetting("Gangetegn")
    If setn = 0 Then
        mgangetegn = VBA.ChrW(183)
    ElseIf setn = 1 Then
        mgangetegn = VBA.ChrW(215)
    Else
        mgangetegn = "*"
    End If
    
    If mlatexstart = vbNullString Then
        LatexStart = "$"
    End If
    If mlatexslut = vbNullString Then
        LatexSlut = "$"
    End If
    
    SettingsRead = True
End Sub
Public Sub SetAllDefaultRegistrySettings()
' s*ae*tter alle indstillinger til default, men kun hvis de ikke eksisterer i forvejen
On Error Resume Next
    If Not RegKeyExists("HKCU\SOFTWARE\WORDMAT\Settings\Forklaring") Then
'    If MsgBox("Indstillingerne kan ikke findes. Vil du genoprette og nulstille alle indstillinger?", vbOKCancel, Sprog.Error) Then
    MaximaForklaring = True
    MaximaKommando = False
    MaximaExact = 0
    Radians = False
    MaximaCifre = 7
    MaximaSeparator = False
    MaximaGangeTegn = "prik"
    MaximaComplex = False
    LmSet = False
    MaximaUnits = False
    MaximaVidNotation = False
    MaximaLogOutput = 0
    ExcelIndlejret = False
    AllTrig = False
    OutUnits = ""
    AutoStart = False
    Antalberegninger = 0
    SettCheckForUpdate = False
    MaximaIndex = False
    PolarOutput = False
#If Mac Then
    GraphApp = 4
#Else
    GraphApp = 0
#End If
    LanguageSetting = 0
    dAsDiffChr = False
    LatexStart = "$"
    LatexSlut = "$"
    LatexUnits = False
    ConvertTexWithMaxima = False
    EqNumPlacement = False
    EqNumType = False
    EqAskRef = False
    BackupType = 0
    BackupNo = 1
    BackupMaxNo = 20
    BackupTime = 5
    LatexSectionNumbering = True
    LatexDocumentclass = 0
    LatexFontsize = 11
    LatexWordMargins = False
    LatexTitlePage = 0
    LatexTOC = 0
    CASengine = 0
    
    GenerateKeyboardShortcuts
'    End If
    End If
    If Not RegKeyExists("HKCU\SOFTWARE\WORDMAT\Settings\BigFloat") Then
        mbigfloat = False
    End If

End Sub

Public Property Get MaximaForklaring() As Boolean
    MaximaForklaring = mforklaring
End Property
Public Property Let MaximaForklaring(xval As Boolean)
    SetRegSetting "Forklaring", Abs(CInt(xval))
    mforklaring = xval
End Property
Public Property Get MaximaKommando() As Boolean
    MaximaKommando = mkommando
End Property
Public Property Let MaximaKommando(xval As Boolean)
    SetRegSetting "MaximaCommand", Abs(CInt(xval))
    mkommando = xval
End Property
Public Property Get MaximaExact() As Integer
' 0 - auto
' 1 - exact
' 2 - num
    MaximaExact = mExact
End Property
Public Property Let MaximaExact(ByVal xval As Integer)
    SetRegSetting "Exact", xval
    mExact = xval
    If Not (MaxProc Is Nothing) Then
        MaxProc.Exact = xval
    End If
End Property
Public Property Get Radians() As Boolean
    Radians = mradians
End Property
Public Property Let Radians(ByVal text As Boolean)
    SetRegSetting "Radians", Abs(CInt(text))
    mradians = text
End Property
Public Property Get MaximaCifre() As Integer
    If mcifre > 1 Then
        MaximaCifre = mcifre
    Else
        ReadAllSettingsFromRegistry
        If mcifre < 2 Then
            MaximaCifre = 7
        Else
            MaximaCifre = mcifre
        End If
    End If
End Property
Public Property Let MaximaCifre(ByVal cifr As Integer)
    SetRegSetting "SigFig", cifr
    mcifre = cifr
End Property
Public Property Get MaximaSeparator() As Boolean
' true er punktum som decimalseparator
    MaximaSeparator = mseparator
End Property
Public Property Let MaximaSeparator(xval As Boolean)
    SetRegSetting "Separator", Abs(CInt(xval))
    mseparator = xval
    If xval Then
        mdecseparator = "."
        mlistseparator = ","
    Else
        mdecseparator = ","
        mlistseparator = ";"
    End If
End Property
Public Property Get DecSeparator() As String
' decimalseparator
    DecSeparator = mdecseparator
End Property
Public Property Let DecSeparator(ByVal Sep As String)
    mdecseparator = Sep
End Property
Public Property Get ListSeparator() As String
' listseparator
    ListSeparator = mlistseparator
End Property
Public Property Let ListSeparator(ByVal Sep As String)
    mlistseparator = Sep
End Property
Public Property Get MaximaGangeTegn() As String
    MaximaGangeTegn = mgangetegn
End Property
Public Property Let MaximaGangeTegn(ByVal nygtegn As String)
    If nygtegn = "prik" Then
        SetRegSetting "Gangetegn", 0
        mgangetegn = VBA.ChrW(183)
    ElseIf nygtegn = "x" Then
        SetRegSetting "Gangetegn", 1
        mgangetegn = VBA.ChrW(215)
    Else '*
        SetRegSetting "Gangetegn", 2
        mgangetegn = "*"
    End If
End Property
Public Property Get MaximaComplex() As Boolean
    MaximaComplex = mComplex
End Property
Public Property Let MaximaComplex(xval As Boolean)
    SetRegSetting "Complex", Abs(CInt(xval))
    mComplex = xval
End Property
Public Property Get LmSet() As Boolean
    LmSet = mlmset
End Property
Public Property Let LmSet(xval As Boolean)
    SetRegSetting "SolveBoolOrSet", Abs(CInt(xval))
    mlmset = xval
End Property
Public Property Get MaximaUnits() As Boolean
    MaximaUnits = mUnits
End Property
Public Property Let MaximaUnits(xval As Boolean)
    SetRegSetting "Units", Abs(CInt(xval))
    mUnits = xval
End Property
Public Property Get MaximaVidNotation() As Boolean
    MaximaVidNotation = mvidnot
End Property
Public Property Let MaximaVidNotation(vidval As Boolean)
    SetRegSetting "VidNot", Abs(CInt(vidval))
    mvidnot = vidval
End Property
Public Property Get MaximaLogOutput() As Integer
    MaximaLogOutput = mlogout
End Property
Public Property Let MaximaLogOutput(xval As Integer)
    SetRegSetting "LogOutput", xval
    mlogout = xval
End Property
Public Property Get ExcelIndlejret() As Boolean
    ExcelIndlejret = mexcelembed
End Property
Public Property Let ExcelIndlejret(vidval As Boolean)
    SetRegSetting "ExcelEmbed", Abs(CInt(vidval))
    mexcelembed = vidval
End Property
Public Property Get AllTrig() As Boolean
    AllTrig = malltrig
End Property
Public Property Let AllTrig(xval As Boolean)
    SetRegSetting "AllTrig", Abs(CInt(xval))
    malltrig = xval
End Property
Public Property Get EqNumPlacement() As Boolean
    EqNumPlacement = meqnumplacement
End Property
Public Property Let EqNumPlacement(ByVal text As Boolean)
    SetRegSetting "EqNumPlacement", Abs(CInt(text))
    meqnumplacement = text
End Property
Public Property Get EqNumType() As Boolean
    EqNumType = meqnumtype
End Property
Public Property Let EqNumType(ByVal text As Boolean)
    SetRegSetting "EqNumType", Abs(CInt(text))
    meqnumtype = text
End Property
Public Property Get EqAskRef() As Boolean
    EqAskRef = maskref
End Property
Public Property Let EqAskRef(ByVal text As Boolean)
    SetRegSetting "EqAskRef", Abs(CInt(text))
    maskref = text
End Property


Public Property Get OutUnits() As String
    OutUnits = moutunits
End Property
Public Property Let OutUnits(ByVal text As String)
    text = Replace(text, "kwh", "kWh")
    text = Replace(text, "hz", "Hz")
    text = Replace(text, "HZ", "Hz")
    text = Replace(text, "bq", "Bq")
    text = Replace(text, "ev", "eV")
    SetRegSettingString "OutUnits", text
    moutunits = text
End Property
Public Property Get AutoStart() As Boolean
    AutoStart = mautostart
End Property
Public Property Let AutoStart(xval As Boolean)
    SetRegSetting "AutoStart", Abs(CInt(xval))
    mautostart = xval
End Property

#If VBA7 Then
Public Property Get Antalberegninger() As LongPtr
    Antalberegninger = mantalb
End Property
Public Property Let Antalberegninger(xval As LongPtr)
    SetRegSettingLong "AntalBeregninger", xval
    mantalb = xval
End Property
#Else
Public Property Get Antalberegninger() As Long
    Antalberegninger = mantalb
End Property
Public Property Let Antalberegninger(xval As Long)
    SetRegSettingLong "AntalBeregninger", xval
    mantalb = xval
End Property
#End If

Public Property Get MaximaBigFloat() As Boolean
    MaximaBigFloat = mbigfloat
End Property
Public Property Let MaximaBigFloat(xval As Boolean)
    SetRegSetting "BigFloat", Abs(CInt(xval))
    mbigfloat = xval
End Property
Public Property Get PolarOutput() As Boolean
    PolarOutput = mpolaroutput
End Property
Public Property Let PolarOutput(xval As Boolean)
    SetRegSetting "PolarOutput", Abs(CInt(xval))
    mpolaroutput = xval
End Property
Public Property Get MaximaIndex() As Boolean
    MaximaIndex = mIndex
End Property
Public Property Let MaximaIndex(xval As Boolean)
    SetRegSetting "Index", Abs(CInt(xval))
    mIndex = xval
End Property
Public Property Get ShowAssum() As Boolean
    ShowAssum = mshowassum
End Property
Public Property Let ShowAssum(xval As Boolean)
    SetRegSetting "ShowAssum", Abs(CInt(xval))
    mshowassum = xval
End Property
Public Property Get GraphApp() As Integer
    GraphApp = mgraphapp
End Property
Public Property Let GraphApp(xval As Integer)
    SetRegSetting "GraphApp", xval
    mgraphapp = xval
End Property
Public Property Get LanguageSetting() As Integer
    LanguageSetting = mlanguage
End Property
Public Property Let LanguageSetting(xval As Integer)
    SetRegSetting "Language", xval
    mlanguage = xval
End Property
Public Property Get dAsDiffChr() As Boolean
    dAsDiffChr = mdasdiffchr
End Property
Public Property Let dAsDiffChr(ByVal text As Boolean)
    SetRegSetting "dAsDiffChr", Abs(CInt(text))
    mdasdiffchr = text
End Property
Public Property Get LatexStart() As String
    LatexStart = mlatexstart
End Property
Public Property Let LatexStart(ByVal Sep As String)
    SetRegSettingString "LatexStart", Sep
    mlatexstart = Sep
End Property
Public Property Get LatexSlut() As String
    LatexSlut = mlatexslut
End Property
Public Property Let LatexSlut(ByVal Sep As String)
    SetRegSettingString "LatexSlut", Sep
    mlatexslut = Sep
End Property
Public Property Get LatexUnits() As Boolean
    LatexUnits = mlatexunits
End Property
Public Property Let LatexUnits(ByVal text As Boolean)
    SetRegSetting "LatexUnits", Abs(CInt(text))
    mlatexunits = text
End Property
Public Property Get ConvertTexWithMaxima() As Boolean
    ConvertTexWithMaxima = mConvertTexWithMaxima
End Property
Public Property Let ConvertTexWithMaxima(ByVal text As Boolean)
    SetRegSetting "ConvertTexWithMaxima", Abs(CInt(text))
    mConvertTexWithMaxima = text
End Property
Public Property Get LatexWordMargins() As Boolean
    LatexWordMargins = mLatexWordMargins
End Property
Public Property Let LatexWordMargins(xval As Boolean)
    SetRegSetting "LatexWordMargins", Abs(CInt(xval))
    mLatexWordMargins = xval
End Property

Public Property Get SettCheckForUpdate() As Boolean
    SettCheckForUpdate = CBool(GetRegSetting("CheckForUpdate"))
End Property
Public Property Let SettCheckForUpdate(xval As Boolean)
    SetRegSetting "CheckForUpdate", Abs(CInt(xval))
End Property
Public Property Get BackupType() As Integer
    If mBackupType = 0 Then
        ReadAllSettingsFromRegistry
    End If
    BackupType = mBackupType
End Property
Public Property Let BackupType(xval As Integer)
    SetRegSetting "BackupType", xval
    mBackupType = xval
End Property
Public Property Get BackupNo() As Integer
    BackupNo = mbackupno
End Property
Public Property Let BackupNo(xval As Integer)
    SetRegSetting "BackupNo", xval
    mbackupno = xval
End Property
Public Property Get BackupMaxNo() As Integer
    BackupMaxNo = mbackupmaxno
End Property
Public Property Let BackupMaxNo(xval As Integer)
    SetRegSetting "BackupMaxNo", xval
    mbackupmaxno = xval
End Property
Public Property Get BackupTime() As Integer
    BackupTime = mbackuptime
End Property
Public Property Let BackupTime(xval As Integer)
    SetRegSetting "BackupTime", xval
    mbackuptime = xval
End Property

Public Property Get LatexPreamble() As String
   If mLatexPreamble = "" Then
      Dim filnavn As String
      filnavn = Environ("AppData") & "\WordMat\WordMatLatexPreamble.tex"
      If Dir(filnavn) <> "" Then mLatexPreamble = ReadTextfileToString(filnavn)
   End If
   LatexPreamble = mLatexPreamble
End Property
Public Property Let LatexPreamble(ByVal preAmble As String)
    Dim filnavn As String
    mLatexPreamble = preAmble
    filnavn = Environ("AppData") & "\WordMat\WordMatLatexPreamble.tex"
    If Dir(filnavn) <> "" Then Kill filnavn
    WriteTextfileToString filnavn, preAmble
End Property
Public Property Get LatexSectionNumbering() As Boolean
    LatexSectionNumbering = mLatexSectionNumbering
End Property
Public Property Let LatexSectionNumbering(xval As Boolean)
    SetRegSetting "LatexSectionNumbering", Abs(CInt(xval))
    mLatexSectionNumbering = xval
End Property
Public Property Get LatexDocumentclass() As Integer
    LatexDocumentclass = mLatexDocumentclass
End Property
Public Property Let LatexDocumentclass(xval As Integer)
    SetRegSetting "LatexDocumentclass", xval
    mLatexDocumentclass = xval
End Property
Public Property Get LatexFontsize() As String
   LatexFontsize = mLatexFontsize
End Property
Public Property Let LatexFontsize(ByVal xval As String)
    mLatexFontsize = xval
    SetRegSettingString "LatexFontsize", xval
End Property
Public Property Get LatexTitlePage() As Integer
    LatexTitlePage = mLatexTitlePage
End Property
Public Property Let LatexTitlePage(xval As Integer)
    SetRegSetting "LatexTitlePage", xval
    mLatexTitlePage = xval
End Property
Public Property Get LatexTOC() As Integer
    LatexTOC = mLatexTOC
End Property
Public Property Let LatexTOC(xval As Integer)
    SetRegSetting "LatexTOC", xval
    mLatexTOC = xval
End Property
Public Property Get CASengine() As Integer
    CASengine = mCASengine
End Property
Public Property Let CASengine(xval As Integer)
    SetRegSetting "CASengine", xval
    mCASengine = xval
End Property


Private Function GetRegSetting(Key As String) As Integer
    GetRegSetting = RegKeyRead("HKCU\SOFTWARE\WORDMAT\Settings\" & Key)
End Function
Private Sub SetRegSetting(Key As String, val As Integer)
    RegKeySave "HKCU\SOFTWARE\WORDMAT\Settings\" & Key, val, "REG_DWORD"
End Sub

#If VBA7 Then
Public Sub SetRegSettingLong(Key As String, val As LongPtr)
    RegKeySave "HKCU\SOFTWARE\WORDMAT\Settings\" & Key, val, "REG_DWORD"
End Sub
Public Function GetRegSettingLong(Key As String) As LongPtr
    GetRegSettingLong = CLngPtr(RegKeyRead("HKCU\SOFTWARE\WORDMAT\Settings\" & Key))
End Function
#Else
Public Sub SetRegSettingLong(Key As String, val As Long)
    RegKeySave "HKCU\SOFTWARE\WORDMAT\Settings\" & Key, val, "REG_DWORD"
End Sub
Public Function GetRegSettingLong(Key As String) As Long
    GetRegSettingLong = CLng(RegKeyRead("HKCU\SOFTWARE\WORDMAT\Settings\" & Key))
End Function
#End If

Private Function GetRegSettingString(Key As String) As String
    GetRegSettingString = RegKeyRead("HKCU\SOFTWARE\WORDMAT\Settings\" & Key)
End Function
Private Sub SetRegSettingString(Key As String, ByVal val As String)
    RegKeySave "HKCU\SOFTWARE\WORDMAT\Settings\" & Key, val, "REG_SZ"
End Sub

