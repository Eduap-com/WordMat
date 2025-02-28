Attribute VB_Name = "MaximaSettingsModule"
Option Explicit

Enum KeybShortcut
    NoShortcut = -1
    InsertNewEquation = 1
    NewNumEquation
    beregnudtryk
    SolveEquation
    Define
    sletdef
    ShowGraph
    Formelsamling
    OmskrivUdtryk
    SolveDiffEq
    ExecuteMaximaCommand
    PrevResult
    SettingsForm
    ToggleNumExact
    ToggleUnitsOnOff
    ConvertEquationToLatex
    OpenLatexPDF
    InsertRefToEqution
End Enum


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
Private mlmset As Boolean ' løsninger som løsningsmængde
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
Private mLastUpdateCheck As String
Private mRegAppVersion As String
Private mDllConnType As Integer ' 0=reg dll  1=direct dll   2=wsh (only Maxima)
Private mInstallLocation As String ' All AppData
Private mDecOutType As Integer ' 1 =dec, 2=bet cif, 3=vidnot
Private mUseVBACAS As Integer  ' 0 = not loaded  1=no  2=yes1
Private mUseShellOnMac As Boolean ' for when applescripttask does not work
Private mSettShortcutAltM As Integer
Private mSettShortcutAltM2 As Integer
Private mSettShortcutAltB As Integer
Private mSettShortcutAltL As Integer
Private mSettShortcutAltD As Integer
Private mSettShortcutAltS As Integer
Private mSettShortcutAltP As Integer
Private mSettShortcutAltF As Integer
Private mSettShortcutAltO As Integer
Private mSettShortcutAltR As Integer
Private mSettShortcutAltJ As Integer
Private mSettShortcutAltN As Integer
Private mSettShortcutAltE As Integer
Private mSettShortcutAltT As Integer
Private mSettShortcutAltQ As Integer
Private mSettShortcutAltG As Integer
Private mSettShortcutAltGr As Integer

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
    If mgraphapp = 0 Then mgraphapp = 2 ' gnuplot er pt afskaffet, så der bruges webapp
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
    mLastUpdateCheck = GetRegSettingString("LastUpdateCheck")
    mDllConnType = CInt(GetRegSetting("DllConnType"))
    mInstallLocation = GetRegSetting("InstallLocation")
    mUseVBACAS = GetRegSetting("UseVBACAS")
    mDecOutType = CInt(GetRegSetting("DecOutType"))
#If Mac Then
    mUseShellOnMac = CBool(GetRegSetting("UseShellOnMac"))
#End If

    mSettShortcutAltM = CInt(GetRegSetting("SettShortcutAltM"))
    mSettShortcutAltM2 = CInt(GetRegSetting("SettShortcutAltM2"))
    mSettShortcutAltB = CInt(GetRegSetting("SettShortcutAltB"))
    mSettShortcutAltL = CInt(GetRegSetting("SettShortcutAltL"))
    mSettShortcutAltP = CInt(GetRegSetting("SettShortcutAltP"))
    mSettShortcutAltD = CInt(GetRegSetting("SettShortcutAltD"))
    mSettShortcutAltS = CInt(GetRegSetting("SettShortcutAltS"))
    mSettShortcutAltF = CInt(GetRegSetting("SettShortcutAltF"))
    mSettShortcutAltO = CInt(GetRegSetting("SettShortcutAltO"))
    mSettShortcutAltR = CInt(GetRegSetting("SettShortcutAltR"))
    mSettShortcutAltJ = CInt(GetRegSetting("SettShortcutAltJ"))
    mSettShortcutAltN = CInt(GetRegSetting("SettShortcutAltN"))
    mSettShortcutAltE = CInt(GetRegSetting("SettShortcutAltE"))
    mSettShortcutAltT = CInt(GetRegSetting("SettShortcutAltT"))
    mSettShortcutAltQ = CInt(GetRegSetting("SettShortcutAltQ"))
    
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
' sætter alle indstillinger til default, men kun hvis de ikke eksisterer i forvejen
On Error Resume Next
    If Not RegKeyExists("HKCU\SOFTWARE\WORDMAT\Settings\Forklaring") Then
'    If MsgBox("Indstillingerne kan ikke findes. Vil du genoprette og nulstille alle indstillinger?", vbOKCancel, Sprog.Error) Then
    MaximaForklaring = True
    MaximaKommando = False
    MaximaExact = 2 ' numerisk
    Radians = False
    MaximaCifre = 7
    MaximaSeparator = False
    MaximaGangeTegn = "prik"
    MaximaComplex = False
    LmSet = False
    MaximaUnits = False
    MaximaLogOutput = 0
    ExcelIndlejret = False
    AllTrig = False
    OutUnits = ""
    AutoStart = False
'    Antalberegninger = 0 ' skal vel aldrig nulstilles
    SettCheckForUpdate = True
    MaximaIndex = False
    PolarOutput = False
#If Mac Then
    GraphApp = 4 ' geogebraweb
#Else
    GraphApp = 4
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
    BackupType = 2 ' spørg ikke
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
    MaximaDecOutType = 2
    SettUseVBACAS = 2
    
    SettShortcutAltM = KeybShortcut.InsertNewEquation
    SettShortcutAltM2 = -1
    SettShortcutAltB = KeybShortcut.beregnudtryk
    SettShortcutAltL = KeybShortcut.SolveEquation
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
Public Property Let Radians(ByVal Text As Boolean)
    SetRegSetting "Radians", Abs(CInt(Text))
    mradians = Text
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
Public Property Get MaximaDecOutType() As Integer
    If mDecOutType = 0 Then
        Dim s As String
        mDecOutType = GetRegSetting("DecOutType")
        If mDecOutType = 0 Then
            mDecOutType = 2
        End If
    End If
    MaximaDecOutType = mDecOutType
End Property
Public Property Let MaximaDecOutType(vidval As Integer)
    SetRegSetting "DecOutType", vidval
    mDecOutType = vidval
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
Public Property Let EqNumPlacement(ByVal Text As Boolean)
    SetRegSetting "EqNumPlacement", Abs(CInt(Text))
    meqnumplacement = Text
End Property
Public Property Get EqNumType() As Boolean
    EqNumType = meqnumtype
End Property
Public Property Let EqNumType(ByVal Text As Boolean)
    SetRegSetting "EqNumType", Abs(CInt(Text))
    meqnumtype = Text
End Property
Public Property Get EqAskRef() As Boolean
    EqAskRef = maskref
End Property
Public Property Let EqAskRef(ByVal Text As Boolean)
    SetRegSetting "EqAskRef", Abs(CInt(Text))
    maskref = Text
End Property
Public Property Get LastUpdateCheck() As String
    LastUpdateCheck = mLastUpdateCheck
End Property
Public Property Let LastUpdateCheck(ByVal Text As String)
    SetRegSettingString "LastUpdateCheck", Text
    mLastUpdateCheck = Text
End Property

Public Property Get OutUnits() As String
    OutUnits = moutunits
End Property
Public Property Let OutUnits(ByVal Text As String)
    Text = Replace(Text, "kwh", "kWh")
    Text = Replace(Text, "hz", "Hz")
    Text = Replace(Text, "HZ", "Hz")
    Text = Replace(Text, "bq", "Bq")
    Text = Replace(Text, "ev", "eV")
    SetRegSettingString "OutUnits", Text
    moutunits = Text
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
Public Property Let dAsDiffChr(ByVal Text As Boolean)
    SetRegSetting "dAsDiffChr", Abs(CInt(Text))
    mdasdiffchr = Text
End Property
Public Property Let dAsDiffChrTemp(ByVal Text As Boolean)
    mdasdiffchr = Text
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
Public Property Let LatexUnits(ByVal Text As Boolean)
    SetRegSetting "LatexUnits", Abs(CInt(Text))
    mlatexunits = Text
End Property
Public Property Get ConvertTexWithMaxima() As Boolean
    ConvertTexWithMaxima = mConvertTexWithMaxima
End Property
Public Property Let ConvertTexWithMaxima(ByVal Text As Boolean)
    SetRegSetting "ConvertTexWithMaxima", Abs(CInt(Text))
    mConvertTexWithMaxima = Text
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
      Dim FilNavn As String
      FilNavn = Environ("AppData") & "\WordMat\WordMatLatexPreamble.tex"
      If Dir(FilNavn) <> "" Then mLatexPreamble = ReadTextfileToString(FilNavn)
   End If
   LatexPreamble = mLatexPreamble
End Property
Public Property Let LatexPreamble(ByVal preAmble As String)
    Dim FilNavn As String
    mLatexPreamble = preAmble
    FilNavn = Environ("AppData") & "\WordMat\WordMatLatexPreamble.tex"
    If Dir(FilNavn) <> "" Then Kill FilNavn
    WriteTextfileToString FilNavn, preAmble
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
Public Property Let CASengineTempOnly(xval As Integer)
    mCASengine = xval
End Property
Public Property Get CASengineRegOnly() As Integer
    mCASengine = CInt(GetRegSettingLong("CASengine"))
    CASengineRegOnly = mCASengine
End Property
Public Property Get RegAppVersion() As String
    If mRegAppVersion <> vbNullString Then
        RegAppVersion = mRegAppVersion
    Else
        RegAppVersion = GetRegSettingString("AppVersion")
        mRegAppVersion = RegAppVersion
    End If
End Property
Public Property Let RegAppVersion(ByVal v As String)
    SetRegSettingString "AppVersion", v
    mRegAppVersion = v
End Property
Public Property Get DllConnType() As Integer ' 0=reg dll  1=direct dll   2=wsh (only Maxima)
    DllConnType = mDllConnType
End Property
Public Property Let DllConnType(xval As Integer)
    SetRegSetting "DllConnType", xval
    mDllConnType = xval
End Property
Public Property Get InstallLocation() As String
    If mInstallLocation <> vbNullString Then
        InstallLocation = mInstallLocation
    Else
        InstallLocation = GetRegSettingString("InstallLocation")
        mInstallLocation = InstallLocation
    End If
End Property
Public Property Let InstallLocation(ByVal l As String)
    SetRegSettingString "InstallLocation", l
    mInstallLocation = l
End Property
Public Property Get SettUseVBACAS() As Boolean
    If QActivePartnership Then
        If mUseVBACAS = 0 Then
            mUseVBACAS = GetRegSetting("UseVBACAS")
        End If
        If mUseVBACAS = 0 Then
            SetRegSetting "UseVBACAS", 2 ' 1=no, 2=yes
            mUseVBACAS = 2
        End If

        SettUseVBACAS = CBool(mUseVBACAS - 1)
    Else ' if no partnership VBACAS will fail
        SettUseVBACAS = False
    End If
End Property
Public Property Let SettUseVBACAS(xval As Boolean)
    mUseVBACAS = Abs(CInt(xval) + 1)
    SetRegSetting "UseVBACAS", mUseVBACAS
End Property
Public Property Get UseShellOnMac() As Boolean
    UseShellOnMac = mUseShellOnMac
End Property
Public Property Let UseShellOnMac(xval As Boolean)
    SetRegSetting "UseShellOnMac", Abs(CInt(xval))
    mUseShellOnMac = xval
End Property
Public Property Get SettShortcutAltM() As Integer
    If mSettShortcutAltM = 0 Then
        mSettShortcutAltM = CInt(GetRegSetting("SettShortcutAltM"))
    End If
    SettShortcutAltM = mSettShortcutAltM
End Property
Public Property Let SettShortcutAltM(xval As Integer)
    SetRegSetting "SettShortcutAltM", xval
    mSettShortcutAltM = xval
End Property
Public Property Get SettShortcutAltM2() As Integer
    If mSettShortcutAltM2 <= 0 Then
        mSettShortcutAltM2 = CInt(GetRegSetting("SettShortcutAltM2"))
    End If
    SettShortcutAltM2 = mSettShortcutAltM2
End Property
Public Property Let SettShortcutAltM2(xval As Integer)
    SetRegSetting "SettShortcutAltM2", xval
    mSettShortcutAltM2 = xval
End Property
Public Property Get SettShortcutAltB() As Integer
    If mSettShortcutAltB = 0 Then
        mSettShortcutAltB = CInt(GetRegSetting("SettShortcutAltB"))
    End If
    SettShortcutAltB = mSettShortcutAltB
End Property
Public Property Let SettShortcutAltB(xval As Integer)
    SetRegSetting "SettShortcutAltB", xval
    mSettShortcutAltB = xval
End Property
Public Property Get SettShortcutAltL() As Integer
    If mSettShortcutAltL = 0 Then
        mSettShortcutAltL = CInt(GetRegSetting("SettShortcutAltL"))
    End If
    SettShortcutAltL = mSettShortcutAltL
End Property
Public Property Let SettShortcutAltL(xval As Integer)
    SetRegSetting "SettShortcutAltL", xval
    mSettShortcutAltL = xval
End Property
Public Property Get SettShortcutAltD() As Integer
    If mSettShortcutAltD = 0 Then
        mSettShortcutAltD = CInt(GetRegSetting("SettShortcutAltD"))
    End If
    SettShortcutAltD = mSettShortcutAltD
End Property
Public Property Let SettShortcutAltD(xval As Integer)
    SetRegSetting "SettShortcutAltD", xval
    mSettShortcutAltD = xval
End Property
Public Property Get SettShortcutAltS() As Integer
    If mSettShortcutAltS = 0 Then
        mSettShortcutAltS = CInt(GetRegSetting("SettShortcutAltS"))
    End If
    SettShortcutAltS = mSettShortcutAltS
End Property
Public Property Let SettShortcutAltS(xval As Integer)
    SetRegSetting "SettShortcutAltS", xval
    mSettShortcutAltS = xval
End Property
Public Property Get SettShortcutAltP() As Integer
    If mSettShortcutAltP = 0 Then
        mSettShortcutAltP = CInt(GetRegSetting("SettShortcutAltP"))
    End If
    SettShortcutAltP = mSettShortcutAltP
End Property
Public Property Let SettShortcutAltP(xval As Integer)
    SetRegSetting "SettShortcutAltP", xval
    mSettShortcutAltP = xval
End Property
Public Property Get SettShortcutAltF() As Integer
    If mSettShortcutAltF = 0 Then
        mSettShortcutAltF = CInt(GetRegSetting("SettShortcutAltF"))
    End If
    SettShortcutAltF = mSettShortcutAltF
End Property
Public Property Let SettShortcutAltF(xval As Integer)
    SetRegSetting "SettShortcutAltF", xval
    mSettShortcutAltF = xval
End Property
Public Property Get SettShortcutAltO() As Integer
    If mSettShortcutAltO = 0 Then
        mSettShortcutAltO = CInt(GetRegSetting("SettShortcutAltO"))
    End If
    SettShortcutAltO = mSettShortcutAltO
End Property
Public Property Let SettShortcutAltO(xval As Integer)
    SetRegSetting "SettShortcutAltO", xval
    mSettShortcutAltO = xval
End Property
Public Property Get SettShortcutAltR() As Integer
    If mSettShortcutAltR = 0 Then
        mSettShortcutAltR = CInt(GetRegSetting("SettShortcutAltR"))
    End If
    SettShortcutAltR = mSettShortcutAltR
End Property
Public Property Let SettShortcutAltR(xval As Integer)
    SetRegSetting "SettShortcutAltR", xval
    mSettShortcutAltR = xval
End Property
Public Property Get SettShortcutAltJ() As Integer
    If mSettShortcutAltJ = 0 Then
        mSettShortcutAltJ = CInt(GetRegSetting("SettShortcutAltJ"))
    End If
    SettShortcutAltJ = mSettShortcutAltJ
End Property
Public Property Let SettShortcutAltJ(xval As Integer)
    SetRegSetting "SettShortcutAltJ", xval
    mSettShortcutAltJ = xval
End Property
Public Property Get SettShortcutAltN() As Integer
    If mSettShortcutAltN = 0 Then
        mSettShortcutAltN = CInt(GetRegSetting("SettShortcutAltN"))
    End If
    SettShortcutAltN = mSettShortcutAltN
End Property
Public Property Let SettShortcutAltN(xval As Integer)
    SetRegSetting "SettShortcutAltN", xval
    mSettShortcutAltN = xval
End Property
Public Property Get SettShortcutAltE() As Integer
    If mSettShortcutAltE = 0 Then
        mSettShortcutAltE = CInt(GetRegSetting("SettShortcutAltE"))
    End If
    SettShortcutAltE = mSettShortcutAltE
End Property
Public Property Let SettShortcutAltE(xval As Integer)
    SetRegSetting "SettShortcutAltE", xval
    mSettShortcutAltE = xval
End Property
Public Property Get SettShortcutAltT() As Integer
    If mSettShortcutAltT = 0 Then
        mSettShortcutAltT = CInt(GetRegSetting("SettShortcutAltT"))
    End If
    SettShortcutAltT = mSettShortcutAltT
End Property
Public Property Let SettShortcutAltT(xval As Integer)
    SetRegSetting "SettShortcutAltT", xval
    mSettShortcutAltT = xval
End Property
Public Property Get SettShortcutAltQ() As Integer
    If mSettShortcutAltQ = 0 Then
        mSettShortcutAltQ = CInt(GetRegSetting("SettShortcutAltQ"))
    End If
    SettShortcutAltQ = mSettShortcutAltQ
End Property
Public Property Let SettShortcutAltQ(xval As Integer)
    SetRegSetting "SettShortcutAltQ", xval
    mSettShortcutAltQ = xval
End Property
Public Property Get SettShortcutAltG() As Integer
    If mSettShortcutAltG = 0 Then
        mSettShortcutAltG = CInt(GetRegSetting("SettShortcutAltG"))
    End If
    SettShortcutAltG = mSettShortcutAltG
End Property
Public Property Let SettShortcutAltG(xval As Integer)
    SetRegSetting "SettShortcutAltG", xval
    mSettShortcutAltG = xval
End Property
Public Property Get SettShortcutAltGr() As Integer
    If mSettShortcutAltGr = 0 Then
        mSettShortcutAltGr = CInt(GetRegSetting("SettShortcutAltGr"))
    End If
    SettShortcutAltGr = mSettShortcutAltGr
End Property
Public Property Let SettShortcutAltGr(xval As Integer)
    SetRegSetting "SettShortcutAltGr", xval
    mSettShortcutAltGr = xval
End Property


'------------------- registry functions --------------------
Public Function GetReg(key As String) As String
    GetReg = GetRegSettingString(key)
End Function
Public Function GetRegSetting(key As String) As Integer
    Dim s As String
    s = RegKeyRead("HKCU\SOFTWARE\WORDMAT\Settings\" & key)
    If s = vbNullString Then
        GetRegSetting = 0
    Else
        On Error Resume Next
        GetRegSetting = CInt(s)
    End If
End Function
Private Sub SetRegSetting(ByVal key As String, ByVal val As Integer)
    RegKeySave "HKCU\SOFTWARE\WORDMAT\Settings\" & key, val, "REG_DWORD"
End Sub

#If VBA7 Then
Public Sub SetRegSettingLong(key As String, val As LongPtr)
    RegKeySave "HKCU\SOFTWARE\WORDMAT\Settings\" & key, val, "REG_DWORD"
End Sub
Public Function GetRegSettingLong(key As String) As LongPtr
    GetRegSettingLong = CLngPtr(RegKeyRead("HKCU\SOFTWARE\WORDMAT\Settings\" & key))
End Function
#Else
Public Sub SetRegSettingLong(key As String, val As Long)
    RegKeySave "HKCU\SOFTWARE\WORDMAT\Settings\" & key, val, "REG_DWORD"
End Sub
Public Function GetRegSettingLong(key As String) As Long
    GetRegSettingLong = CLng(RegKeyRead("HKCU\SOFTWARE\WORDMAT\Settings\" & key))
End Function
#End If

Private Function GetRegSettingString(key As String) As String
    GetRegSettingString = RegKeyRead("HKCU\SOFTWARE\WORDMAT\Settings\" & key)
End Function
Private Sub SetRegSettingString(key As String, ByVal val As String)
    RegKeySave "HKCU\SOFTWARE\WORDMAT\Settings\" & key, val, "REG_SZ"
End Sub

