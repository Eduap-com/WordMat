Attribute VB_Name = "ModuleSettings"
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
    GradTegn
    Open3DPLot
End Enum

Public UFMSettings As UserFormSettings
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
Private mantalb As LongPtr
Private mbigfloat As Boolean
Private mIndex As Boolean
Private mshowassum As Boolean
Private mpolaroutput As Boolean
Private mgraphapp As Integer
Private mlanguage As Integer
Private mlmset As Boolean ' solutions as a set
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
Private mUseCodeFile As Boolean
Private mUseCodeBlocks As Boolean
Private mOutputColor As Integer  ' 0 = wdauto ellers WdColorIndex enumeration (Word) 1= black, 2= blue

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

Public ShowMenuDone As Boolean
Private mSMformula As Boolean
Private mSMformulaOld As Boolean
Private mSMSettingAEN As Boolean
Private mSMSettingUnits As Boolean
Private mSMSettingDecimals As Boolean
Private mSMSettingRadians As Boolean
Private mSMCASCalculate As Boolean
Private mSMCASReduceMenu As Boolean
Private mSMCASReduce As Boolean
Private mSMCASFactor As Boolean
Private mSMCASExpand As Boolean
Private mSMCASDifferentiate As Boolean
Private mSMCASIntegrate As Boolean
Private mSMCASMaximaCommand As Boolean
Private mSMCASSolve As Boolean
Private mSMCASNSolve As Boolean
Private mSMCASEliminate As Boolean
Private mSMCASTest As Boolean
Private mSMCASSolveDE As Boolean
Private mSMCASSolveDEC As Boolean
Private mSMDef As Boolean
Private mSMDefDelete As Boolean
Private mSMDefFunction As Boolean
Private mSMDefConstants As Boolean
Private mSMGraphs As Boolean
Private mSMGraphsGeogebraSuite As Boolean
Private mSMGraphsGeogebra5 As Boolean
Private mSMGraphsGnuPlot As Boolean
Private mSMGraphsGraph As Boolean
Private mSMGraphsExcel As Boolean
Private mSMGraphsDirectionField As Boolean
Private mSMGraphs3D As Boolean
Private mSMGraphs3DSolidRev As Boolean
Private mSMGraphsStatistics As Boolean
Private mSMGraphsStatisticsUgrup As Boolean
Private mSMGraphsStatisticsGrup As Boolean
Private mSMGraphsStatisticsStickChart As Boolean
Private mSMGraphsStatisticsHistogram As Boolean
Private mSMGraphsStatisticsStair As Boolean
Private mSMGraphsStatisticsSumCurve As Boolean
Private mSMGraphsStatisticsBoxPlot As Boolean
Private mSMProb As Boolean
Private mSMProbRegression As Boolean
Private mSMProbRegressionTable As Boolean
Private mSMProbRegressionLin As Boolean
Private mSMProbRegressionExp As Boolean
Private mSMProbRegressionPow As Boolean
Private mSMProbRegressionPol As Boolean
Private mSMProbRegressionSin As Boolean
Private mSMProbRegressionExcel As Boolean
Private mSMProbRegressionUser As Boolean
Private mSMProbDistributions As Boolean
Private mSMProbDistributionsBinomial As Boolean
Private mSMProbDistributionsNormal As Boolean
Private mSMProbDistributionsChi As Boolean
Private mSMProbDistributionsT As Boolean
Private mSMProbTests As Boolean
Private mSMProbTestsBinomial As Boolean
Private mSMProbTestsNormal As Boolean
Private mSMProbTestsChi As Boolean
Private mSMProbTestsSimulation As Boolean
Private mSMOther As Boolean
Private mSMOtherNewEq As Boolean
Private mSMOtherNewEqNum As Boolean
Private mSMOtherLatex As Boolean
Private mSMOtherTriangle As Boolean
Private mSMOtherFigurs As Boolean
Private mSMOtherTable As Boolean
Private mSMCodeBlock As Boolean

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
    mbigfloat = CBool(GetRegSetting("BigFloat"))
    mantalb = val(GetRegSetting("AntalBeregninger"))
    mIndex = CBool(GetRegSetting("Index"))
    mshowassum = CBool(GetRegSetting("ShowAssum"))
    mpolaroutput = CBool(GetRegSetting("PolarOutput"))
    mgraphapp = val(GetRegSetting("GraphApp"))
#If Mac Then
    If mgraphapp = 0 Then mgraphapp = 2 ' gnuplot is not available on mac
#End If
    mlanguage = val(GetRegSetting("Language"))
    mdasdiffchr = CBool(GetRegSetting("dAsDiffChr"))
    mlatexstart = GetRegSettingString("LatexStart")
    mlatexslut = GetRegSettingString("LatexSlut")
    mlatexunits = CBool(GetRegSetting("LatexUnits"))
    mConvertTexWithMaxima = CBool(GetRegSetting("ConvertTexWithMaxima"))
    meqnumplacement = CBool(GetRegSetting("EqNumPlacement"))
    meqnumtype = CBool(GetRegSetting("EqNumType"))
    maskref = CBool(GetRegSetting("EqAskRef"))
    mBackupType = val(GetRegSetting("BackupType"))
    mbackupno = val(GetRegSetting("BackupNo"))
    mbackupmaxno = val(GetRegSetting("BackupMaxNo"))
    mbackuptime = val(GetRegSetting("BackupTime"))
    mLatexSectionNumbering = CBool(GetRegSetting("LatexSectionNumbering"))
    mLatexDocumentclass = val(GetRegSettingLong("LatexDocumentclass"))
    mLatexFontsize = GetRegSettingString("LatexFontsize")
    mLatexWordMargins = CBool(GetRegSetting("LatexWordMargins"))
    mLatexTitlePage = val(GetRegSettingLong("LatexTitlePage"))
    mLatexTOC = val(GetRegSettingLong("LatexToc"))
    mCASengine = val(GetRegSetting("CASengine"))
    mLastUpdateCheck = GetRegSettingString("LastUpdateCheck")
    mDllConnType = val(GetRegSetting("DllConnType"))
    mInstallLocation = GetRegSetting("InstallLocation")
    mUseVBACAS = GetRegSetting("UseVBACAS")
    mDecOutType = val(GetRegSetting("DecOutType"))
    mUseCodeFile = CBool(GetRegSetting("UseCodeFile"))
    mUseCodeBlocks = CBool(GetRegSetting("UseCodeBlocks"))
    mOutputColor = val(GetRegSetting("OutputColor"))
    
    mSettShortcutAltM = val(GetRegSetting("SettShortcutAltM"))
    mSettShortcutAltM2 = val(GetRegSetting("SettShortcutAltM2"))
    mSettShortcutAltB = val(GetRegSetting("SettShortcutAltB"))
    mSettShortcutAltL = val(GetRegSetting("SettShortcutAltL"))
    mSettShortcutAltP = val(GetRegSetting("SettShortcutAltP"))
    mSettShortcutAltD = val(GetRegSetting("SettShortcutAltD"))
    mSettShortcutAltS = val(GetRegSetting("SettShortcutAltS"))
    mSettShortcutAltF = val(GetRegSetting("SettShortcutAltF"))
    mSettShortcutAltO = val(GetRegSetting("SettShortcutAltO"))
    mSettShortcutAltR = val(GetRegSetting("SettShortcutAltR"))
    mSettShortcutAltJ = val(GetRegSetting("SettShortcutAltJ"))
    mSettShortcutAltN = val(GetRegSetting("SettShortcutAltN"))
    mSettShortcutAltE = val(GetRegSetting("SettShortcutAltE"))
    mSettShortcutAltT = val(GetRegSetting("SettShortcutAltT"))
    mSettShortcutAltQ = val(GetRegSetting("SettShortcutAltQ"))
    
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
        mgangetegn = VBA.ChrW$(183)
    ElseIf setn = 1 Then
        mgangetegn = VBA.ChrW$(215)
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
' sets all settings to default, but only if they don't already exist
On Error Resume Next
    If Not RegKeyExists("HKEY_CURRENT_USER\SOFTWARE\WORDMAT\Settings\Forklaring") Then
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
    BackupType = 2 ' dont ask
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
    OutputColor = wdGreen
    ShowAssum = True
    
    
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
    
    End If
    If Not RegKeyExists("HKEY_CURRENT_USER\SOFTWARE\WORDMAT\Settings\BigFloat") Then
        mbigfloat = False
    End If
End Sub

Sub ReadSM(Optional ForceRead As Boolean = False)
    If ForceRead Or Not ShowMenuDone Then
        Dim s As String
        s = GetRegSettingString("ShowMenus")
        If Len(s) < 70 Then
            If IsNumeric(s) Then
                s = Trim$(s) & String$(70 - Len(s), "1")
            Else
                s = "11" ' formula
                s = s & "1111" ' Settings
                s = s & "11111111111111" ' CAS
                s = s & "1111" ' def
                s = s & "11111111111111111" ' graphs
                s = s & "1111111111" ' regression
                s = s & "11111" ' distributions
                s = s & "11111" ' Tests
                s = s & "1111111" ' Other
                s = s & String$(70 - Len(s), "1")
            End If
        End If
        '        On Error Resume Next
        mSMformula = CBool(Mid$(s, 1, 1))
        mSMformulaOld = CBool(Mid$(s, 2, 1))
        mSMSettingAEN = CBool(Mid$(s, 3, 1))
        mSMSettingDecimals = CBool(Mid$(s, 4, 1))
        mSMSettingUnits = CBool(Mid$(s, 5, 1))
        mSMSettingRadians = CBool(Mid$(s, 6, 1))
        mSMCASCalculate = CBool(Mid$(s, 7, 1))
        mSMCASReduceMenu = CBool(Mid$(s, 8, 1))
        mSMCASReduce = CBool(Mid$(s, 9, 1))
        mSMCASFactor = CBool(Mid$(s, 10, 1))
        mSMCASExpand = CBool(Mid$(s, 11, 1))
        mSMCASDifferentiate = CBool(Mid$(s, 12, 1))
        mSMCASIntegrate = CBool(Mid$(s, 13, 1))
        mSMCASMaximaCommand = CBool(Mid$(s, 14, 1))
        mSMCASSolve = CBool(Mid$(s, 15, 1))
        mSMCASNSolve = CBool(Mid$(s, 16, 1))
        mSMCASEliminate = CBool(Mid$(s, 17, 1))
        mSMCASTest = CBool(Mid$(s, 18, 1))
        mSMCASSolveDE = CBool(Mid$(s, 19, 1))
        mSMCASSolveDEC = CBool(Mid$(s, 20, 1))
        mSMDef = CBool(Mid$(s, 21, 1))
        mSMDefDelete = CBool(Mid$(s, 22, 1))
        mSMDefFunction = CBool(Mid$(s, 23, 1))
        mSMDefConstants = CBool(Mid$(s, 24, 1))
        mSMGraphs = CBool(Mid$(s, 25, 1))
        mSMGraphsGeogebraSuite = CBool(Mid$(s, 26, 1))
        mSMGraphsGeogebra5 = CBool(Mid$(s, 27, 1))
        mSMGraphsGnuPlot = CBool(Mid$(s, 28, 1))
        mSMGraphsGraph = CBool(Mid$(s, 29, 1))
        mSMGraphsExcel = CBool(Mid$(s, 30, 1))
        mSMGraphsDirectionField = CBool(Mid$(s, 31, 1))
        mSMGraphs3D = CBool(Mid$(s, 32, 1))
        mSMGraphs3DSolidRev = CBool(Mid$(s, 33, 1))
        mSMGraphsStatistics = CBool(Mid$(s, 34, 1))
        mSMGraphsStatisticsUgrup = CBool(Mid$(s, 35, 1))
        mSMGraphsStatisticsGrup = CBool(Mid$(s, 36, 1))
        mSMGraphsStatisticsStickChart = CBool(Mid$(s, 37, 1))
        mSMGraphsStatisticsHistogram = CBool(Mid$(s, 38, 1))
        mSMGraphsStatisticsStair = CBool(Mid$(s, 39, 1))
        mSMGraphsStatisticsSumCurve = CBool(Mid$(s, 40, 1))
        mSMGraphsStatisticsBoxPlot = CBool(Mid$(s, 41, 1))
        mSMProb = CBool(Mid$(s, 42, 1))
        mSMProbRegression = CBool(Mid$(s, 43, 1))
        mSMProbRegressionTable = CBool(Mid$(s, 44, 1))
        mSMProbRegressionLin = CBool(Mid$(s, 45, 1))
        mSMProbRegressionExp = CBool(Mid$(s, 46, 1))
        mSMProbRegressionPow = CBool(Mid$(s, 47, 1))
        mSMProbRegressionPol = CBool(Mid$(s, 48, 1))
        mSMProbRegressionSin = CBool(Mid$(s, 49, 1))
        mSMProbRegressionExcel = CBool(Mid$(s, 50, 1))
        mSMProbRegressionUser = CBool(Mid$(s, 51, 1))
        mSMProbDistributions = CBool(Mid$(s, 52, 1))
        mSMProbDistributionsBinomial = CBool(Mid$(s, 53, 1))
        mSMProbDistributionsNormal = CBool(Mid$(s, 54, 1))
        mSMProbDistributionsChi = CBool(Mid$(s, 55, 1))
        mSMProbDistributionsT = CBool(Mid$(s, 56, 1))
        mSMProbTests = CBool(Mid$(s, 57, 1))
        mSMProbTestsBinomial = CBool(Mid$(s, 58, 1))
        mSMProbTestsNormal = CBool(Mid$(s, 59, 1))
        mSMProbTestsChi = CBool(Mid$(s, 60, 1))
        mSMProbTestsSimulation = CBool(Mid$(s, 61, 1))
        mSMOther = CBool(Mid$(s, 62, 1))
        mSMOtherNewEq = CBool(Mid$(s, 63, 1))
        mSMOtherNewEqNum = CBool(Mid$(s, 64, 1))
        mSMOtherLatex = CBool(Mid$(s, 65, 1))
        mSMOtherTriangle = CBool(Mid$(s, 66, 1))
        mSMOtherFigurs = CBool(Mid$(s, 67, 1))
        mSMOtherTable = CBool(Mid$(s, 68, 1))
        mSMCodeBlock = CBool(Mid$(s, 69, 1))
        
        ShowMenuDone = True
    End If
End Sub

Public Property Get SMformula() As Boolean
    ReadSM
    SMformula = mSMformula
End Property
Public Property Get SMformulaOld() As Boolean
    ReadSM
    SMformulaOld = mSMformulaOld
End Property
Public Property Get SMSettingAEN() As Boolean
    ReadSM
    SMSettingAEN = mSMSettingAEN
End Property
Public Property Get SMSettingUnits() As Boolean
    ReadSM
    SMSettingUnits = mSMSettingUnits
End Property
Public Property Get SMSettingDecimals() As Boolean
    ReadSM
    SMSettingDecimals = mSMSettingDecimals
End Property
Public Property Get SMSettingRadians() As Boolean
    ReadSM
    SMSettingRadians = mSMSettingRadians
End Property
Public Property Get SMCASCalculate() As Boolean
    ReadSM
    SMCASCalculate = mSMCASCalculate
End Property
Public Property Get SMCASReduceMenu() As Boolean
    ReadSM
    SMCASReduceMenu = mSMCASReduceMenu
End Property
Public Property Get SMCASReduce() As Boolean
    ReadSM
    SMCASReduce = mSMCASReduce
End Property
Public Property Get SMCASFactor() As Boolean
    ReadSM
    SMCASFactor = mSMCASFactor
End Property
Public Property Get SMCASExpand() As Boolean
    ReadSM
    SMCASExpand = mSMCASExpand
End Property
Public Property Get SMCASDifferentiate() As Boolean
    ReadSM
    SMCASDifferentiate = mSMCASDifferentiate
End Property
Public Property Get SMCASIntegrate() As Boolean
    ReadSM
    SMCASIntegrate = mSMCASIntegrate
End Property
Public Property Get SMCASMaximaCommand() As Boolean
    ReadSM
    SMCASMaximaCommand = mSMCASMaximaCommand
End Property
Public Property Get SMCASSolve() As Boolean
    ReadSM
    SMCASSolve = mSMCASSolve
End Property
Public Property Get SMCASNSolve() As Boolean
    ReadSM
    SMCASNSolve = mSMCASNSolve
End Property
Public Property Get SMCASEliminate() As Boolean
    ReadSM
    SMCASEliminate = mSMCASEliminate
End Property
Public Property Get SMCASTest() As Boolean
    ReadSM
    SMCASTest = mSMCASTest
End Property
Public Property Get SMCASSolveDE() As Boolean
    ReadSM
    SMCASSolveDE = mSMCASSolveDE
End Property
Public Property Get SMCASSolveDEC() As Boolean
    ReadSM
    SMCASSolveDEC = mSMCASSolveDEC
End Property
Public Property Get SMDef() As Boolean
    ReadSM
    SMDef = mSMDef
End Property
Public Property Get SMDefDelete() As Boolean
    ReadSM
    SMDefDelete = mSMDefDelete
End Property
Public Property Get SMDefFunction() As Boolean
    ReadSM
    SMDefFunction = mSMDefFunction
End Property
Public Property Get SMDefConstants() As Boolean
    ReadSM
    SMDefConstants = mSMDefConstants
End Property
Public Property Get SMCodeBlock() As Boolean
    ReadSM
    SMCodeBlock = mSMCodeBlock
End Property
Public Property Get SMGraphs() As Boolean
    ReadSM
    SMGraphs = mSMGraphs
End Property
Public Property Get SMGraphsGeogebraSuite() As Boolean
    ReadSM
    SMGraphsGeogebraSuite = mSMGraphsGeogebraSuite
End Property
Public Property Get SMGraphsGeogebra5() As Boolean
    ReadSM
    SMGraphsGeogebra5 = mSMGraphsGeogebra5
End Property
Public Property Get SMGraphsGnuPlot() As Boolean
    ReadSM
    SMGraphsGnuPlot = mSMGraphsGnuPlot
End Property
Public Property Get SMGraphsGraph() As Boolean
    ReadSM
    SMGraphsGraph = mSMGraphsGraph
End Property
Public Property Get SMGraphsExcel() As Boolean
    ReadSM
    SMGraphsExcel = mSMGraphsExcel
End Property
Public Property Get SMGraphsDirectionField() As Boolean
    ReadSM
    SMGraphsDirectionField = mSMGraphsDirectionField
End Property
Public Property Get SMGraphs3D() As Boolean
    ReadSM
    SMGraphs3D = mSMGraphs3D
End Property
Public Property Get SMGraphs3DSolidRev() As Boolean
    ReadSM
    SMGraphs3DSolidRev = mSMGraphs3DSolidRev
End Property
Public Property Get SMGraphsStatistics() As Boolean
    ReadSM
    SMGraphsStatistics = mSMGraphsStatistics
End Property
Public Property Get SMGraphsStatisticsUgrup() As Boolean
    ReadSM
    SMGraphsStatisticsUgrup = mSMGraphsStatisticsUgrup
End Property
Public Property Get SMGraphsStatisticsGrup() As Boolean
    ReadSM
    SMGraphsStatisticsGrup = mSMGraphsStatisticsGrup
End Property
Public Property Get SMGraphsStatisticsStickChart() As Boolean
    ReadSM
    SMGraphsStatisticsStickChart = mSMGraphsStatisticsStickChart
End Property
Public Property Get SMGraphsStatisticsHistogram() As Boolean
    ReadSM
    SMGraphsStatisticsHistogram = mSMGraphsStatisticsHistogram
End Property
Public Property Get SMGraphsStatisticsStair() As Boolean
    ReadSM
    SMGraphsStatisticsStair = mSMGraphsStatisticsStair
End Property
Public Property Get SMGraphsStatisticsSumCurve() As Boolean
    ReadSM
    SMGraphsStatisticsSumCurve = mSMGraphsStatisticsSumCurve
End Property
Public Property Get SMGraphsStatisticsBoxPlot() As Boolean
    ReadSM
    SMGraphsStatisticsBoxPlot = mSMGraphsStatisticsBoxPlot
End Property
Public Property Get SMProb() As Boolean
    ReadSM
    SMProb = mSMProb
End Property
Public Property Get SMProbRegression() As Boolean
    ReadSM
    SMProbRegression = mSMProbRegression
End Property
Public Property Get SMProbRegressionTable() As Boolean
    ReadSM
    SMProbRegressionTable = mSMProbRegressionTable
End Property
Public Property Get SMProbRegressionLin() As Boolean
    ReadSM
    SMProbRegressionLin = mSMProbRegressionLin
End Property
Public Property Get SMProbRegressionExp() As Boolean
    ReadSM
    SMProbRegressionExp = mSMProbRegressionExp
End Property
Public Property Get SMProbRegressionPow() As Boolean
    ReadSM
    SMProbRegressionPow = mSMProbRegressionPow
End Property
Public Property Get SMProbRegressionPol() As Boolean
    ReadSM
    SMProbRegressionPol = mSMProbRegressionPol
End Property
Public Property Get SMProbRegressionSin() As Boolean
    ReadSM
    SMProbRegressionSin = mSMProbRegressionSin
End Property
Public Property Get SMProbRegressionExcel() As Boolean
    ReadSM
    SMProbRegressionExcel = mSMProbRegressionExcel
End Property
Public Property Get SMProbRegressionUser() As Boolean
    ReadSM
    SMProbRegressionUser = mSMProbRegressionUser
End Property
Public Property Get SMProbDistributions() As Boolean
    ReadSM
    SMProbDistributions = mSMProbDistributions
End Property
Public Property Get SMProbDistributionsBinomial() As Boolean
    ReadSM
    SMProbDistributionsBinomial = mSMProbDistributionsBinomial
End Property
Public Property Get SMProbDistributionsNormal() As Boolean
    ReadSM
    SMProbDistributionsNormal = mSMProbDistributionsNormal
End Property
Public Property Get SMProbDistributionsChi() As Boolean
    ReadSM
    SMProbDistributionsChi = mSMProbDistributionsChi
End Property
Public Property Get SMProbDistributionsT() As Boolean
    ReadSM
    SMProbDistributionsT = mSMProbDistributionsT
End Property
Public Property Get SMProbTests() As Boolean
    ReadSM
    SMProbTests = mSMProbTests
End Property
Public Property Get SMProbTestsBinomial() As Boolean
    ReadSM
    SMProbTestsBinomial = mSMProbTestsBinomial
End Property
Public Property Get SMProbTestsNormal() As Boolean
    ReadSM
    SMProbTestsNormal = mSMProbTestsNormal
End Property
Public Property Get SMProbTestsChi() As Boolean
    ReadSM
    SMProbTestsChi = mSMProbTestsChi
End Property
Public Property Get SMProbTestsSimulation() As Boolean
    ReadSM
    SMProbTestsSimulation = mSMProbTestsSimulation
End Property
Public Property Get SMOther() As Boolean
    ReadSM
    SMOther = mSMOther
End Property
Public Property Get SMOtherNewEq() As Boolean
    ReadSM
    SMOtherNewEq = mSMOtherNewEq
End Property
Public Property Get SMOtherLatex() As Boolean
    ReadSM
    SMOtherLatex = mSMOtherLatex
End Property
Public Property Get SMOtherNewEqNum() As Boolean
    ReadSM
    SMOtherNewEqNum = mSMOtherNewEqNum
End Property
Public Property Get SMOtherTriangle() As Boolean
    ReadSM
    SMOtherTriangle = mSMOtherTriangle
End Property
Public Property Get SMOtherFigurs() As Boolean
    ReadSM
    SMOtherFigurs = mSMOtherFigurs
End Property
Public Property Get SMOtherTable() As Boolean
    ReadSM
    SMOtherTable = mSMOtherTable
End Property
'Public Property Get () As Boolean
'    ReadSM
'     = m
'End Property


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
#If Mac Then
#Else
    If Not (MaxProc Is Nothing) Then
        MaxProc.Exact = xval
    End If
#End If
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
        mgangetegn = VBA.ChrW$(183)
    ElseIf nygtegn = "x" Then
        SetRegSetting "Gangetegn", 1
        mgangetegn = VBA.ChrW$(215)
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
Public Property Get LastUpdateCheck() As String
    LastUpdateCheck = mLastUpdateCheck
End Property
Public Property Let LastUpdateCheck(ByVal text As String)
    SetRegSettingString "LastUpdateCheck", text
    mLastUpdateCheck = text
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
Public Property Let dAsDiffChrTemp(ByVal text As Boolean)
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
    mCASengine = CInt(GetRegSetting("CASengine"))
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
    If mInstallLocation <> vbNullString And mInstallLocation <> "0" Then
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
Public Property Get UseCodeFile() As Boolean
    UseCodeFile = mUseCodeFile
End Property
Public Property Let UseCodeFile(xval As Boolean)
    SetRegSetting "UseCodeFile", Abs(CInt(xval))
    mUseCodeFile = xval
End Property
Public Property Get UseCodeBlocks() As Boolean
    UseCodeBlocks = mUseCodeBlocks
End Property
Public Property Let UseCodeBlocks(xval As Boolean)
    SetRegSetting "UseCodeBlocks", Abs(CInt(xval))
    mUseCodeBlocks = xval
End Property
Public Property Get OutputColor() As Integer
    OutputColor = mOutputColor
End Property
Public Property Let OutputColor(ByVal xval As Integer)
    SetRegSetting "OutputColor", xval
    mOutputColor = xval
End Property


Public Property Get ReadSettingsFromFile() As Integer ' 0= not set, 1=dont read settings from file, 2=read from appdata, 3=read from program files, 4= first try appdata then program files, 5= first try program files, then appdata
    ReadSettingsFromFile = CInt(GetRegSetting("ReadSettingsFromFile"))
    If ReadSettingsFromFile <= 0 Then ' If a computer is run as a shared pc from intune. You can set this global key to look for the registry file
        ReadSettingsFromFile = val(RegKeyRead("HKEY_LOCAL_MACHINE\SOFTWARE\WORDMAT\Settings\ReadSettingsFromFile"))
    End If
End Property
Public Property Let ReadSettingsFromFile(xval As Integer)
    SetRegSetting "ReadSettingsFromFile", xval
End Property

'------------------- registry functions --------------------

Public Function GetRegSetting(key As String) As Integer
    Dim s As String
#If Mac Then
    s = RegKeyRead("HKEY_CURRENT_USER\SOFTWARE\WORDMAT\Settings\" & key)
#Else
    s = GetRegistryValue("HKCU", "SOFTWARE\WORDMAT\Settings", key)
#End If
    If s = vbNullString Then
        GetRegSetting = 0
    Else
        On Error Resume Next
        GetRegSetting = CInt(s)
    End If
End Function
Private Sub SetRegSetting(ByVal key As String, ByVal val As Integer)
#If Mac Then
    RegKeySave "HKEY_CURRENT_USER\SOFTWARE\WORDMAT\Settings\" & key, val, "REG_DWORD"
#Else
    SetRegistryValue "HKCU", "SOFTWARE\WORDMAT\Settings", key, REG_SZ, val
#End If
End Sub

#If VBA7 Then
Public Sub SetRegSettingLong(key As String, val As LongPtr)
#If Mac Then
    RegKeySave "HKEY_CURRENT_USER\SOFTWARE\WORDMAT\Settings\" & key, val, "REG_DWORD"
#Else
    SetRegistryValue "HKCU", "SOFTWARE\WORDMAT\Settings", key, REG_DWORD, val
#End If
End Sub
'Public Function GetRegSettingLong(key As String) As LongPtr
'#If Mac Then
'    GetRegSettingLong = CLngPtr(RegKeyRead("HKEY_CURRENT_USER\SOFTWARE\WORDMAT\Settings\" & key))
'#Else
'    GetRegSettingLong = CLngPtr(GetRegistryValue("HKCU", "SOFTWARE\WORDMAT\Settings", key, REG_DWORD))
'#End If
'End Function
'#Else
'Public Sub SetRegSettingLong(key As String, val As Long)
'    RegKeySave "HKEY_CURRENT_USER\SOFTWARE\WORDMAT\Settings\" & key, val, "REG_DWORD"
'End Sub
Public Function GetRegSettingLong(key As String) As Long
#If Mac Then
    GetRegSettingLong = CLng(RegKeyRead("HKEY_CURRENT_USER\SOFTWARE\WORDMAT\Settings\" & key))
#Else
    GetRegSettingLong = CLng(GetRegistryValue("HKCU", "SOFTWARE\WORDMAT\Settings", key, REG_DWORD))
#End If
End Function
#End If
Public Function GetRegSettingString(key As String) As String
#If Mac Then
    GetRegSettingString = RegKeyRead("HKEY_CURRENT_USER\SOFTWARE\WORDMAT\Settings\" & key)
#Else
    GetRegSettingString = GetRegistryValue("HKCU", "SOFTWARE\WORDMAT\Settings", key, REG_SZ)
#End If
End Function
Public Sub SetRegSettingString(key As String, ByVal val As String)
#If Mac Then
    RegKeySave "HKEY_CURRENT_USER\SOFTWARE\WORDMAT\Settings\" & key, val, "REG_SZ"
#Else
    SetRegistryValue "HKCU", "SOFTWARE\WORDMAT\Settings", key, REG_SZ, val
#End If
End Sub

Sub SaveSettingsToData()
    SaveSettingsToFile GetSettingsFilePath
End Sub
Function LoadSettingsFromData() As Boolean
    LoadSettingsFromData = LoadSettingsFromFile(GetSettingsFilePath, True)
End Function
Function LoadSettingsFromWMfolder() As Boolean
    LoadSettingsFromWMfolder = LoadSettingsFromFile(GetWordMatDir & "settings.txt", True)
End Function

Private Function GetSettingsFilePath() As String
#If Mac Then
    GetSettingsFilePath = DataFolder & "settings.txt"
#Else
    GetSettingsFilePath = Environ("AppData") & "\WordMat\"
    If Dir(GetSettingsFilePath, vbDirectory) = vbNullString Then
        MkDir GetSettingsFilePath
    End If
    GetSettingsFilePath = GetSettingsFilePath & "settings.txt"
#End If

End Function
Sub SaveSettingsToFile(Optional SettingsFileName As String)
    Dim s As String
    s = "# WordMat Settings" & vbCrLf
    
#If Mac Then
    SettingsFileName = GetDocumentsDir & "/settings.txt"
    s = s & "# Mac" & vbCrLf
#Else
    If SettingsFileName = vbNullString Then
        SettingsFileName = SaveAsFilePath(GetDocumentsDir & "\settings.txt", "Text files,*.txt")
    End If
    s = s & "# Win" & vbCrLf
#End If
    AddSetting s, "Version", AppVersion & PatchVersion
    
    AddSetting s, "Language", LanguageSetting
    AddSetting s, "GraphApp", GraphApp
    AddSetting s, "UseVBACAS", SettUseVBACAS
    AddSetting s, "CASengine", CASengine
    AddSetting s, "DllConnType", DllConnType
'    AddSetting s, "InstallLocation", InstallLocation
    AddSetting s, "DecOutType", MaximaDecOutType
    AddSetting s, "SigFig", MaximaCifre
    AddSetting s, "Exact", MaximaExact
    AddSetting s, "Radians", Radians
    AddSetting s, "Complex", MaximaComplex
    AddSetting s, "PolarOutput", PolarOutput
    AddSetting s, "AllTrig", AllTrig
    AddSetting s, "LogOutput", MaximaLogOutput
    AddSetting s, "BigFloat", MaximaBigFloat
    AddSetting s, "Units", MaximaUnits
    AddSetting s, "OutUnits", OutUnits
    AddSetting s, "Separator", MaximaSeparator
    AddSetting s, "Index", MaximaIndex
    AddSetting s, "ShowAssum", ShowAssum
    AddSetting s, "dAsDiffChr", dAsDiffChr
    AddSetting s, "LastUpdateCheck", LastUpdateCheck
    AddSetting s, "GangeTegn", GetRegSetting("Gangetegn")
    
    AddSetting s, "Forklaring", MaximaForklaring
    AddSetting s, "MaximaCommand", MaximaKommando
    AddSetting s, "EqNumPlacement", EqNumPlacement
    AddSetting s, "EqNumType", EqNumType
    AddSetting s, "EqAskRef", EqAskRef
    AddSetting s, "ExcelEmbed", ExcelIndlejret
    
    AddSetting s, "SettShortcutAltB", SettShortcutAltB
    AddSetting s, "SettShortcutAltD", SettShortcutAltD
    AddSetting s, "SettShortcutAltE", SettShortcutAltE
    AddSetting s, "SettShortcutAltF", SettShortcutAltF
    AddSetting s, "SettShortcutAltG", SettShortcutAltG
    AddSetting s, "SettShortcutAltGr", SettShortcutAltGr
    AddSetting s, "SettShortcutAltJ", SettShortcutAltJ
    AddSetting s, "SettShortcutAltL", SettShortcutAltL
    AddSetting s, "SettShortcutAltM2", SettShortcutAltM2
    AddSetting s, "SettShortcutAltN", SettShortcutAltN
    AddSetting s, "SettShortcutAltO", SettShortcutAltO
    AddSetting s, "SettShortcutAltP", SettShortcutAltP
    AddSetting s, "SettShortcutAltQ", SettShortcutAltQ
    AddSetting s, "SettShortcutAltR", SettShortcutAltR
    AddSetting s, "SettShortcutAltS", SettShortcutAltS
    AddSetting s, "SettShortcutAltT", SettShortcutAltT
    
    AddSetting s, "BackupType", BackupType
    AddSetting s, "BackupNo", BackupNo
    AddSetting s, "BackupMaxNo", BackupMaxNo
    AddSetting s, "BackupTime", BackupTime
    AddSetting s, "LatexStart", LatexStart
    AddSetting s, "LatexSlut", LatexSlut
    AddSetting s, "LatexUnits", LatexUnits
    AddSetting s, "ConvertTexWithMaxima", ConvertTexWithMaxima
    AddSetting s, "LatexSectionNumbering", LatexSectionNumbering
    AddSetting s, "LatexDocumentclass", LatexDocumentclass
    AddSetting s, "LatexFontsize", LatexFontsize
    AddSetting s, "LatexWordMargins", LatexWordMargins
    AddSetting s, "LatexTitlePage", LatexTitlePage
    AddSetting s, "LatexToc", LatexTOC
    AddSetting s, "ShowMenus", GetRegSettingString("ShowMenus")
    AddSetting s, "FormelFag", GetRegSettingString("FormelFag")
    AddSetting s, "FormelMatNiveau", GetRegSettingString("FormelMatNiveau")
    AddSetting s, "FormelFysNiveau", GetRegSettingString("FormelFysNiveau")
    AddSetting s, "FormelKemiNiveau", GetRegSettingString("FormelKemiNiveau")
    AddSetting s, "FormelUddannelse", GetRegSettingString("FormelUddannelse")
    AddSetting s, "FormelSamlingClose", GetRegSetting("FormelSamlingClose")
    AddSetting s, "FormelSamlingDouble", GetRegSetting("FormelSamlingDouble")
    AddSetting s, "FormelSamlingEnheder", GetRegSetting("FormelSamlingEnheder")
    AddSetting s, "FormelSamlingKonstanter", GetRegSetting("FormelSamlingKonstanter")
    AddSetting s, "UseCodeFile", UseCodeFile
    AddSetting s, "UseCodeBlocks", UseCodeBlocks
    
'    AddSetting s, "AntalBeregninger", Antalberegninger
'    AddSetting s, "",
    
    WriteTextfileToString SettingsFileName, s
#If Mac Then
    MsgBox2 "Settingsfile saved to " & vbCrLf & SettingsFileName, vbOKOnly, "Saved"
#End If
End Sub
Function LoadSettingsFromFile(filePath As String, Optional Silent As Boolean = False, Optional SaveToReg As Boolean = False) As Boolean
    Dim s As String, Arr() As String, Arr2() As String, i As Integer
    On Error GoTo fejl
    If filePath = vbNullString Then
#If Mac Then
        filePath = GetDocumentsDir & "/settings.txt"
#Else
        filePath = GetFilePath
#End If
    End If
    
    If Dir(filePath) = vbNullString Then
        If Not Silent Then
            MsgBox2 "Could not load settingsfile:" & vbCrLf & filePath, vbOKOnly, TT.Error
        End If
        Exit Function
    End If
    s = ReadTextfileToString(filePath)
    
    Arr = Split(s, vbCrLf)
    
    For i = 0 To UBound(Arr)
        If Left$(Trim$(Arr(i)), 1) <> "#" Then
            Arr2 = Split(Arr(i), "=")
            If UBound(Arr2) > 0 Then
                SetSetting Arr2(0), Arr2(1), SaveToReg
            End If
        End If
    Next
'    ReadAllSettingsFromRegistry
    LoadSettingsFromFile = True
    GoTo TheEnd
fejl:

TheEnd:
End Function
Private Sub AddSetting(ByRef s As String, Sett As String, SettVal As String)
    s = s & Sett & "=" & SettVal & vbCrLf
End Sub
Private Sub SetSetting(Sett As String, SettVal As String, Optional SaveToReg As Boolean = False)
'    SetRegSetting Sett, SettVal ' registry approach is slower 0,04s to load or set all in 2025
    
    If Sett = "Forklaring" Then
        mforklaring = CBool(SettVal)
        If SaveToReg Then MaximaForklaring = mforklaring
    ElseIf Sett = "MaximaCommand" Then
        mkommando = CBool(SettVal)
        If SaveToReg Then MaximaKommando = mkommando
    ElseIf Sett = "Exact" Then
        mExact = CInt(SettVal)
        If SaveToReg Then MaximaExact = mExact
    ElseIf Sett = "Radians" Then
        mradians = CBool(SettVal)
        If SaveToReg Then Radians = mradians
    ElseIf Sett = "SigFig" Then
        mcifre = CInt(SettVal)
        If SaveToReg Then MaximaCifre = mcifre
    ElseIf Sett = "Complex" Then
        mComplex = CBool(SettVal)
        If SaveToReg Then MaximaComplex = mComplex
    ElseIf Sett = "SolveBoolOrSet" Then
        mlmset = CBool(SettVal)
        If SaveToReg Then LmSet = mlmset
    ElseIf Sett = "Units" Then
        mUnits = CBool(SettVal)
        If SaveToReg Then MaximaUnits = mUnits
    ElseIf Sett = "LogOutput" Then
        mlogout = CBool(SettVal)
        If SaveToReg Then MaximaLogOutput = mlogout
    ElseIf Sett = "ExcelEmbed" Then
        mexcelembed = CBool(SettVal)
        If SaveToReg Then ExcelIndlejret = mexcelembed
    ElseIf Sett = "AllTrig" Then
        malltrig = CBool(SettVal)
        If SaveToReg Then AllTrig = malltrig
    ElseIf Sett = "OutUnits" Then
        moutunits = SettVal
        If SaveToReg Then OutUnits = moutunits
    ElseIf Sett = "BigFloat" Then
        mbigfloat = SettVal
        If SaveToReg Then MaximaBigFloat = mbigfloat
    ElseIf Sett = "AntalBeregninger" Then
        mantalb = CInt(SettVal)
        If SaveToReg Then
            Antalberegninger = mantalb
            AntalB = mantalb
        End If
    ElseIf Sett = "Index" Then
        mIndex = CBool(SettVal)
        If SaveToReg Then MaximaIndex = mIndex
    ElseIf Sett = "ShowAssum" Then
        mshowassum = CBool(SettVal)
        If SaveToReg Then ShowAssum = mshowassum
    ElseIf Sett = "PolarOutput" Then
        mpolaroutput = CBool(SettVal)
        If SaveToReg Then PolarOutput = mpolaroutput
    ElseIf Sett = "GraphApp" Then
        mgraphapp = CInt(SettVal)
#If Mac Then
        If mgraphapp = 0 Then mgraphapp = 2 ' gnuplot is not available on mac
#End If
        If SaveToReg Then GraphApp = mgraphapp
    ElseIf Sett = "Language" Then
        mlanguage = CInt(SettVal)
        If SaveToReg Then LanguageSetting = mlanguage
    ElseIf Sett = "dAsDiffChr" Then
        mdasdiffchr = CBool(SettVal)
        If SaveToReg Then dAsDiffChr = mdasdiffchr
    ElseIf Sett = "LatexStart" Then
        mlatexstart = SettVal
        If SaveToReg Then LatexStart = mlatexstart
    ElseIf Sett = "LatexSlut" Then
        mlatexslut = SettVal
        If SaveToReg Then LatexSlut = mlatexslut
    ElseIf Sett = "LatexUnits" Then
        mlatexunits = CBool(SettVal)
        If SaveToReg Then LatexUnits = mlatexunits
    ElseIf Sett = "ConvertTexWithMaxima" Then
        mConvertTexWithMaxima = CBool(SettVal)
        If SaveToReg Then ConvertTexWithMaxima = mConvertTexWithMaxima
    ElseIf Sett = "EqNumPlacement" Then
        meqnumplacement = CBool(SettVal)
        If SaveToReg Then EqNumPlacement = meqnumplacement
    ElseIf Sett = "EqNumType" Then
        meqnumtype = CBool(SettVal)
        If SaveToReg Then EqNumType = meqnumtype
    ElseIf Sett = "EqAskRef" Then
        maskref = CBool(SettVal)
        If SaveToReg Then EqAskRef = maskref
    ElseIf Sett = "BackupType" Then
        mBackupType = CInt(SettVal)
        If SaveToReg Then BackupType = mBackupType
    ElseIf Sett = "BackupNo" Then
        mbackupno = CInt(SettVal)
        If SaveToReg Then BackupNo = mbackupno
    ElseIf Sett = "BackupMaxNo" Then
        mbackupmaxno = CInt(SettVal)
        If SaveToReg Then BackupMaxNo = mbackupmaxno
    ElseIf Sett = "BackupTime" Then
        mbackuptime = CInt(SettVal)
        If SaveToReg Then BackupTime = mbackuptime
    ElseIf Sett = "LatexSectionNumbering" Then
        mLatexSectionNumbering = CBool(SettVal)
        If SaveToReg Then LatexSectionNumbering = mLatexSectionNumbering
    ElseIf Sett = "LatexDocumentclass" Then
        mLatexDocumentclass = CInt(SettVal)
        If SaveToReg Then LatexDocumentclass = mLatexDocumentclass
    ElseIf Sett = "LatexFontsize" Then
        mLatexFontsize = SettVal
        If SaveToReg Then LatexFontsize = mLatexFontsize
    ElseIf Sett = "LatexWordMargins" Then
        mLatexWordMargins = CBool(SettVal)
        If SaveToReg Then LatexWordMargins = mLatexWordMargins
    ElseIf Sett = "LatexTitlePage" Then
        mLatexTitlePage = CInt(SettVal)
        If SaveToReg Then LatexTitlePage = mLatexTitlePage
    ElseIf Sett = "LatexToc" Then
        mLatexTOC = CInt(SettVal)
        If SaveToReg Then LatexTOC = mLatexTOC
    ElseIf Sett = "CASengine" Then
        mCASengine = CInt(SettVal)
        If SaveToReg Then CASengine = mCASengine
    ElseIf Sett = "LastUpdateCheck" Then
        mLastUpdateCheck = SettVal
        If SaveToReg Then LastUpdateCheck = mLastUpdateCheck
    ElseIf Sett = "DllConnType" Then
        mDllConnType = CInt(SettVal)
        If SaveToReg Then DllConnType = mDllConnType
    ElseIf Sett = "InstallLocation" Then
        mInstallLocation = SettVal
        If SaveToReg Then InstallLocation = mInstallLocation
    ElseIf Sett = "UseVBACAS" Then
        SettUseVBACAS = SettVal ' mUseVBACAS
'        If SaveToReg Then SettUseVBACAS = msettusevbacas
    ElseIf Sett = "DecOutType" Then
        mDecOutType = CInt(SettVal)
        If SaveToReg Then MaximaDecOutType = mDecOutType
    ElseIf Sett = "FormelUddannelse" Then
        SetRegSettingString Sett, SettVal
    ElseIf Sett = "FormelFag" Then
        SetRegSettingString Sett, SettVal
    ElseIf Sett = "FormelMatNiveau" Then
        SetRegSettingString Sett, SettVal
    ElseIf Sett = "FormelFysNiveau" Then
        SetRegSettingString Sett, SettVal
    ElseIf Sett = "FormelKemiNiveau" Then
        SetRegSettingString Sett, SettVal
    ElseIf Sett = "FormelSamlingClose" Then
        SetRegSettingString Sett, SettVal
    ElseIf Sett = "FormelSamlingDouble" Then
        SetRegSettingString Sett, SettVal
    ElseIf Sett = "FormelSamlingEnheder" Then
        SetRegSettingString Sett, SettVal
    ElseIf Sett = "FormelSamlingKonstanter" Then
        SetRegSettingString Sett, SettVal
    ElseIf Sett = "ShowMenus" Then
        SetRegSettingString Sett, SettVal
'        RegKeySave "ShowMenus", SettVal
    ElseIf Sett = "UseCodeFile" Then
        mUseCodeFile = SettVal
        If SaveToReg Then UseCodeFile = mUseCodeFile
    ElseIf Sett = "UseCodeBlocks" Then
        mUseCodeBlocks = SettVal
        If SaveToReg Then UseCodeBlocks = mUseCodeBlocks
    End If

End Sub

Function GetFilePath(Optional Filter As String = "All Files,*.*") As String

    Dim fd As FileDialog
    Dim filterParts() As String
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    ' Set the dialog properties
    With fd
        .Title = "Select a Folder"
        .AllowMultiSelect = False ' Only allow selecting one folder
        
        ' Apply filter (format: "Description,*.ext")
        .Filters.Clear ' Clear existing filters
        filterParts = Split(Filter, ",")
        If UBound(filterParts) = 1 Then
            .Filters.Add filterParts(0), filterParts(1)
        End If
        
        If .Show = -1 Then ' User clicked OK
            GetFilePath = .SelectedItems(1)
        Else
            GetFilePath = ""
        End If
    End With
    
    Set fd = Nothing
End Function
Function SaveAsFilePath(DefaultFileName As String, Optional Filter As String = "All Files,*.*") As String

    Dim fd As FileDialog
    Dim filterParts() As String
    
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    
    ' Set the dialog properties
    With fd
        .Title = "Save As"
        .InitialFileName = DefaultFileName
        .AllowMultiSelect = False ' Only allow selecting one folder
        
        .FilterIndex = 13
        If .Show = -1 Then ' User clicked OK
            SaveAsFilePath = .SelectedItems(1)
        Else
            SaveAsFilePath = ""
        End If
    End With
    
    Set fd = Nothing
End Function

Sub TestReadSett()
Dim tid As Single

    tid = Timer
'    ReadAllSettingsFromRegistry
    LoadSettingsFromData
    MsgBox Timer - tid
End Sub
Public Function GetHardwareUUID() As String

#If Mac Then
    Dim scriptToRun As String
        
    On Error Resume Next
    scriptToRun = "do shell script ""system_profiler SPHardwareDataType | awk '/Hardware UUID/{print $3}'"""
    GetHardwareUUID = MacScript(scriptToRun)
    If Err.Number <> 0 Then
        GetHardwareUUID = "" ' Clear in case of error
    End If
    On Error GoTo 0
        
#Else
    On Error Resume Next
            
    GetHardwareUUID = RegKeyRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\IDConfigDB\Hardware Profiles\0001\HwProfileGuid")
    GetHardwareUUID = Trim$(GetHardwareUUID)
    If Left$(GetHardwareUUID, 1) = "{" Then GetHardwareUUID = Right$(GetHardwareUUID, Len(GetHardwareUUID) - 1)
    If Right$(GetHardwareUUID, 1) = "}" Then GetHardwareUUID = Left$(GetHardwareUUID, Len(GetHardwareUUID) - 1)
    
    On Error GoTo 0
        
#End If

End Function
