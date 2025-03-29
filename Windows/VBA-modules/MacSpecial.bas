Attribute VB_Name = "MacSpecial"
Option Explicit
#If Mac Then
'Private Declare PtrSafe Function bits2pict Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/libpicture64.dylib" () As LongPtr
'Private Declare PtrSafe Function uSleep Lib "/usr/lib/libc.dylib" Alias "usleep" (ByVal seconds As Long) As Long

Private m_screenwidth As Integer
Private mDatafolder As String

Function ScreenWidth() As Integer
    If m_screenwidth > 0 Then
        ScreenWidth = m_screenwidth
    Else
'        ScreenWidth = RunScript("getscreenwidth", "")
    End If
End Function

Function MacDrawDims(Optional x As Long = 0, Optional y As Long = 0) As String
Dim xdrawdim As Long, ydrawdim As Long
    If x > 0 Then
        xdrawdim = x
    End If
    If y > 0 Then
        ydrawdim = y
    End If
    
    Dim dx As Long
    Dim dy As Long
    
    dx = (254 * xdrawdim) / 72
    dy = (254 * ydrawdim) / 72
    
    MacDrawDims = "dimensions=[" & dx & "," & dy & "]"
End Function

Sub OpenExcelMac(fileName As String, Optional ParamS As String)
    RunScript "OpenExcel", GetWordMatDir & "Excelfiles/" & fileName & ";" & ParamS
End Sub

Sub requestFileAccess()

'Declare Variables_
    Dim fileAccessGranted As Boolean
    Dim filePermissionCandidates

'Create an array with file paths for which permissions are needed_
    filePermissionCandidates = Array("/Library/Application scripts/com.microsoft.Word/WordMatScripts.scpt")
'Request access from user_
    fileAccessGranted = GrantAccessToMultipleFiles(filePermissionCandidates)
'returns true if access granted, otherwise, false
End Sub

Function CheckAppleScriptTaskWordScriptFile(ScriptFileName As String) As Boolean
    'Function to Check if the AppleScriptTask script file exists
    'Ron de Bruin : 6-March-2016
    Dim AppleScriptTaskFolder As String
    Dim TestStr As String

    AppleScriptTaskFolder = MacScript("return POSIX path of (path to desktop folder) as string")
    AppleScriptTaskFolder = Replace(AppleScriptTaskFolder, "/Desktop", "") & _
        "Library/Application Scripts/com.microsoft.Word/"

    On Error Resume Next
    TestStr = Dir(AppleScriptTaskFolder & ScriptFileName, vbDirectory)
    On Error GoTo 0
    If TestStr = vbNullString Then
        CheckAppleScriptTaskWordScriptFile = False
    Else
        CheckAppleScriptTaskWordScriptFile = True
    End If
End Function

Function DataFolder() As String
    If mDatafolder <> "" Then
        DataFolder = mDatafolder
        Exit Function
   End If
    DataFolder = MacScript("return POSIX path of (path to desktop folder) as string")
    DataFolder = Replace(DataFolder, "/Desktop", "") & "Library/containers/com.microsoft.Word/Data/"
    mDatafolder = DataFolder
End Function

Function RunScript(ScriptName As String, Param As String) As String
' scriptfile must be placed in ~/Library/Application Scripts/com.microsoft.Word/
' ~/library is a hidden folder in the user folder
' filetype: .scpt or .applescript
On Error GoTo Fejl
    RunScript = AppleScriptTask("WordMatScripts.scpt", ScriptName, Param)
GoTo slut
Fejl:
    RunScript = "ScriptError"
slut:
End Function

Public Function ExecuteMaximaViaFile(MaximaCommand As String, Optional ByVal MaxWait As Integer = 10, Optional UnitCore As Boolean = False) As String
' M1 via textfiler
' scriptfile must be placed in ~/Library/Application Scripts/com.microsoft.Word/
' ~/library is a hidden folder in the user folder
' filetype: .scpt or .applescript
On Error GoTo Fejl

'    SaveCommandFile MaximaCommand
    If UnitCore Then
'        AppleScriptTask "WordMatScripts.scpt", "RunMaximaUnit", CStr(MaxWait)
'        If OutUnits <> "" Then ' removed as units are now handled by RunMaximaFile via GetSettingString
'            ExecuteMaximaViaFile = AppleScriptTask("WordMatScripts.scpt", "RunMaximaUnit", CStr(MaxWait) & "£" & "setunits(" & omax.ConvertUnits(OutUnits) & ")$" & MaximaCommand)
'        Else
            ExecuteMaximaViaFile = AppleScriptTask("WordMatScripts.scpt", "RunMaximaUnit", CStr(MaxWait) & "£" & MaximaCommand)
'        End If
    Else
            ExecuteMaximaViaFile = AppleScriptTask("WordMatScripts.scpt", "RunMaxima", CStr(MaxWait) & "£" & MaximaCommand)
    End If
'    ExecuteMaximaViaFile = ReadMaximaOutputFile()
'MsgBox ExecuteMaximaViaFile
    
    GoTo slut
Fejl:
    ExecuteMaximaViaFile = "Fejln" & Err.Number
slut:
    
End Function

#Else
#End If
