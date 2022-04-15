Attribute VB_Name = "MacSpecial"
Option Explicit
#If Mac Then
Private Declare PtrSafe Function bits2pict Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/libpicture64.dylib" () As LongPtr
Private Declare PtrSafe Function uSleep Lib "/usr/lib/libc.dylib" Alias "usleep" (ByVal seconds As Long) As Long

Private m_screenwidth As Integer
Private mDatafolder As String


Sub TestMaxima()

    PrepareMaxima
    MaxProc.ExecuteMaximaCommand "2+4;", 5
'    MaxProc.WaitForMaximaUntil 5
    MsgBox MaxProc.LastMaximaOutput
    Application.Windows(1).WindowState = 1
    'ActiveDocument.ActiveWindow.View = wdReadingView
    
    
End Sub


Function ScreenWidth() As Integer
    If m_screenwidth > 0 Then
        ScreenWidth = m_screenwidth
    Else
'        ScreenWidth = RunScript("getscreenwidth", "")
    End If
End Function
'Mac: LoadPicture
'
'Function LoadPictureMac(pathPict) As IPictureDisp
'    On Error GoTo Err
'    Dim pathBmp As String
'    Const sIID_IPicture As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}" ' der var fejl om manglende def af denne. Er fundet p*aa* nettet s*aa* ikke sikker p*aa* om den er korrekt
'
'    If (LCase(Right(pathPict, 4)) = ".bmp") Then
'        pathBmp = pathPict
'    Else
'        pathBmp = GetTempDir() & "WordMatGraf.bmp"
'
'        On Error Resume Next
'        Kill pathBmp
'        On Error GoTo Err
'
'        Application.ScreenUpdating = False
'        Dim ishp As InlineShape
'        Set ishp = Application.ActiveDocument.InlineShapes.AddPicture(pathPict)
'        ishp.ScaleHeight = 100
'        ishp.ScaleWidth = 100
'        ishp.SaveAsPicture msoPictureTypeBMP, pathBmp
'        ishp.Delete
'        Application.ScreenUpdating = True
'    End If
'
'    Exit Function
'
'    Dim num As Integer
'    num = FreeFile
'    Open pathBmp For Binary Access Read As num
'    Dim siz As Long
'    siz = LOF(num)
'    Dim buf() As Byte
'    ReDim buf(0 To siz - 1)
'    Get #num, 1, buf
'    Close #num
'
'    Dim hr As LongPtr
'    Dim ipic As IPicture
'    hr = bits2pict(buf(0), StrPtr(sIID_IPicture), ipic)
'    Set LoadPictureMac = ipic
'    Exit Function
'
'Err:
'    MsgBox Err.Description, , "LoadPicture"
'End Function


Function MacDrawDims(Optional x As Long = 0, Optional Y As Long = 0) As String
Dim xdrawdim As Long, ydrawdim As Long
    If x > 0 Then
        xdrawdim = x
    End If
    If Y > 0 Then
        ydrawdim = Y
    End If
    
    Dim dx As Long
    Dim dy As Long
    
    dx = (254 * xdrawdim) / 72
    dy = (254 * ydrawdim) / 72
    
    MacDrawDims = "dimensions=[" & dx & "," & dy & "]"
End Function


Sub OpenExcelMac(FileName As String)
    RunScript "OpenExcel", GetWordMatDir & "Excelfiles/" & FileName
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


Sub TestCreateSCPTfile()
    Dim FileName As String
    Dim ScriptString As String
    
    'Name of the file you want to create
    FileName = "FileCheckSCPT.scpt"

    'Script that you want in this file
    ScriptString = "on ExistsFile(filePath)" & ChrW(13)
    ScriptString = ScriptString & "tell application ""System Events"" to return (exists disk item filePath) and class of disk item filePath = file " & ChrW(13)
    ScriptString = ScriptString & "end ExistsFile"
    
    CreateSCPTfile FileName, ScriptString
    

End Sub
Sub CreateSCPTfile(FileName As String, ScriptString As String)
'Code example to create or update scpt files in the
'/Library/Application Scripts/com.microsoft.Word folder
'location for the files used by the AppleScriptTask function
'The MakeSCPTFile.scpt file must be in this location
    Dim AppleScriptTaskFolder As String
    Dim AppleScriptTaskScript As String
    Dim RunMyScript

    '***** Do not change the code below *****
    If CheckAppleScriptTaskWordScriptFile(ScriptFileName:="MakeSCPTFile.scpt") = False Then
        MsgBox "Sorry the MakeSCPTFile.scpt is not in the correct location"
        Exit Sub
    End If

    AppleScriptTaskFolder = MacScript("return POSIX path of (path to desktop folder) as string")
    AppleScriptTaskFolder = Replace(AppleScriptTaskFolder, "/Desktop", "") & _
        "Library/Application Scripts/com.microsoft.Word/"
    AppleScriptTaskFolder = AppleScriptTaskFolder & FileName

    'Call the AppleScriptTask function
    AppleScriptTaskScript = ScriptString & ";" & AppleScriptTaskFolder
    RunMyScript = AppleScriptTask("MakeSCPTFile.scpt", "CreateSCPTFile", AppleScriptTaskScript)
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
Public Function ExecuteMaximaViaFile(MaximaCommand As String, Optional ByVal MaxWait As Integer = 10, Optional UnitCore As Boolean = False) As String
' M1 via textfiler
' scriptfile must be placed in ~/Library/Application Scripts/com.microsoft.Word/
' ~/library is a hidden folder in the user folder
' filetype: .scpt or .applescript
On Error GoTo fejl
'    SaveCommandFile MaximaCommand
    If UnitCore Then
'        AppleScriptTask "WordMatScripts.scpt", "RunMaximaUnit", CStr(MaxWait)
        If OutUnits <> "" Then
            ExecuteMaximaViaFile = AppleScriptTask("WordMatScripts.scpt", "RunMaximaUnit", CStr(MaxWait) & "£" & "setunits(" & omax.ConvertUnits(OutUnits) & ")$" & MaximaCommand)
        Else
            ExecuteMaximaViaFile = AppleScriptTask("WordMatScripts.scpt", "RunMaximaUnit", CStr(MaxWait) & "£" & MaximaCommand)
        End If
    Else
        ExecuteMaximaViaFile = AppleScriptTask("WordMatScripts.scpt", "RunMaxima", CStr(MaxWait) & "£" & MaximaCommand)
    End If
'    ExecuteMaximaViaFile = ReadMaximaOutputFile()
'MsgBox ExecuteMaximaViaFile
    GoTo slut
fejl:
    ExecuteMaximaViaFile = "Fejln" & Err.Number
slut:
    
End Function
Function RunScript(ScriptName As String, Param As String) As String
' scriptfile must be placed in ~/Library/Application Scripts/com.microsoft.Word/
' ~/library is a hidden folder in the user folder
' filetype: .scpt or .applescript
On Error GoTo fejl
    RunScript = AppleScriptTask("WordMatScripts.scpt", ScriptName, Param)
GoTo slut
fejl:
    RunScript = "ScriptError"
slut:
End Function
#Else
Function RunScript(ScriptName As String, Param As String) As String
' lige nu en dummy shell s*aa* der ikke kommer compilefejl
'    RunScript = Shell("WordMatScripts.scpt", ScriptName, Param)
End Function
#End If




