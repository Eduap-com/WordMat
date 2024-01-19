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
'    Const sIID_IPicture As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}" ' der var fejl om manglende def af denne. Er fundet på nettet så ikke sikker på om den er korrekt
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


Function MacDrawDims(Optional X As Long = 0, Optional Y As Long = 0) As String
Dim xdrawdim As Long, ydrawdim As Long
    If X > 0 Then
        xdrawdim = X
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


Sub OpenExcelMac(fileName As String)
    RunScript "OpenExcel", GetWordMatDir & "Excelfiles/" & fileName
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
    Dim fileName As String
    Dim ScriptString As String
    
    'Name of the file you want to create
    fileName = "FileCheckSCPT.scpt"

    'Script that you want in this file
    ScriptString = "on ExistsFile(filePath)" & ChrW(13)
    ScriptString = ScriptString & "tell application ""System Events"" to return (exists disk item filePath) and class of disk item filePath = file " & ChrW(13)
    ScriptString = ScriptString & "end ExistsFile"
    
    CreateSCPTfile fileName, ScriptString
    

End Sub
Sub CreateSCPTfile(fileName As String, ScriptString As String)
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
    AppleScriptTaskFolder = AppleScriptTaskFolder & fileName

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

Function RunScript(ScriptName As String, Param As String) As String
' scriptfile must be placed in ~/Library/Application Scripts/com.microsoft.Word/
' ~/library is a hidden folder in the user folder
' filetype: .scpt or .applescript
On Error GoTo Fejl
    RunScript = AppleScriptTask("WordMatScripts.scpt", ScriptName, Param)
GoTo Slut
Fejl:
    RunScript = "ScriptError"
Slut:
End Function
#Else
Function RunScript(ScriptName As String, Param As String) As String
' lige nu en dummy shell så der ikke kommer compilefejl
'    RunScript = Shell("WordMatScripts.scpt", ScriptName, Param)
End Function

#End If

Public Function ExecuteMaximaViaFile(MaximaCommand As String, Optional ByVal MaxWait As Integer = 10, Optional UnitCore As Boolean = False) As String
' M1 via textfiler
' scriptfile must be placed in ~/Library/Application Scripts/com.microsoft.Word/
' ~/library is a hidden folder in the user folder
' filetype: .scpt or .applescript
On Error GoTo Fejl
#If Mac Then
#Else
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
#End If

'    SaveCommandFile MaximaCommand
    If UnitCore Then
'        AppleScriptTask "WordMatScripts.scpt", "RunMaximaUnit", CStr(MaxWait)
        If OutUnits <> "" Then
#If Mac Then
            ExecuteMaximaViaFile = AppleScriptTask("WordMatScripts.scpt", "RunMaximaUnit", CStr(MaxWait) & "£" & "setunits(" & omax.ConvertUnits(OutUnits) & ")$" & MaximaCommand)
#Else
#End If
        Else
#If Mac Then
            ExecuteMaximaViaFile = AppleScriptTask("WordMatScripts.scpt", "RunMaximaUnit", CStr(MaxWait) & "£" & MaximaCommand)
#Else
#End If
        End If
    Else
#If Mac Then
        ExecuteMaximaViaFile = AppleScriptTask("WordMatScripts.scpt", "RunMaxima", CStr(MaxWait) & "£" & MaximaCommand)
#Else
        ExecuteMaximaViaFile = WshShell.Exec("maxima.bat --batch-string ""2+3;""").StdOut.ReadAll
#End If
    End If
'    ExecuteMaximaViaFile = ReadMaximaOutputFile()
'MsgBox ExecuteMaximaViaFile
    

    GoTo Slut
Fejl:
    ExecuteMaximaViaFile = "Fejln" & Err.Number
Slut:
    
#If Mac Then
#Else
    Set WshShell = Nothing
#End If
End Function


Sub TestSHell()
' WshShell.Exec kan ike skjule vinduet. .run kan ikke få output tilbage
' .exec() kører asynkront og via det object den returnerer kan man checke om den færdig. WshScriptExec.status=1
' WshScriptExec.StdIn.write "2+3;"  kan bruges til at sende input
' WshScriptExec.StdIn.ReadAll()  eller readline eller read(1) henter output. Desværre er alle blocking, så låser hvis der ikke er output
' WshScriptExec.terminate kan bruges til at forcelukke
' AtEndOfStream er også blocking

    Dim WshShell As Object, Output As String, WshScriptExec As Object, i As Integer, t As String
    Set WshShell = CreateObject("WScript.Shell")
    
    
'    Set WshScriptExec = WshShell.Exec("""C:\Program Files (x86)\WordMat\Maxima-5.47.0\bin\maxima.bat"" --batch-string ""2+3;""")
    Set WshScriptExec = WshShell.Exec("""C:\Program Files (x86)\WordMat\Maxima-5.47.0\bin\maxima.bat""")
'    Set WshScriptExec = WshShell.Exec("cmd /c start /B cmd /c ""C:\Program Files (x86)\WordMat\Maxima-5.47.0\bin\maxima.bat"" & timeout /t 5 & taskkill /IM cmd.exe")
    
    'start /B cmd /c "C:\Program Files (x86)\WordMat\Maxima-5.47.0\bin\maxima.bat" & timeout /t 5 & taskkill /IM cmd.exe
    
    WshScriptExec.StdIn.Write "2+3;" & vbCrLf
    
    Do While WshScriptExec.Status = 0 And i < 50 ' 0=running  1=finished
        Sleep (100)
        DoEvents
        'start cmd /c "C:\Program Files (x86)\WordMat\Maxima-5.47.0\bin\maxima.bat" & timeout /t 5 & taskkill /im Maxima*
        Do
            t = WshScriptExec.StdOut.Read(1)
'            t = WshScriptExec.StdOut.ReadLine
            Output = Output & t '& vbCrLf
            Debug.Assert t <> "2"
        Loop Until t = vbNullString Or WshScriptExec.StdOut.AtEndOfStream
        
        'waitform code here
        i = i + 1
    Loop
        
    If WshScriptExec.Status = 1 Then
        Output = Output & WshScriptExec.StdOut.ReadAll
    Else
        Do While Not WshScriptExec.StdOut.AtEndOfStream
            Output = Output & WshScriptExec.StdOut.Read(1)
        Loop
    End If
    MsgBox Output
    
    If WshScriptExec.Status = 1 Then ' Hvis den stadig kører, så tving den til at lukke
        WshScriptExec.Terminate
    End If
    Set WshShell = Nothing
End Sub
